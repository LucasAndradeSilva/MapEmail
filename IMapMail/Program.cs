using IMapMail.Model;
using IMapMail.Service;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Pop3;
using MailKit.Search;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using MimeKit;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace IMapMail
{
    class Program
    {
        public static IConfigurationRoot _Configuration;
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
               .SetBasePath(Directory.GetCurrentDirectory())
               .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)              
               .AddEnvironmentVariables();

            IConfigurationRoot configuration = builder.Build();
            var mySettingsConfig = new MySettingsConfig();
            configuration.GetSection("MySettings").Bind(mySettingsConfig);

            _Configuration = configuration;

            Console.WriteLine("==============================");
            Console.WriteLine("===== IMAP MAIL INICIADO =====");
            Console.WriteLine("==============================");

            Parametros parametros = new Parametros();
            parametros.Email = "";
            parametros.Senha = "";
            if (args.Length == 0)
            {
                Console.WriteLine("");
                Console.WriteLine("Informe seu email: ");
                parametros.Email = Console.ReadLine();
                Console.WriteLine("");
                Console.WriteLine("Informe sua senha: ");
                parametros.Senha = Console.ReadLine();
                Console.WriteLine("");
                Console.WriteLine("Deseja Deletar os Emails? (S/N)");
                var result = Console.ReadLine();
                if (result == "S" || result == "s") parametros.ApagaEmails = true;
                else parametros.ApagaEmails = Convert.ToBoolean(_Configuration.GetSection("Pagar_Emails").Value);
                Console.WriteLine("");
                Console.WriteLine("Informe o Caminho que deseja salver os Anexos: ");
                parametros.Caminho = Console.ReadLine();
                Console.WriteLine("");
                Console.WriteLine("Digite a pasta que você deseja baixar: ");
                parametros.Pasta = Console.ReadLine();
            }
            else
            {
                parametros.Email = string.IsNullOrEmpty(args[0]) ? "" : args[0];
                parametros.Senha = string.IsNullOrEmpty(args[1]) ? "" : args[1];
                parametros.ApagaEmails = Convert.ToBoolean(args[2]);
                parametros.Caminho = string.IsNullOrEmpty(args[3]) ? "" : args[3];
                parametros.Pasta = string.IsNullOrEmpty(args[4]) ? "" : args[4];
            }

            Console.Clear();

            //Get Emails
            var ListEmails = GetMails(parametros);

            if(ListEmails != null && ListEmails.Count > 0)
            {
                Console.Clear();
                Console.WriteLine("\n Salvando Dados no Banco...");
                Thread.Sleep(3000);

                //Salvando Emails
                var services = new DBService();
                var Emails = JsonConvert.SerializeObject(ListEmails);
                services.GravaEmails(Emails,_Configuration);
                Console.Clear();
            }

            Console.Clear();
            Console.WriteLine("===============================");
            Console.WriteLine("===== IMAP MAIL FINALIZADO ====");
            Console.WriteLine("===============================");
            Console.ReadKey();
        }

        public static List<object> GetMails(Parametros parametros)
        {
            try
            {
                #region IMAP
                using (ImapClient client = new ImapClient())
                {

                    // Connect to the server                    
                    Console.WriteLine("\n Conectando Aguarde...");
                    client.Connect("mail.visaogrupo.com.br", 143, false); //Visao                    
                    client.Authenticate(parametros.Email.Replace(" ", ""), parametros.Senha.Replace(" ", ""));
                    Console.Clear();
                    if (!client.IsConnected)
                    {
                        Console.WriteLine("Erro na Conexão!");
                        return null;
                    }

                    if (!client.IsAuthenticated)
                    {
                        Console.WriteLine("Erro na Autenticação!");
                        return null;
                    }

                    Console.WriteLine("\n Email Conectado com Sucess!");
                    Thread.Sleep(3000);

                    List<object> allMessages = new List<object>();
                    int messageCount = 0;
                                     
                    if (string.IsNullOrEmpty(parametros.Pasta))
                    {
                        client.Inbox.Open(FolderAccess.ReadWrite);
                        messageCount = client.Inbox.Count;
                        if (messageCount > 0)
                        {
                            var mails = client.Inbox.Search(SearchQuery.All);
                            var Baixados = BaixaEmails(client, mails, messageCount, client.Inbox.FullName, client.Inbox, parametros);
                            if (messageCount > 0) allMessages.AddRange(Baixados);
                        }
                        client.Inbox.Close();
                    }
                    else
                    {
                        var Folders = client.GetFolder(parametros.Pasta);
                        Folders?.Open(FolderAccess.ReadWrite);
                        if (Folders != null)
                        {
                            messageCount += Folders != null ? Folders.Count : messageCount;
                            var Baixados = BaixaEmails(client, Folders.Search(SearchQuery.All), messageCount, Folders.FullName , Folders, parametros);
                            if (messageCount > 0) allMessages.AddRange(Baixados);
                            Folders.Close();
                        }
                    }
                                   
                    Console.Clear();
                    if (messageCount > 0) Console.WriteLine("\n Foram Baixados " + messageCount.ToString() + " email(s)");
                    else Console.WriteLine("\n Nenhuma Email Encontrado!");
                    Thread.Sleep(2000);
                                                                                                               
                    return allMessages;                                       
                }
                #endregion
            }
            catch (Exception exe)
            {
                Console.WriteLine("\nErro: "+exe.Message);
                Console.ReadKey();
                return null;
            }

        }

        public static List<object> BaixaEmails(ImapClient client, IList<UniqueId> mails, int count, string pasta, IMailFolder folder, Parametros parametros)
        {           
            if(count > 0)
            {
                Console.Clear();
                Console.WriteLine(" Fazendo Downloads da(o) " + pasta + "...\n");
                Thread.Sleep(3000);

                List<object> allMessages = new List<object>();
                IList<UniqueId> UniMails = new List<UniqueId>();
                double i = 1.0;
                double ii = 1.0;
                double porcentagem = 0;
                var qtdGrava = 0;
                foreach (var mail 
                    in mails)               
                {
                    porcentagem = (i/count)*100;
                    if (ii == 50)
                    {
                        ii = 1;
                        Console.Clear();
                        Console.WriteLine(" Fazendo Downloads da(o) " + pasta + "...\n");
                        Console.SetCursorPosition(1, 1);
                        Console.Write(i + "/" + count+" - "+ Convert.ToInt32(porcentagem) +"/100%");
                        Console.SetCursorPosition(Convert.ToInt32(ii), 2);
                        Console.Write("█");                        
                    }
                    else
                    {
                        Console.SetCursorPosition(1, 1);
                        Console.Write(i + "/" + count + " - " + Convert.ToInt32(porcentagem) + "/100%");
                        Console.SetCursorPosition(Convert.ToInt32(ii), 2);
                        Console.Write("█");                       
                    }
                    
                    var message = folder.GetMessage(mail);                   

                    Emails emails = new Emails();
                    emails.IdEmail = message.MessageId;
                    emails.Titulo = message.Subject;
                    emails.DtHrEnvio = message.Date.Date.ToString("dd/MM/yyyy");
                    emails.De = message.From[0].ToString().Contains("<") ? message.From[0].ToString()?.Split('<')?[1]?.Replace(">", "") : message.From[0].ToString();
                    emails.Para = message.To.ToString().Contains("<") ? message.To.ToString()?.Split('<')?[1]?.Replace(">", "") : message.To.ToString();
                    emails.Html = message.HtmlBody;
                    emails.Body = message.TextBody;                    
                    foreach (var Cc in message.Cc)
                    {
                        var Address = Cc.ToString().Contains("<") ? Cc.ToString()?.Split('<')?[1]?.Replace(">", "") : Cc.ToString();

                        emails.CC += string.IsNullOrEmpty(emails.CC) ? Address : "; " + Address;
                    }                    

                    foreach (var attachment in message.Attachments)
                    {
                        if (attachment is MessagePart)
                        {
                            var fileName = attachment.ContentDisposition?.FileName;
                            var rfc822 = (MessagePart)attachment;

                            if (string.IsNullOrEmpty(fileName))
                                fileName = "attached-message.eml";

                            fileName = fileName.Replace("/", "_").Replace(":", "").Replace("?", "").Replace("*", "").Replace(",", "").Replace("&", "").Replace("!", "").Replace("%", "").Replace("\"","").Replace("|", "").Replace("<", "").Replace(">", "");  
                            var caminho = string.IsNullOrEmpty(parametros.Caminho) ? _Configuration.GetSection("Folder").Value : parametros.Caminho;
                            caminho = caminho.EndsWith("//") ? caminho : caminho+"//";
                            caminho += (mail.Id + "_" + fileName).Replace(" ", "");
                            emails.CaminhoAnexos += emails.CaminhoAnexos.Contains(";") ? "; " + caminho : caminho;
                            using (var stream = File.Create(caminho))
                                rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;
                            var fileName = part.FileName.Replace("/", "_").Replace(":", "").Replace("?", "").Replace("*", "").Replace(",", "").Replace("&", "").Replace("!", "").Replace("%", "").Replace("\"", "").Replace("|", "").Replace("<", "").Replace(">", "");
                            var caminho = string.IsNullOrEmpty(parametros.Caminho) ? _Configuration.GetSection("Folder").Value : parametros.Caminho;
                            caminho = caminho.EndsWith("//") ? caminho : caminho + "//";
                            caminho += (mail.Id + "_" + fileName).Replace(" ", "");
                            emails.CaminhoAnexos = "";
                            emails.CaminhoAnexos += emails.CaminhoAnexos.Contains(";") ? "; " + caminho : caminho;
                            using (var stream = File.Create(caminho))
                                part.Content.DecodeTo(stream);
                        }
                    }

                    allMessages.Add(emails);
                    UniMails.Add(mail);
                    i++;
                    ii++;
                    qtdGrava++;
                    if(qtdGrava == 10)
                    {
                        qtdGrava = 0;                       
                        GravaBanco(allMessages);
                        if(parametros.ApagaEmails) DeleteEmails(client, UniMails, folder);
                        allMessages.Clear();
                        UniMails.Clear();
                    }
                }

                if (parametros.ApagaEmails) DeleteEmails(client, mails, folder);
                Console.Clear();
                return allMessages;
            }

            return null;                      
        }

        public static void GravaBanco(List<object> ListEmails) 
        {                      
            //Salvando Emails
            var services = new DBService();
            var Emails = JsonConvert.SerializeObject(ListEmails);
            services.GravaEmails(Emails, _Configuration);                  
        }

        public static void DeleteEmails(ImapClient client, IList<UniqueId> emails, IMailFolder folder)
        {
            folder.AddFlags(emails, MessageFlags.Deleted, false);
            folder.Expunge();
        }
    }
}
