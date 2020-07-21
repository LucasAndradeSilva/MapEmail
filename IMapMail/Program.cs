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

            var email = "";
            var senha = "";
            if (args.Length == 0)
            {
                Console.WriteLine("");
                Console.WriteLine("Informe seu email: ");
                email = Console.ReadLine();
                Console.WriteLine("");
                Console.WriteLine("Informe sua senha: ");
                senha = Console.ReadLine();
            }
            else
            {
                email = args[0];
                senha = args[1];
            }

            Console.Clear();

            //Get Emails
            var ListEmails = GetMails(email, senha);

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

        public static List<object> GetMails(string email, string senha)
        {
            try
            {
                #region IMAP
                using (ImapClient client = new ImapClient())
                {

                    // Connect to the server                    
                    Console.WriteLine("\n Conectando Aguarde...");
                    client.Connect("mail.visaogrupo.com.br", 143, false); //Visao                    
                    client.Authenticate(email.Replace(" ", ""), senha);
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

                    client.Inbox.Open(FolderAccess.ReadWrite);     
                    int messageCount = client.Inbox.Count;
                    if (messageCount > 0)
                    {
                        var mails = client.Inbox.Search(SearchQuery.All);
                        var Baixados = BaixaEmails(client, mails, messageCount,"Caixa de Entrada",client.Inbox);
                        if (messageCount > 0) allMessages.AddRange(Baixados);
                    }                    
                    client.Inbox.Close();

                    var Folders = client.GetFolder(SpecialFolder.Archive);
                    Folders?.Open(FolderAccess.ReadWrite);
                    if(Folders != null)
                    {
                        messageCount += Folders != null ? Folders.Count : messageCount;
                        var Baixados = BaixaEmails(client, Folders.Search(SearchQuery.All), messageCount, "Arquivados", Folders);
                        if (messageCount > 0) allMessages.AddRange(Baixados);
                        Folders.Close();
                    }

                    Folders = client.GetFolder(SpecialFolder.Drafts);
                    Folders?.Open(FolderAccess.ReadWrite);
                    if (Folders != null)
                    {
                        messageCount += Folders != null ? Folders.Count : messageCount;
                        var Baixados = BaixaEmails(client, Folders.Search(SearchQuery.All), messageCount, "Lixeira", Folders);
                        if (messageCount > 0) allMessages.AddRange(Baixados);
                        Folders.Close();
                    }

                    Folders = client.GetFolder(SpecialFolder.Important);
                    Folders?.Open(FolderAccess.ReadWrite);
                    if (Folders != null)
                    {
                        messageCount += Folders != null ? Folders.Count : messageCount;
                        var Baixados = BaixaEmails(client, Folders.Search(SearchQuery.All), messageCount,"Importante", Folders);
                        if (messageCount > 0) allMessages.AddRange(Baixados);
                        Folders.Close();
                    }

                    Folders = client.GetFolder(SpecialFolder.Sent);
                    Folders?.Open(FolderAccess.ReadWrite);
                    if (Folders != null)
                    {
                        messageCount += Folders != null ? Folders.Count : messageCount;
                        var Baixados = BaixaEmails(client, Folders.Search(SearchQuery.All), messageCount, "Enviados", Folders);
                        if(messageCount > 0) allMessages.AddRange(Baixados);
                        Folders.Close();
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

        public static void DeleteEmails(ImapClient client, IList<UniqueId> emails, IMailFolder folder)
        {
            folder.AddFlags(emails, MessageFlags.Deleted, false);
            folder.Expunge();         
        }

        public static List<object> BaixaEmails(ImapClient client, IList<UniqueId> mails, int count, string pasta, IMailFolder folder)
        {           
            if(count > 0)
            {
                Console.Clear();
                Console.WriteLine(" Fazendo Downloads da(o) " + pasta + "...\n");
                Thread.Sleep(3000);

                List<object> allMessages = new List<object>();

                var i = 1;
                var ii = 1;
                double porcentagem = 0;
                var qtdGrava = 0;
                foreach (var mail in mails)               
                {
                    porcentagem = (i / count);//(double)(i / count)*100;
                    if (ii == 50)
                    {
                        ii = 1;
                        Console.Clear();
                        Console.WriteLine(" Fazendo Downloads da(o) " + pasta + "...\n");
                        Console.SetCursorPosition(ii, 1);
                        Console.Write("█");
                        Console.Write(porcentagem + "/100");
                    }
                    else
                    {
                        Console.SetCursorPosition(ii, 1);
                        Console.Write("█");
                        Console.Write(porcentagem + "/100%");
                    }
                    
                    var message = folder.GetMessage(mail);                   

                    Emails emails = new Emails();
                    emails.IdEmail = message.MessageId;
                    emails.Titulo = message.Subject;
                    emails.Data = message.Date.ToString();
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

                            fileName = fileName.Replace("/", "_");
                            var caminho = _Configuration.GetSection("Folder").Value;
                            caminho += (mail.Id + "_" + fileName).Replace(" ", "");
                            emails.CaminhoAnexos += emails.CaminhoAnexos.Contains(";") ? "; " + caminho : caminho;
                            using (var stream = File.Create(caminho))
                                rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;
                            var fileName = part.FileName.Replace("/", "_");
                            var caminho = _Configuration.GetSection("Folder").Value;
                            caminho += (mail.Id + "_" + fileName).Replace(" ", "");
                            emails.CaminhoAnexos = "";
                            emails.CaminhoAnexos += emails.CaminhoAnexos.Contains(";") ? "; " + caminho : caminho;
                            using (var stream = File.Create(caminho))
                                part.Content.DecodeTo(stream);
                        }
                    }

                    allMessages.Add(emails);
                    i++;
                    ii++;
                    qtdGrava++;
                    if(qtdGrava == 10)
                    {
                        GravaBanco(allMessages);
                        allMessages.Clear();
                    }
                }

                //DeleteEmails(client, mails, folder);
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
            Console.Clear();            
        } 
    }
}
