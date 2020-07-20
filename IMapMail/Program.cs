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
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace IMapMail
{
    class Program
    {
        public static IConfiguration _Configuration;
        static void Main(string[] args)
        {
           
            Console.WriteLine("==============================");
            Console.WriteLine("===== IMAP MAIL INICIADO =====");
            Console.WriteLine("==============================");
            
            Console.WriteLine("");
            Console.WriteLine("Informe seu email: ");
            var email = Console.ReadLine();
            Console.WriteLine("");
            Console.WriteLine("Informe sua senha: ");
            var senha = Console.ReadLine();

            Console.Clear();

            //Get Emails
            var ListEmails = GetMails(email, senha);

            if(ListEmails != null)
            {
                Console.Clear();
                Console.WriteLine("\n Salvando Dados no Banco...");
                Thread.Sleep(3000);

                //Delete Emails
                var services = new DBService();
                var Emails = JsonConvert.SerializeObject(ListEmails);
                services.GravaEmails(Emails);
                Console.Clear();
            }
           

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

                    client.Inbox.Open(FolderAccess.ReadOnly);

                    int messageCount = client.Inbox.Count;
                    
                    Console.WriteLine("\n Email Conectado com Sucess!");                   
                    Console.WriteLine("\n Foram Encontrados "+messageCount.ToString()+" email(s)");

                    Thread.Sleep(2000);
                    List<object> allMessages = new List<object>(messageCount);

                    var mails = client.Inbox.Search(SearchQuery.All);                   
                    
                    Console.Clear();
       
                    Console.WriteLine(" Fazendo Downloads dos Emails...\n");
                    Thread.Sleep(3000);
                    var i = 1;
                    foreach (var mail in mails)
                    {
                        Console.SetCursorPosition(1, 1);
                        Console.Write("█");
                        Console.Write(i + "/" + messageCount);

                        var message = client.Inbox.GetMessage(mail);

                        message.WriteTo(string.Format("{0}.eml", mail));

                        Emails emails = new Emails();
                        emails.IdEmail = message.MessageId;
                        emails.Titulo = message.Subject;
                        emails.Data = message.Date.ToString();
                        emails.De = message.From[0].ToString().Split('<')[1].Replace(">", "");
                        emails.Para = message.To.ToString();
                        emails.Html = message.HtmlBody;
                        emails.Body = message.TextBody;
                        emails.CC = message.Cc;

                        foreach (var attachment in message.Attachments)
                        {
                            if (attachment is MessagePart)
                            {
                                var fileName = attachment.ContentDisposition?.FileName;
                                var rfc822 = (MessagePart)attachment;

                                if (string.IsNullOrEmpty(fileName))
                                    fileName = "attached-message.eml";

                                using (var stream = File.Create(ConfigurationManager.AppSettings["Folder"] + email + "_" + fileName))
                                    rfc822.Message.WriteTo(stream);
                            }
                            else
                            {
                                var part = (MimePart)attachment;
                                var fileName = part.FileName;
                                var caminho = ConfigurationManager.AppSettings["Folder"];
                                using (var stream = File.Create(caminho+mail.Id+"_"+fileName))
                                    part.Content.DecodeTo(stream);
                            }
                        }
                     
                        allMessages.Add(emails);
                        i++;
                    }

                    //Deletar Menssagens                          
                    Console.Clear();
                    Console.WriteLine(" Deletando Emails...");
                    Thread.Sleep(3000);
                    DeleteEmails(email, senha);

                    return allMessages;
                }
                #endregion
            }
            catch (Exception exe)
            {
                Console.WriteLine("\nErro: "+exe.Message);
                return null;
            }

        }

        public static void DeleteEmails(string email, string senha)
        {
            using (Pop3Client client = new Pop3Client())
            {
                client.Connect("mail.visaogrupo.com.br", 110, false);
                client.Authenticate(email.Replace(" ", ""), senha);

                int messageCount = client.GetMessageCount();

                if(messageCount != 0) client.DeleteMessages(0, messageCount);

                client.Dispose();
            }
        }
    }
}
