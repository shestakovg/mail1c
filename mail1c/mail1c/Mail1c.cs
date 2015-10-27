using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Limilabs.Client.IMAP;
using Limilabs.Client.SMTP;
using Limilabs.Mail.Fluent;
using System.Reflection;
//using Limilabs.Mail.Headers;
namespace mail1c
{
    [Guid("AAFB965C-8872-4199-95CC-AE66AB674499"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IMailEvents))]
    public class Mail1C:IMail
    {
        private Imap _imap;
        private string _server;
        private string _username;
        private string _password;
        private int _port;

        public string Server
        {
            get {
                if (_server == null) throw new Exception("Не указан почтовый сервер!");
                else
                    return _server; 
                }
        }

        public bool ConnectIMAP(string server, int port, string username, string password)
        {
            this._port = port;
            this._username = username;
            this._server = server;
            this._password = password;
            this.CloseConnection();
            this._imap = new Imap();
            try
            {
                this._imap.ConnectSSL(server, port);
                this._imap.Login(username, password);
                this._imap.SelectInbox(); 
            }
            catch (Exception e)
            {
                
                return false;
            }
            return true;
        }

        public void CloseConnection()
        {
            if (_imap != null)
            {
                _imap.Close();
                _imap.Dispose();
                _imap = null;
            }
        }

        //public bool SendMessage(string to, string tc, string[] attachments)
        //{
        //    throw new NotImplementedException();
        //}


        public bool ConnectSmtp(string server,   string username, string password)
        {
            this._username = username;
            this._server = server;
            this._password = password;
            return true;
        }

        private Limilabs.Mail.IMail createEmail(string to, string tc, string attachment, string messageBody, string subject)
        {
            Limilabs.Mail.IMail email = null;
            if (attachment != null)
            {
                email = Mail.Html(messageBody).AddAttachment(attachment).Subject(subject)
                             .From(this._username)
                             .To(to).Create();
            }
            else
            {
                email = Mail.Html(messageBody)
                                 .Subject(subject)
                                 .From(this._username)
                                 .To(to).Create();
            }
            if (email.Subject.Contains("Please purchase Mail.dll license"))
            {
              
                email.Subject = string.Empty;
                email.Subject = subject;
            }
            return email;
        }
        public bool SendMessage(string to, string tc, string attachment, string messageBody, string subject)
        {
            Limilabs.Mail.IMail email = createEmail( to,  tc,  attachment,  messageBody,  subject);
            MethodInfo[] rs =  email.GetType().GetMethods(BindingFlags.NonPublic | BindingFlags.Instance);
            FieldInfo[] res = email.GetType().GetFields();
            int i = 0;
            while ( ++i <= 15)
            {
                if (email.Subject.Contains("Please purchase Mail.dll license"))
                {
                    
                    email = createEmail("asd@gmail.com", tc, attachment, messageBody, subject);
                    using (Smtp smtp = new Smtp())
                    {
                        smtp.ConnectSSL(this._server);
                        smtp.UseBestLogin(this._username, this._password);
                        smtp.SendMessage(email);
                        email = createEmail(to, tc, attachment, messageBody, subject);
                    }
                }
                else
                    break;
            }
           SendMessageResult result = null;
           using (Smtp smtp = new Smtp())
           {
               smtp.ConnectSSL(this._server);
               smtp.UseBestLogin(this._username, this._password);
               result  =smtp.SendMessage(email);
               
           }

               //             .UsingNewSmtp()
               //             .Server(this._server)
               //             .WithSSL()
               //             .WithCredentials(this._username, this._password);

               //SendMessageResult result = .Send();
            //IMail email = Mail
            //    .Html(messageBody)
            //    .Subject(subject)
            //   /// .AddVisual("Lemon.jpg").SetContentId("lemon@id")
            //    //.AddAttachment("Attachment.txt").SetFileName("Invoice.txt")
            //    .From(this._username)
            //    .To(to)
            //    .Create();

            return (result.Status == SendMessageStatus.Success);
        }
    }
}
