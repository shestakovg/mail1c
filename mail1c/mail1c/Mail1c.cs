using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using AE.Net.Mail;

namespace mail1c
{
    [Guid("AAFB965C-8872-4199-95CC-AE66AB674499"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IMailEvents))]
    public class Mail1C:IMail
    {
       // private Imap _imap;
        private string _server;
        private string _username;
        private string _password;
        private int _port;
        private List<MessageQueue> _messageQueue = new List<MessageQueue>();
        private List<MailMessage> messages = new List<MailMessage>();
//        private POPClient _popCLient = null;
        private ImapClient ic = null;
        public string Server
        {
            get {
                if (_server == null) throw new Exception("Не указан почтовый сервер!");
                else
                    return _server; 
                }
        }

        //public bool ConnectPop3(string server, int port, string username, string password)
        //{
        //    try
        //    {
        //        this._popCLient = new POPClient(server, 110, username, password, AuthenticationMethod.USERPASS);
        //        this._popCLient.ReceiveTimeOut = 10000000;
        //    }
        //    catch (Exception e)
        //    {

        //        return false;
        //    }
            
            
        //    return true;
        //}

        public bool ConnectIMAP(string server, int port, string username, string password)
        {
            this._port = port;
            this._username = username;
            this._server = server;
            this._password = password;

            try
            {
                ic = new ImapClient(server, username, password, AuthMethods.Login, port, true);
                ic.SelectMailbox("INBOX");
                return true;
            }
            catch (Exception)
            {
                this.CloseConnection();    
                return false;
            }
        }

        public MailStructure[] GetMessages(string datefrom = "", string dateto = "", string email = "", int limit = 0,
            int exitAttachment = 0)
        {
            if (this.ic == null) throw new Exception("IMAP connection isn't init");
            if (limit == 0) limit = 100;
            DateTime dateFrom = new DateTime(2000, 1, 1);
            DateTime dateTo = new DateTime(5000, 1, 1);
            int currentCount = 0;
            if (datefrom != String.Empty) dateFrom = new DateTime(Int32.Parse(datefrom.Substring(1, 4)), Int32.Parse(datefrom.Substring(5, 2)), Int32.Parse(datefrom.Substring(7, 2)));
            if (dateto != String.Empty) dateTo = new DateTime(Int32.Parse(dateto.Substring(1, 4)), Int32.Parse(dateto.Substring(5, 2)), Int32.Parse(dateto.Substring(7, 2)));
            int mc = this.ic.GetMessageCount();
            this.ic.IdleTimeout = 0;
            List<MailStructure> result = new List<MailStructure>();
            int startIndex = (mc - limit < 0 ? 0 : mc - limit);
            
            MailMessage[] mm = ic.GetMessages(startIndex, mc -1, false);//(mc-limit-1, limit);

            messages.Clear();
            int i = 0;
            //for (int i = mc; i >= 1; i--)
            foreach (MailMessage m in mm)
            {
                bool searchStatus = true;
                if ((exitAttachment > 0) && (m.Attachments.Count == 0)) searchStatus = false;
               
                if (searchStatus && (currentCount++ <= limit - 1))
                {
                    try
                    {
                        messages.Add(m);

                        MailStructure ms = new MailStructure()
                        {
                            Id = i,
                            MailDate = m.Date.ToString(),
                            From = m.From.ToString(),
                            Subject = m.Subject,
                            Body = m.Body,
                            Uuid = m.Uid,
                            AttachmentExist = (m.Attachments.Count > 0)
                        };

                        string att = "";
                        int attCnt = 1;
                        foreach (AE.Net.Mail.Attachment attachment in m.Attachments)
                        {
                            if (att != "") att += ";";
                            att += attachment.Filename;
                        }
                        ms.AttachmentFiles = att;
                        result.Add(
                            ms
                            );
                    }
                    catch (Exception)
                    {

                     
                    }
                }
                else
                {
                    break;
                }

            }

            return result.ToArray();

        }


        //public bool ConnectIMAP(string server, int port, string username, string password)
        //{
        //    this._port = port;
        //    this._username = username;
        //    this._server = server;
        //    this._password = password;
        //    this.CloseConnection();
        //    this._imap = new Imap();
            
        //    try
        //    {
        //        this._imap.ConnectSSL(server, port);
        //        this._imap.Login(username, password);
        //        this._imap.SelectInbox(); 
        //    }
        //    catch (Exception e)
        //    {
        //        this.CloseConnection();
        //        return false;
        //    }
        //    return true;
        //}

        public void CloseConnection()
        {
            if (this.ic != null)
            {
                ic.Disconnect();
                ic.Dispose();
                ic = null;
            }
        }

        //public bool SendMessage(string to, string tc, string[] attachments)
        //{
        //    throw new NotImplementedException();
        //}


        public bool ConnectSmtp(string server, string username, string password, int port)
        {
            this._username = username;
            this._server = server;
            this._password = password;
            this._port = port;
            return true;
        }

      // private Limilabs.Mail.IMail createEmail(string to, string tc, string attachment, string messageBody, string subject)
        //{
        //    Limilabs.Mail.IMail email = null;
        //    if (attachment != null)
        //    {
        //        email = Mail.Html(messageBody).AddAttachment(attachment).Subject(subject)
        //                     .From(this._username)
        //                     .To(to).Create();
        //    }
        //    else
        //    {
        //        email = Mail.Html(messageBody)
        //                         .Subject(subject)
        //                         .From(this._username)
        //                         .To(to).Create();
        //    }
        //    if (email.Subject.Contains("Please purchase Mail.dll license"))
        //    {
              
        //        email.Subject = string.Empty;
        //        email.Subject = subject;
        //    }
        //    return email;
        //}
        //public bool SendMessage(string to, string tc, string attachment, string messageBody, string subject)
        //{
        //    Limilabs.Mail.IMail email = createEmail( to,  tc,  attachment,  messageBody,  subject);
        //    MethodInfo[] rs =  email.GetType().GetMethods(BindingFlags.NonPublic | BindingFlags.Instance);
        //    FieldInfo[] res = email.GetType().GetFields();
        //    int i = 0;
        //    while ( ++i <= 15)
        //    {
        //        if (email.Subject.Contains("Please purchase Mail.dll license"))
        //        {
                    
        //            email = createEmail("asd@gmail.com", tc, attachment, messageBody, subject);
        //            using (Smtp smtp = new Smtp())
        //            {
        //                smtp.ConnectSSL(this._server);
        //                smtp.UseBestLogin(this._username, this._password);
        //                smtp.SendMessage(email);
        //                email = createEmail(to, tc, attachment, messageBody, subject);
        //            }
        //        }
        //        else
        //            break;
        //    }
        //   SendMessageResult result = null;
        //   using (Smtp smtp = new Smtp())
        //   {
        //       smtp.ConnectSSL(this._server);
        //       smtp.UseBestLogin(this._username, this._password);
        //       result  =smtp.SendMessage(email);
               
        //   }

        //       //             .UsingNewSmtp()
        //       //             .Server(this._server)
        //       //             .WithSSL()
        //       //             .WithCredentials(this._username, this._password);

        //       //SendMessageResult result = .Send();
        //    //IMail email = Mail
        //    //    .Html(messageBody)
        //    //    .Subject(subject)
        //    //   /// .AddVisual("Lemon.jpg").SetContentId("lemon@id")
        //    //    //.AddAttachment("Attachment.txt").SetFileName("Invoice.txt")
        //    //    .From(this._username)
        //    //    .To(to)
        //    //    .Create();

        //    return (result.Status == SendMessageStatus.Success);
        //}

        private System.Web.Mail.MailMessage createEmail2(string to, string tc, string attachment, string messageBody, string subject)
        {
            System.Web.Mail.MailMessage message = new System.Web.Mail.MailMessage();
            message.Fields.Add
            ("http://schemas.microsoft.com/cdo/configuration/smtpserver",
                          this._server);
            message.Fields.Add
                ("http://schemas.microsoft.com/cdo/configuration/smtpserverport",
                              "465");
            message.Fields.Add
                ("http://schemas.microsoft.com/cdo/configuration/sendusing",
                              "2");
            message.Fields.Add
        ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1");
            //Use 0 for anonymous
            message.Fields.Add
            ("http://schemas.microsoft.com/cdo/configuration/sendusername",
                this._username);
            message.Fields.Add
            ("http://schemas.microsoft.com/cdo/configuration/sendpassword",
                 this._password);
            message.Fields.Add
            ("http://schemas.microsoft.com/cdo/configuration/smtpusessl",
                 "true");
            message.From = this._username;
            message.To = to;
            if (!String.IsNullOrEmpty(tc))                message.Cc = tc;

            message.Subject = subject;
            message.Body = messageBody;
            if (attachment != null && attachment != String.Empty)
                message.Attachments.Add(new System.Web.Mail.MailAttachment(attachment));
            return message;
        }

        public bool SendMessage(string to, string subject, string messageBody = "", string tc = "", string attachment = "")
        {
            System.Web.Mail.SmtpMail.SmtpServer = this._server + ":" + this._port.ToString();
            try
            {
                System.Web.Mail.SmtpMail.Send(createEmail2(to, tc, attachment, messageBody, subject));
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
            
        }


        //public MailStructure[] GetMessages(string datefrom = "", string dateto = "", string email = "", int limit = 0, int exitAttachment = 0)
        //{
        //    if (_imap == null) throw new Exception("IMAP connection is't init");
        //    if (limit == 0) limit = 1000000;
        //    DateTime dateFrom = new DateTime(2000,1,1);
        //    DateTime dateTo=new DateTime(5000,1,1);
        //    int currentCount = 0;
        //    if (datefrom != String.Empty) dateFrom = new DateTime(Int32.Parse(datefrom.Substring(1,4)), Int32.Parse(datefrom.Substring(5,2)), Int32.Parse(datefrom.Substring(7,2)));
        //    if (dateto != String.Empty) dateTo = new DateTime(Int32.Parse(dateto.Substring(1,4)), Int32.Parse(dateto.Substring(5,2)), Int32.Parse(dateto.Substring(7,2)));

            
        //    SimpleImapQuery simpleImapQuery = new SimpleImapQuery();
        //    if (email != String.Empty) simpleImapQuery.From = email;

        //    List<long> reports = _imap.Search(simpleImapQuery);
        //    List<MailStructure> result = new List<MailStructure>();
        //    foreach (long uid in reports)
        //    {
        //        bool searchStatus = true;
        //        Limilabs.Mail.IMail mail = new MailBuilder().CreateFromEml(
        //            _imap.GetMessageByUID(uid));
        //        if (!((mail.Date >= dateFrom) && (mail.Date <= dateTo))) searchStatus = false;
        //        if ((exitAttachment > 0) && (mail.Attachments.Count == 0)) searchStatus = false;
        //        if (searchStatus && (currentCount++ <= limit-1))
        //        {
        //            MailStructure ms = new MailStructure()
        //            {
        //                Id = uid,
        //                MailDate = mail.Date.ToString(),
        //                From = mail.From.ToString(),
        //                Subject = mail.Subject,
        //                Body = mail.GetBodyAsText(),
        //                AttachmentExist = (mail.Attachments.Count > 0)
        //            };
        //            string att = "";
        //            for (int i = 0; i < mail.Attachments.Count; i++)
        //            {
        //                if (att != "") att += ";";
        //                att += mail.Attachments[i].FileName;
        //            }
        //            ms.AttachmentFiles = att;
        //            result.Add(
        //                ms
        //                );
        //        }
        //        else
        //        {
        //            break;
        //        }
                
        //    }
        //    return result.ToArray();
        //}
        //private string convertChars(string str)
        //{
        //    byte[] data = Convert.FromBase64String(str);
        //    return Encoding.UTF8.GetString(data);
        //    //string str = "UTF8 Encoded string.";
        //    Encoding srcEncodingFormat = Encoding.Unicode;
        //    Encoding dstEncodingFormat = Encoding.GetEncoding("windows-1251");
        //    byte[] originalByteString = srcEncodingFormat.GetBytes(str);
        //    byte[] convertedByteString = Encoding.Convert(srcEncodingFormat,
        //    dstEncodingFormat, originalByteString);
        //    string finalString = dstEncodingFormat.GetString(convertedByteString);
        //    return finalString;
        //}
        //public MailStructure[] GetMessages2(string datefrom = "", string dateto = "", string email = "", int limit = 0,
        //    int exitAttachment = 0)
        //{
        //    if (this._popCLient == null) throw new Exception("POP3 connection is't init");
        //    if (limit == 0) limit = 1000000;
        //    DateTime dateFrom = new DateTime(2000, 1, 1);
        //    DateTime dateTo = new DateTime(5000, 1, 1);
        //    int currentCount = 0;
        //    if (datefrom != String.Empty) dateFrom = new DateTime(Int32.Parse(datefrom.Substring(1, 4)), Int32.Parse(datefrom.Substring(5, 2)), Int32.Parse(datefrom.Substring(7, 2)));
        //    if (dateto != String.Empty) dateTo = new DateTime(Int32.Parse(dateto.Substring(1, 4)), Int32.Parse(dateto.Substring(5, 2)), Int32.Parse(dateto.Substring(7, 2)));
        //    int mc = this._popCLient.GetMessageCount();
        //    List<MailStructure> result = new List<MailStructure>();
        //    for (int i = mc; i >= 1; i--)
        //    {

        //        Message mail =this._popCLient.GetMessage(i,false);// c.GetMessage(i, false);
        //        //string msgBody = "";
        //        //mail.GetMessageBody(msgBody);
        //        //string rr= mail.GetTextBody(msgBody);
        //        bool searchStatus = true;
                
        //        //if (!((mail.Date >= dateFrom) && (mail.Date <= dateTo))) searchStatus = false;
        //        if ((exitAttachment > 0) && (mail.Attachments.Count == 0)) searchStatus = false;
        //        if (searchStatus && (currentCount++ <= limit-1))
        //        {
        //            MailStructure ms = new MailStructure()
        //            {
        //                Id = i,
        //                MailDate = mail.Date.ToString(),
        //                From = mail.From.ToString(),
        //                Subject = mail.Subject,
        //                //Body = mail.MessageBody.ToString(),
        //                AttachmentExist = (mail.Attachments.Count > 0)
        //            };
        //            for (int msBodyCount = 0;msBodyCount<= mail.MessageBody.Count - 1; msBodyCount++)
        //            {
        //              //ms.Body += convertChars(mail.MessageBody[msBodyCount].ToString());
                        
        //            }
                    
        //            string att = "";
        //            for (int j = 0; j < mail.Attachments.Count; j++)
        //            {
        //                if (att != "") att += ";";
        //                att += mail.Attachments[j];
        //            }
        //            ms.AttachmentFiles = att;
        //            result.Add(
        //                ms
        //                );
        //        }
        //        else
        //        {
        //            break;
        //        }
                
        //    }
        //    return result.ToArray();
            
        //}


        //public void ClosePop3()
        //{
        //    if (this._popCLient != null)
        //    {
        //        this._popCLient.Disconnect();
        //        this._popCLient = null;
        //    }
        //}


        public bool GetAttachments(string uuid, string path)
        {
            if (this.ic == null) return false;
            try
            {
                MailMessage m = this.ic.GetMessage(uuid, false);
                foreach (Attachment att in m.Attachments)
                {
                    att.Save(path + att.Filename);
                }
                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }


        public string GetMessagesToStringJson(string datefrom = "", string dateto = "", string email = "", int limit = 10, int exitAttachment = 0)
        {
            MailJsonSerializer serializer = new MailJsonSerializer("emails",
                GetMessages(datefrom, dateto, email, limit, exitAttachment));
            //return serializer.ToJSONRepresentation( );
            return serializer.ToJSONRepresentation();
        }


        public string GetMessagesToStringXml(string datefrom = "", string dateto = "", string email = "", int limit = 10, int exitAttachment = 0)
        {
            MailJsonSerializer serializer = new MailJsonSerializer("emails",
               GetMessages(datefrom, dateto, email, limit, exitAttachment));
            //return serializer.ToJSONRepresentation( );
            return serializer.ToXMLRepresentation();
        }


        public string GetMessagesCount(string datefrom = "", string dateto = "", string email = "", int limit = 10, int exitAttachment = 0)
        {
            MailStructure[] structure = GetMessages(datefrom, dateto, email, limit, exitAttachment);
            Guid id = Guid.NewGuid();
            this._messageQueue.Add(new MessageQueue(structure, id));
            return id.ToString() + "#" + structure.Length.ToString();
        }




        public string GetMessageField(string id, int messageNumber, string fieldName)
        {
            Guid iGuid;
            try
            {
                iGuid = new Guid(id);
            }
            catch (Exception e)
            {

                throw new Exception("Can't convert " + id + " to Guid");
            }
            MessageQueue queue = null;
            string result="";
            foreach (MessageQueue q in this._messageQueue)
            {
                if (q.QueueId == iGuid)
                {
                    queue = q;
                    MailStructure mailStructure= queue.GetMailStructure(messageNumber);
                    result = mailStructure.GetType().GetField(fieldName).GetValue(mailStructure).ToString();
                    //GetMembers(BindingFlags.NonPublic | BindingFlags.Instance);
                }
            }
            if (queue==null) throw new Exception("Messages with id "+id+" not exists");
            return result;
        }

        public void DeleteMessage(int messageId)
        {
            MailMessage m;
            if (messages.Count > messageId)
            {
                m = messages[messageId];
                //ic.SelectMailbox("Trash");
                // ic.DeleteMessage(m);
                //ic.SelectMailbox("Trash");
                try
                {
                    ic.MoveMessage(m.Uid, "Trash");
                }
                catch  { }
                
            }
        }
        public void RemoveMessages(string id)
        {
            Guid iGuid;
            try
            {
                iGuid = new Guid(id);
            }
            catch (Exception e)
            {
                
                throw new Exception("Can't convert "+id+" to Guid");
            }

            foreach (MessageQueue q in this._messageQueue)
            {
                if (q.QueueId == iGuid)
                {
                    this._messageQueue.Remove(q);
                    return;
                }
            }
        }
    }
}

