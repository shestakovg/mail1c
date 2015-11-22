using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using mail1c;

namespace mailTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestSendMail()
        {
            IMail mail = new Mail1C();
            mail.ConnectSmtp("smtp.ukr.net", "gennadiy78@ukr.net", "jgnbvf", 465);
            for (int i = 0; i < 1; i++ )
                mail.SendMessage("shestakov.g@gmail.com", null, "d:\\D\\Projects\\Current\\fb\\Mail\\Examples\\VS2012\\SmtpSend\\Program.cs", "Привет", "test");
        }

        [TestMethod]
        public void TestRecive()
        {
            IMail mail = new Mail1C();
            //mail.ConnectIMAP("imap.ukr.net", 993, "mrzed@ukr.net", "Drag0n");
            mail.ConnectIMAP("imap.ukr.net", 993, "gennadiy78@ukr.net", "jgnbvf");
            var msg = mail.GetMessagesToStringJson("", "", "", 100);
            mail.CloseConnection();
            //IMail mail = new Mail1C();
            //mail.ConnectPop3("imap.ukr.net", 993, "gennadiy78@ukr.net", "jgnbvf");
            //mail.GetMessages2("","","",100);
            //mail.ClosePop3();
        }



        [TestMethod]
        public void TestRecive7()
        {
            IMail mail = new Mail1C();
            //mail.ConnectIMAP("imap.ukr.net", 993, "mrzed@ukr.net", "Drag0n");
            mail.ConnectIMAP("imap.ukr.net", 993, "gennadiy78@ukr.net", "jgnbvf");
           // var msg = mail.GetMessagesToStringJson("", "", "", 100);
            string s = mail.GetMessagesCount("", "", "", 100);
            int sepIndex = s.IndexOf("#");
            string merssageId = s.Substring(0, sepIndex);
            int msCount=Int32.Parse(s.Replace(merssageId + "#", ""));
            for (int i = 0; i < msCount; i++)
            {
                string val = mail.GetMessageField(merssageId, i, "AttachmentExist");
            }
            mail.RemoveMessages(merssageId);
            mail.CloseConnection();
            //IMail mail = new Mail1C();
            //mail.ConnectPop3("imap.ukr.net", 993, "gennadiy78@ukr.net", "jgnbvf");
            //mail.GetMessages2("","","",100);
            //mail.ClosePop3();
        }

        
        
    }
}
