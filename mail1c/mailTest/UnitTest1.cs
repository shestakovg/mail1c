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
            mail.ConnectIMAP("imap.ukr.net", 993, "gennadiy78@ukr.net", "jgnbvf");
            var msg = mail.GetMessages();
            mail.CloseConnection();
            //IMail mail = new Mail1C();
            //mail.ConnectPop3("imap.ukr.net", 993, "gennadiy78@ukr.net", "jgnbvf");
            //mail.GetMessages2("","","",100);
            //mail.ClosePop3();
        }

        
        
    }
}
