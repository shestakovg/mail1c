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
            mail.ConnectSmtp("smtp.ukr.net", "gennadiy78@ukr.net", "jgnbvf");
            for (int i = 0; i < 1; i++ )
                mail.SendMessage("shestakov.g@gmail.com", null, "d:\\D\\Projects\\Current\\fb\\Mail\\Examples\\VS2012\\SmtpSend\\Program.cs", "Привет", "test");
        }
    }
}
