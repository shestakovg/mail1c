using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenPOP.POP3;

namespace pop3test
{
    class Program
    {
        static void Main(string[] args)
        {
            DateTime start = DateTime.Now;
            POPClient c = new POPClient("212.42.77.244", 110, "gennadiy78@ukr.net", "jgnbvf", AuthenticationMethod.USERPASS);
            c.ReceiveTimeOut = 10000000;
            int mc = c.GetMessageCount();
            for (int i = mc; i >= 1; i--)
            {

                OpenPOP.MIMEParser.Message m = c.GetMessageHeader(i);// c.GetMessage(i, false);
                
                Console.WriteLine(String.Format("{0} - {1}", i.ToString(), m.FromEmail));
                m = null;
            }
            c.Disconnect();
            Console.WriteLine(DateTime.Now - start);
        }
    }
}
