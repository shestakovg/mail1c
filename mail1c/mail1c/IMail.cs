using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace mail1c
{
    [Guid("36637F39-9209-419C-AC11-CC32C6E2BFBE")]
    public interface IMail
    {
        [DispId(1)]
        bool ConnectIMAP(string server, int port, string username, string password);
        bool ConnectSmtp(string server,   string username, string password, int port);
        void CloseConnection();
        //bool SendMessage(string to, string tc, string attachment, string messageBody, string subject);
        bool SendMessage(string to,string subject, string messageBody = "", string tc = "", string attachment ="" );
        MailStructure[] GetMessages(string datefrom = "", string dateto = "", string email = "", int limit = 10, int exitAttachment = 0);
        //MailStructure[] GetMessages2(string datefrom = "", string dateto = "", string email = "", int limit = 1000, int exitAttachment = 0);
        //bool ConnectPop3(string server, int port, string username, string password);
        //void ClosePop3();
        bool GetAttachments(string uuid, string path);
    }

    [Guid("A4420A21-152A-4D2A-9FCB-7FE20E5D1B71"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMailEvents
    {

    }
}
