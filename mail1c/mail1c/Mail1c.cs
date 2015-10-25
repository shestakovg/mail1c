using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Limilabs.Client.IMAP;
using Limilabs.Mail.Fluent.Mail;

namespace mail1c
{
    [Guid("AAFB965C-8872-4199-95CC-AE66AB674499"), ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(typeof(IMailEvents))]
    public class Mail1C:IMail
    {
        private Imap _imap;
        
        public bool ConnectIMAP(string server, int port, string username, string password)
        {
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

        public bool SendMessage(string to, string tc, string[] attachments)
        {
            throw new NotImplementedException();
        }
    }
}
