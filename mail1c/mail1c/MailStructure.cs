using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mail1c
{
  
    public class MailStructure
    {
        public long Id;
        public string Uuid;
        public string MailDate;
        public string From;
        public string Subject;
        public string Body;
        public bool AttachmentExist;
        public string AttachmentFiles;
    }
}
