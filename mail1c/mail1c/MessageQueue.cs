using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace mail1c
{
    internal class MessageQueue
    {
        private Guid _queueId;
        private MailStructure[] _mailStructures;
        public Guid QueueId
        {
            get { return _queueId; }
            set { _queueId = value; }
        }

        public MessageQueue(MailStructure[] mailStructures, Guid id)
        {
            this._queueId = id;
            this._mailStructures = mailStructures;
        }

        public MailStructure GetMailStructure(int index)
        {
            return this._mailStructures[index];
        }
    }
}
