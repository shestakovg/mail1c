using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Newtonsoft.Json;

namespace mail1c
{
    internal class MailJsonSerializer
    {
        private MailStructure[] _mailStructures;
        private string _description = "";
        public MailStructure[] MailStructures
        {
            set { this._mailStructures = value; }
            get { return this._mailStructures; }
        }

        public MailJsonSerializer(string description, MailStructure[] mailStructures)
        {
            this._description = description;
            this._mailStructures = mailStructures;
        }

        public String ToJSONRepresentation()
        {
            StringBuilder sb = new StringBuilder();
            JsonWriter jw = new JsonTextWriter(new StringWriter(sb));

            jw.Formatting = Formatting.Indented;
            jw.WriteStartObject();
            jw.WritePropertyName(this._description);
            jw.WriteStartArray();

            foreach(MailStructure mail in this._mailStructures)
            {
                jw.WriteStartObject();
                jw.WritePropertyName("Id");
                jw.WriteValue(mail.Id);
                jw.WritePropertyName("Uuid");
                jw.WriteValue(mail.Uuid);
                jw.WritePropertyName("MailDate");
                jw.WriteValue(mail.MailDate);
                jw.WritePropertyName("From");
                jw.WriteValue(mail.From);
                jw.WritePropertyName("Subject");
                jw.WriteValue(mail.Subject);
                jw.WritePropertyName("Body");
                jw.WriteValue(mail.Body);
                jw.WritePropertyName("AttachmentExist");
                jw.WriteValue((mail.AttachmentExist ? "1" : "0"));
                jw.WritePropertyName("AttachmentFiles");
                jw.WriteValue(mail.AttachmentFiles);
                jw.WriteEndObject();
            }
            jw.WriteEndArray();
            jw.WriteEndObject();
            return sb.ToString();
        }

        public string ToXMLRepresentation()
        {
            //FileStream fs = new FileStream("emails.xml", FileMode.OpenOrCreate);
            MemoryStream ms = new  MemoryStream();
            XmlSerializer x = new XmlSerializer(typeof(MailStructure[]));
            x.Serialize(ms, this._mailStructures);
            return Encoding.UTF8.GetString(ms.GetBuffer());
            
        }
    }
}
