using System;
using System.Xml.Serialization;

namespace BirthdayEmailer {
   [Serializable]
   [XmlRoot(ElementName = "config")]
   public class EmailConfig {

      [XmlElement(ElementName = "mail")]
      public EmailEnvelope Envelope { get; set; }

      [XmlElement(ElementName = "authentication")]
      public EmailAuthentication Authentication { get; set; }

   }
}
