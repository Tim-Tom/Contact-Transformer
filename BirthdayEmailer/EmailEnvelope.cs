using System;
using System.Xml.Serialization;

namespace BirthdayEmailer {
   [Serializable]
   public class EmailEnvelope {
      [XmlElement(ElementName = "from")]
      public EmailAddress From { get; set; }

      [XmlElement(ElementName = "to")]
      public EmailAddress To { get; set; }

      [XmlElement(ElementName = "self")]
      public EmailAddress Self { get; set; }
   }
}
