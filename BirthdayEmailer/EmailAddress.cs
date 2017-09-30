using System;
using System.Xml.Serialization;
using System.Net.Mail;

namespace BirthdayEmailer {
   [Serializable]
   public class EmailAddress {

      [XmlElement(ElementName = "name")]
      public string Name { get; set; }

      [XmlElement(ElementName = "address")]
      public string Address { get; set; }

      public MailAddress AsAddress() {
         return new MailAddress(Address, Name);
      }
   }
}
