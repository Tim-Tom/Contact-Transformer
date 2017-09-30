using System;
using System.Net;
using System.Xml.Serialization;

namespace BirthdayEmailer {
   [Serializable]
   public class EmailAuthentication {
      [XmlElement(ElementName = "account")]
      public string Account { get; set; }

      [XmlElement(ElementName = "password")]
      public string Password { get; set; }

      public NetworkCredential AsCredential() {
         return new NetworkCredential(Account, Password);
      }
   }
}
