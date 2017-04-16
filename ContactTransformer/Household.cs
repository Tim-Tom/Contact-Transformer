using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContactTransformer {
   class Household {
      public Household(List<Contact> contacts) {
         Address = contacts[0].Address;

      }
      string Address;
      string Phone;
      string Email;
      List<Contact> Contacts;
   }
}
