using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ContactTransformer {
   class Household {
      public Household(List<Contact> contacts) {
         Contact first = contacts[0];
         Contacts = contacts;
         Address = first.Address;
         Phone = first.HomePhone;
         Email = first.Email;
         foreach(Contact c in contacts) {
            if (c.Email != Email) {
               Email = "";
               break;
            }
         }
         bool LastNameSame = contacts.All(c => c.Last == first.Last);
         if (LastNameSame) {
            Name = ListJoin(contacts.Select(c => c.First).ToArray()) + " " + first.Last;
         } else {
            Name = ListJoin(contacts.Select(c => c.First + " " + c.Last).ToArray());
         }
      }
      private static string ListJoin(string[] parts) {
         if (parts.Length == 1) {
            return parts[0];
         } else if (parts.Length == 2) {
            return parts[0] + " and " + parts[1];
         } else {
            return String.Join(", ", parts.Take(parts.Length - 1)) + ", and " + parts[parts.Length - 1];
         }
      }
      public readonly string Name;
      public readonly string Address;
      public readonly string Phone;
      public readonly string Email;
      public readonly List<Contact> Contacts;
   }
}
