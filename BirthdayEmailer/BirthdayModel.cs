using ContactCommon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BirthdayEmailer {
   public class BirthdayModel {
      internal BirthdayModel(IEnumerable<Contact> allContacts, DateTime month) {
         _month = month;
         _contacts = (from c in allContacts
                      where c.Birthday.HasValue && c.Birthday.Value.Month == month.Month
                      orderby c.Birthday.Value.Day
                      select c).ToList();
      }

      private readonly DateTime _month;
      private readonly List<Contact> _contacts;
      public DateTime Month {
         get {
            return _month;
         }
      }
      public IEnumerable<Contact> Contacts {
         get {
            return _contacts;
         }
      }
      public int Age(Contact c) {
         return _month.Year - c.Birthday.Value.Year;
      }
   }
}
