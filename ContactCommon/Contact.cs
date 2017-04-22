using System;
using System.Collections;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace ContactCommon {
   internal class ProtoContact {
      public ProtoContact(Excel.ListColumns columns, Excel.ListRow data) {
         IEnumerator columnEnumerator = columns.GetEnumerator();
         StringBuilder address = new StringBuilder();
         foreach(Excel.Range c in data.Range) {
            columnEnumerator.MoveNext();
            Excel.ListColumn col = (Excel.ListColumn)columnEnumerator.Current;
            switch(col.Name) {
            case "First":
               First = ((String)c.Value ?? "").Trim();
               break;
            case "Last":
               Last = ((String)c.Value ?? "").Trim();
               break;
            case "Birthday":
               Birthday = c.Value;
               break;
            case "Home":
               HomePhone = ((String)c.Value ?? "").Trim();
               break;
            case "Cell":
               CellPhone = ((String)c.Value ?? "").Trim();
               break;
            case "Email":
               Email = ((String)c.Value ?? "").Trim();
               break;
            case "Work Email":
               WorkEmail = ((String)c.Value ?? "").Trim();
               break;
            case "Address 1":
               Address1 = ((String)c.Value ?? "").Trim();
               address.Append(Address1);
               break;
            case "Address 2":
               Address2 = ((String)c.Value ?? "").Trim();
               if (Address2 != "") {
                  address.Append('\n');
                  address.Append(Address2);
               }
               break;
            default:
               throw new Exception("Unknown Table Column " + col.Name);
            }
         }
         Address = address.ToString();
      }
      public readonly string First;
      public readonly string Last;
      public readonly DateTime? Birthday;
      public readonly string HomePhone;
      public readonly string CellPhone;
      public readonly string Email;
      public readonly string WorkEmail;
      public readonly string Address1;
      public readonly string Address2;
      public readonly string Address;
   }
   public class Contact {
      internal Contact(ProtoContact pc, string shortName) {
         this.First = pc.First;
         this.Last = pc.Last;
         this.Birthday = pc.Birthday;
         this.HomePhone = pc.HomePhone;
         this.CellPhone = pc.CellPhone;
         this.Email = pc.Email;
         this.WorkEmail = pc.WorkEmail;
         this.Address1 = pc.Address1;
         this.Address2 = pc.Address2;
         this.Address = pc.Address;
         this.ShortName = shortName;
      }
      public readonly string First;
      public readonly string Last;
      public readonly string ShortName;
      public readonly DateTime? Birthday;
      public readonly string HomePhone;
      public readonly string CellPhone;
      public readonly string Email;
      public readonly string WorkEmail;
      public readonly string Address1;
      public readonly string Address2;
      public readonly string Address;
   }
}
