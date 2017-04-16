using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace ContactTransformer {
   class Contact {
      public Contact(Excel.ListColumns columns, Excel.ListRow data) {
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
            case "Address 1":
               Address1 = ((String)c.Value ?? "").Trim();
               address.Append(Address1);
               break;
            case "Address 2":
               Address2 = ((String)c.Value ?? "").Trim();
               if (Address2 != "") {
                  address.AppendLine();
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
      public readonly string Address1;
      public readonly string Address2;
      public readonly string Address;
   }
}
