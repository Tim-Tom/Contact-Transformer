using System;
using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace ContactCommon {
   public static class ContactBuilder {
      public static Tuple<List<Household>, List<Contact>> ReadExcel(string filename) {
         Excel.Application excel = null;
         Excel.Workbook workbook = null;
         List<Household> Households;
         List<Contact> Contacts;
         try {
            excel = new Excel.Application();
            excel.Visible = false;
            workbook = excel.Workbooks.Open(filename, ReadOnly: true);
            Excel.Worksheet sheet = workbook.Sheets[1];
            Excel.ListObject table = sheet.ListObjects["Information_Table"];
            List<ProtoContact> contacts = new List<ProtoContact>();
            foreach (Excel.ListRow row in table.ListRows) {
               ProtoContact c = new ProtoContact(table.ListColumns, row);
               contacts.Add(c);
            }
            Dictionary<string, bool> multiple = contacts.GroupBy(c => c.First).ToDictionary(cg => cg.Key, cg => cg.Count() > 1);
            Contacts = contacts.Select(c => {
               string shortName = multiple[c.First] ? c.First + " " + c.Last.Substring(0, 1) + "." : c.First;
               return new Contact(c, shortName);
            }).ToList();
            Households = (
               from c in Contacts
               where c.Address != ""
               group c by c.Address into cg
               select new Household(cg.ToList())
            ).Concat(
               from c in Contacts
               where c.Address == ""
               select new Household(new List<Contact>() { c })
            ).ToList();
         } finally {
            if (workbook != null) {
               workbook.Close();
            }
            if (excel != null) {
               excel.Quit();
            }
         }
         return new Tuple<List<Household>, List<Contact>>(Households, Contacts);
      }
   }
}
