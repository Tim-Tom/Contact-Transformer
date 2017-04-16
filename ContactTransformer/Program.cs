using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ContactTransformer {
   class Program {

      private static List<Household> Households;
      private static string changeExtension(string filename, string extension) {
         return Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + extension);
      }
      static void Main(string[] args) {
         string excelFilename = args[0];
         string textFilename = changeExtension(excelFilename, ".txt");
         string wordFilename = changeExtension(excelFilename, ".docx");
         ReadExcel(excelFilename);
         WriteText(textFilename);
         WriteWord(wordFilename);
      }

      private static void ReadExcel(string excelFilename) {
         AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
         Excel.Application excel = null;
         Excel.Workbook workbook = null;
         try {
            excel = new Excel.Application();
            workbook = excel.Workbooks.Open(excelFilename, ReadOnly: true);
            Excel.Worksheet sheet = workbook.Sheets[1];
            Excel.ListObject table = sheet.ListObjects["Information_Table"];
            List<Contact> contacts = new List<Contact>();
            foreach (Excel.ListRow row in table.ListRows) {
               Contact c = new Contact(table.ListColumns, row);
               contacts.Add(c);
            }
            //households = contacts
            //   .Where(c => c.Address != "")
            //   .GroupBy(c => c.Address)
            //   .Select(cg => new Household(cg.ToList()))
            //.Union(contacts
            //   .Where(c => c.Address == "")
            //   .Select(c => new Household(new List<Contact>() { c }))
            //).ToList();
            Households = (
               from c in contacts
               where c.Address != ""
               group c by c.Address into cg
               select new Household(cg.ToList())
           ).Concat(
               from c in contacts
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
      }

      private static void WriteText(string textFilename) {
         HashSet<Contact> seen = new HashSet<Contact>();
         using (TextWriter writer = new StreamWriter(textFilename, false, Encoding.UTF8)) {
            foreach(Household h in Households) {
               writer.WriteLine(h.Name);
               if (h.Phone != "") {
                  writer.WriteLine("  Phone: {0}", h.Phone);
               }
               foreach(Contact c in h.Contacts) {
                  if (c.CellPhone != "") {
                     writer.WriteLine("  {0} Cell: {1}", c.First, c.CellPhone);
                  }
               }
               if (h.Email != "") {
                  writer.WriteLine("  Email: {0}", h.Email);
               } else {
                  foreach (Contact c in h.Contacts) {
                     if (c.Email != "") {
                        writer.WriteLine("  {0} Email: {1}", c.First, c.Email);
                     }
                  }
               }
               foreach (Contact c in h.Contacts) {
                  if (c.Birthday.HasValue) {
                     writer.WriteLine("  {0} Birthday: {1:MMM d, yyyy}", c.First, c.Birthday.Value);
                  }
               }
               if (h.Address != "") {
                  writer.WriteLine("  Address:");
                  foreach(string line in h.Address.Split('\n')) {
                     writer.WriteLine("    {0}", line);
                  }
               }
               writer.WriteLine();
            }
         }
      }

      private static void WriteWord(string wordFilename) {
         // throw new NotImplementedException();
      }

      private static void UnhandledExceptionTrapper(object sender, UnhandledExceptionEventArgs eargs) {
         if (!System.Diagnostics.Debugger.IsAttached) {
            Exception e = eargs.ExceptionObject as Exception;
            if (e == null) {
               e = new Exception("Unknown Error");
            }
            Console.Error.WriteLine(e);
            Environment.Exit(1);
         }
      }
   }
}
