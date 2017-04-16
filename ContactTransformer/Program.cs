using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ContactTransformer {
   class Program {

      private static List<Contact> contacts;
      private static Dictionary<string, List<Contact>> households;
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
            contacts = new List<Contact>();
            households = new Dictionary<string, List<Contact>>();
            foreach (Excel.ListRow row in table.ListRows) {
               Contact c = new Contact(table.ListColumns, row);
               contacts.Add(c);
               if (c.Address != "") {
                  List<Contact> household;
                  if (!households.TryGetValue(c.Address, out household)) {
                     households[c.Address] = household = new List<Contact>();
                  }
                  household.Add(c);
               }
            }
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
            foreach(Contact c in contacts) {
               if (seen.Contains(c)) {
                  continue;
               }
               IEnumerable<Contact> household;
               if (c.Address == "") {
                  household = new Contact[] { c };
               } else {
                  household = households[c.Address];
               }
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
