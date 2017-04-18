using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ContactTransformer {
   class Program {

      private static List<Household> Households;
      private static List<Contact> Contacts;
      private static string changeExtension(string filename, string extension) {
         return Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + extension);
      }

      

      static void Main(string[] args) {
         string excelFilename = args[0];
         string textFilename = changeExtension(excelFilename, ".txt");
         string wordFilename = changeExtension(excelFilename, ".docx");
         string vcardFilename = changeExtension(excelFilename, ".vcf");
         ReadExcel(excelFilename);
         // WriteText(textFilename);
         // WriteWord(wordFilename);
         WriteVCard(vcardFilename);
      }

      private static void ReadExcel(string filename) {
         AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
         Excel.Application excel = null;
         Excel.Workbook workbook = null;
         try {
            excel = new Excel.Application();
            excel.Visible = false;
            workbook = excel.Workbooks.Open(filename, ReadOnly: true);
            Excel.Worksheet sheet = workbook.Sheets[1];
            Excel.ListObject table = sheet.ListObjects["Information_Table"];
            Contacts = new List<Contact>();
            foreach (Excel.ListRow row in table.ListRows) {
               Contact c = new Contact(table.ListColumns, row);
               Contacts.Add(c);
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
               from c in Contacts
               where c.Address != ""
               group c by c.Address into cg
               select new Household(cg.ToList())
            ).Concat(
               from c in Contacts
               where c.Address == ""
               select new Household(new List<Contact>() { c })
            ).ToList(); 
            foreach(var cg in Contacts.GroupBy(c => c.First)) {
               List<Contact> group = cg.ToList();
               if (group.Count > 1) {
                  foreach (Contact c in group) {
                     c.ShortName = c.First + " " + c.Last.Substring(0, 1) + ".";
                  }
               } else {
                  group[0].ShortName = group[0].First;
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

      private static void WriteText(string filename) {
         using (TextWriter writer = new StreamWriter(filename, false, Encoding.UTF8)) {
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
                     if (c.Birthday.Value > DateTime.Today) {
                        writer.WriteLine("  {0} Birthday: {1:MMM d}", c.First, c.Birthday.Value);
                     } else {
                        writer.WriteLine("  {0} Birthday: {1:MMM d, yyyy}", c.First, c.Birthday.Value);
                     }
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

      private static void WriteWord(string filename) {
         Word.Application word = null;
         Word.Document doc = null;
         try {
            word = new Word.Application();
            word.Visible = false;
            doc = word.Documents.Add(Visible: false);
            {
               Word.PageSetup ps = doc.PageSetup;
               ps.Orientation = Word.WdOrientation.wdOrientLandscape;
               ps.TopMargin = ps.BottomMargin = ps.LeftMargin = ps.RightMargin = word.InchesToPoints(0.5f);
               Word.TextColumns ts = ps.TextColumns;
               ts.SetCount(4);
            }
            int columnLines = 0;
            Word.Paragraph p = null;
            foreach (Household h in Households) {
               int paraLines = 1;
               string content = h.Name;
               if (h.Phone != "") {
                  ++paraLines;
                  content += String.Format("\vPhone: {0}", h.Phone);
               }
               foreach (Contact c in h.Contacts) {
                  if (c.CellPhone != "") {
                     ++paraLines;
                     content += String.Format("\v{0} Cell: {1}", c.First, c.CellPhone);
                  }
               }
               if (h.Email != "") {
                  ++paraLines;
                  content += String.Format("\vEmail: {0}", h.Email);
               } else {
                  foreach (Contact c in h.Contacts) {
                     if (c.Email != "") {
                        ++paraLines;
                        content += String.Format("\v{0} Email: {1}", c.First, c.Email);
                     }
                  }
               }
               foreach (Contact c in h.Contacts) {
                  if (c.WorkEmail != "") {
                     ++paraLines;
                     content += String.Format("\v{0} Work Email: {1}", c.First, c.Email);
                  }
               }
               foreach (Contact c in h.Contacts) {
                  if (c.Birthday.HasValue) {
                     ++paraLines;
                     content += String.Format("\v{0} Birthday: {1:MMM d}", c.First, c.Birthday.Value);
                  }
               }
               if (h.Address != "") {
                  ++paraLines;
                  content += "\vAddress:";
                  foreach (string line in h.Address.Split('\n')) {
                     ++paraLines;
                     content += String.Format("\v\t{0}", line);
                  }
               }
               if (columnLines + paraLines > 50) {
                  content = (char)14 + content;
                  columnLines = paraLines;
               } else {
                  columnLines += paraLines;
               }
               p = doc.Paragraphs.Add();
               p.Range.Text = content;
               p.Range.Font.Bold = 0;
               p.Range.Font.Size = 8.0f;
               p.TabStops.Add(16.0f);
               doc.Range(p.Range.Start, p.Range.Start + content.IndexOf('\v')).Font.Bold = 1;
               p.Range.InsertParagraphAfter();
            }
            p = doc.Paragraphs.Add();
            p.Range.InsertBreak(Word.WdBreakType.wdPageBreak);
            p.Range.InsertParagraphAfter();
            Dictionary<int, List<Contact>> birthdays = Contacts
               .Where(c => c.Birthday.HasValue)
               .GroupBy(c => c.Birthday.Value.Month)
               .ToDictionary(
                  cg => cg.Key,
                  cg => cg.OrderBy(c => c.Birthday).ToList()
               );
            foreach(var month in birthdays.OrderBy(k => k.Key)) {
               string content = month.Value[0].Birthday.Value.ToString("MMMM");
               foreach(Contact c in month.Value.OrderBy(k => k.Birthday.Value.Day)) {
                  content += String.Format("\v\t{0:MMM-d}: {1}", c.Birthday.Value, c.ShortName);
                  if (c.Birthday.Value < DateTime.Today) {
                     content += String.Format(" ({0})", c.Birthday.Value.Year);
                  }
               }
               p = doc.Paragraphs.Last;
               p.Range.Text = content;
               p.Range.Font.Bold = 0;
               p.Range.Font.Size = 12.0f;
               p.TabStops.Add(16.0f);
               doc.Range(p.Range.Start, p.Range.Start + content.IndexOf('\v')).Font.Bold = 1;
               p.Range.InsertParagraphAfter();
            }
         } finally {
            if (doc != null) {
               doc.SaveAs2(FileName: filename);
               doc.Close();
            }
            if (word != null) {
               word.Quit();
            }
         }
      }

      private static void WriteVCard(string filename) {
         using (TextWriter writer = new StreamWriter(filename, false, Encoding.UTF8)) {
            foreach (Household h in Households) {

            }
         }
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
