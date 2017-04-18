using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

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
         string csvFilename = changeExtension(excelFilename, ".csv");
         ReadExcel(excelFilename);
         WriteText(textFilename);
         WriteWord(wordFilename);
         WriteVCard(vcardFilename);
         WriteCSV(csvFilename);
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
               if (h.Contacts.Count == 1) {
                  Contact c = h.Contacts[0];
                  if (c.CellPhone != "") {
                     writer.WriteLine("  Cell: {0}", c.CellPhone);
                  }
                  if (c.Email != "") {
                     writer.WriteLine("  Email: {0}", c.Email);
                  }
                  if (c.WorkEmail != "") {
                     writer.WriteLine("  Work Email: {0}", c.Email);
                  }
                  if (c.Birthday.HasValue) {
                     if (c.Birthday.Value > DateTime.Today) {
                        writer.WriteLine("  Birthday: {0:MMM d}", c.Birthday.Value);
                     } else {
                        writer.WriteLine("  Birthday: {0:MMM d, yyyy}", c.Birthday.Value);
                     }
                  }
               } else {
                  foreach (Contact c in h.Contacts) {
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
                     if (c.WorkEmail != "") {
                        writer.WriteLine("{0} Work Email: {1}", c.First, c.WorkEmail);
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
               int paraLines = 2;
               string content = h.Name;
               if (h.Phone != "") {
                  ++paraLines;
                  content += String.Format("\vPhone: {0}", h.Phone);
               }
               if (h.Contacts.Count == 1) {
                  Contact c = h.Contacts[0];
                  if (c.CellPhone != "") {
                     ++paraLines;
                     content += String.Format("\vCell: {0}", c.CellPhone);
                  }
                  if (c.Email != "") {
                     ++paraLines;
                     content += String.Format("\vEmail: {0}", h.Email);
                  }
                  if (c.WorkEmail != "") {
                     ++paraLines;
                     content += String.Format("\vWork Email: {0}", c.WorkEmail);
                  }
                  //if (c.Birthday.HasValue) {
                  //   ++paraLines;
                  //   content += String.Format("\vBirthday: {0:MMM d}", c.Birthday.Value);
                  //}
               } else {
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
                  //foreach (Contact c in h.Contacts) {
                  //   if (c.Birthday.HasValue) {
                  //      ++paraLines;
                  //      content += String.Format("\v{0} Birthday: {1:MMM d}", c.First, c.Birthday.Value);
                  //   }
                  //}
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
         Regex notDigit = new Regex(@"\D");
         Regex addr2Parse = new Regex(@"([^,]+),\s*(\S+)\s*(\d+)");
         using (TextWriter writer = new StreamWriter(filename, false, Encoding.UTF8)) {
            foreach (Contact c in Contacts) {
               if (c.Birthday.HasValue && c.Birthday <= DateTime.Today && c.Birthday > new DateTime(2010, 1, 1)) {
                  continue;
               }
               writer.WriteLine("BEGIN:VCARD");
               writer.WriteLine("VERSION:3.0");
               writer.WriteLine("FN:{0} {1}", c.First, c.Last);
               writer.WriteLine("N:{0};{1};;;", c.Last, c.First);
               if (c.HomePhone != "") {
                  writer.WriteLine("TEL;TYPE=HOME:{0}", notDigit.Replace(c.HomePhone, ""));
               }
               if (c.CellPhone != "") {
                  writer.WriteLine("TEL;TYPE=CELL:{0}", notDigit.Replace(c.CellPhone, ""));
               }
               if (c.Email != "") {
                  writer.WriteLine("EMAIL;TYPE=PERSONAL:{0}", c.Email);
               }
               if (c.WorkEmail != "") {
                  writer.WriteLine("EMAIL;TYPE=WORK:{0}", c.WorkEmail);
               }
               if (c.Birthday.HasValue) {
                  writer.WriteLine("BDAY:{0:yyyyMMdd}", c.Birthday.Value);
               }
               Match m = addr2Parse.Match(c.Address2);
               if (m.Success) {
                  writer.WriteLine("ADR;TYPE=HOME:;;{0};{1};{2};{3};", c.Address1, m.Groups[1].Value, m.Groups[2].Value, m.Groups[3].Value);
               }
               writer.WriteLine("KIND:INDIVIDUAL");
               writer.WriteLine("END:VCARD");
            }
         }
      }

      private static void WriteCSV(string filename) {
         string[] columns = new string[] {
            "First Name", "Middle Name", "Last Name", "Title", "Suffix",
            "Initials", "Web Page", "Gender", "Birthday", "Anniversary",
            "Location", "Language", "Internet Free Busy", "Notes", "E-mail Address",
            "E-mail 2 Address", "E-mail 3 Address", "Primary Phone", "Home Phone", "Home Phone 2",
            "Mobile Phone", "Pager", "Home Fax", "Home Address", "Home Street",
            "Home Street 2", "Home Street 3", "Home Address PO Box", "Home City", "Home State",
            "Home Postal Code", "Home Country", "Spouse", "Children", "Manager's Name",
            "Assistant's Name", "Referred By", "Company Main Phone", "Business Phone", "Business Phone 2",
            "Business Fax", "Assistant's Phone", "Company", "Job Title", "Department",
            "Office Location", "Organizational ID Number", "Profession", "Account", "Business Address",
            "Business Street", "Business Street 2", "Business Street 3", "Business Address PO Box", "Business City",
            "Business State", "Business Postal Code", "Business Country", "Other Phone", "Other Fax",
            "Other Address", "Other Street", "Other Street 2", "Other Street 3", "Other Address PO Box",
            "Other City", "Other State", "Other Postal Code", "Other Country", "Callback",
            "Car Phone", "ISDN", "Radio Phone", "TTY/TDD Phone", "Telex",
            "User 1", "User 2", "User 3", "User 4", "Keywords",
            "Mileage", "Hobby", "Billing Information", "Directory Server", "Sensitivity",
            "Priority", "Private", "Categories"
         };
         Regex addr2Parse = new Regex(@"([^,]+),\s*(\S+)\s*(\d+)");
         using (TextWriter writer = new StreamWriter(filename, false, Encoding.UTF8)) {
            writer.WriteLine(String.Join(", ", columns));
            foreach (Contact c in Contacts) {
               if (c.Birthday.HasValue && c.Birthday <= DateTime.Today && c.Birthday > new DateTime(2010, 1, 1)) {
                  continue;
               }
               Dictionary<string, string> cc = new Dictionary<string, string>() {
               { "First Name", c.First },
               { "Last Name", c.Last },
            };
               Match m = addr2Parse.Match(c.Address2);
               if (m.Success) {
                  cc["Home Address"] = '"' + c.Address.Replace("\n", Environment.NewLine) + '"';
                  cc["Home Street"] = c.Address1;
                  cc["Home City"] = m.Groups[1].Value;
                  cc["Home State"] = m.Groups[2].Value;
                  cc["Home Postal Code"] = m.Groups[3].Value;
               }
               if (c.HomePhone != "") {
                  cc["Home Phone"] = c.HomePhone;
               }
               if (c.CellPhone != "") {
                  cc["Mobile Phone"] = c.CellPhone;
               }
               if (c.Email != "") {
                  cc["E-Mail Address"] = c.Email;
               }
               if (c.WorkEmail != "") {
                  cc["E-Mail Address 2"] = c.WorkEmail;
               }
               if (c.Birthday.HasValue) {
                  cc["Birthday"] = String.Format("{0:yyyy-MM-dd}", c.Birthday.Value);
               }
               writer.WriteLine(String.Join(",", columns.Select(k => cc.ContainsKey(k) ? cc[k] : "")));
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
