using System;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Xml.Serialization;

using RazorEngine.Configuration;
using RazorEngine.Templating;

using ContactCommon;

namespace BirthdayEmailer {
   class Program {

      static EmailConfig GetConfig(string filename) {
         XmlSerializer serializer = new XmlSerializer(typeof(EmailConfig));
         using (TextReader reader = new StreamReader(filename)) {
            return (EmailConfig)serializer.Deserialize(reader);
         }
      }

      static void Main(string[] args) {
         AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
         string excelFilename = args[0];
         string emailConfigFilename = args[1];
         EmailConfig eConfig = GetConfig(emailConfigFilename);
         BirthdayModel birthdays = new BirthdayModel(ContactBuilder.ReadExcel(excelFilename), DateTime.Today);
         if (birthdays.Contacts.Count() == 0) {
            return;
         }
         var tConfig = new TemplateServiceConfiguration() {
            DisableTempFileLocking = true,
            TemplateManager = new ResolvePathTemplateManager(new string[] { AppDomain.CurrentDomain.BaseDirectory }),
            CachingProvider = new DefaultCachingProvider(t => { }),
         };
         using (var service = RazorEngineService.Create(tConfig)) {
            service.Compile("MonthlyBirthdays", typeof(BirthdayModel));
            string template = service.Run("MonthlyBirthdays", modelType: typeof(BirthdayModel), model: birthdays);
            using (SmtpClient client = new SmtpClient() {
               Host = "smtp.gmail.com",
               Port = 587,
               EnableSsl = true,
               DeliveryMethod = SmtpDeliveryMethod.Network,
               UseDefaultCredentials = false,
               Credentials = eConfig.Authentication.AsCredential()
            }) {
               using (var message = new MailMessage(eConfig.Envelope.From.AsAddress(), eConfig.Envelope.To.AsAddress()) {
                  Subject = String.Format("{0:MMMM} Birthdays", birthdays.Month),
                  IsBodyHtml = true,
                  Body = template
               }) {
                  message.Bcc.Add(eConfig.Envelope.Self.AsAddress());
                  client.Send(message);
               }
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
