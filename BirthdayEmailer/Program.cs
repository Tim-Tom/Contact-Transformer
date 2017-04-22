using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using ContactCommon;
using RazorEngine;
using RazorEngine.Configuration;
using RazorEngine.Templating;

namespace BirthdayEmailer {
   class Program {
      static void Main(string[] args) {
         AppDomain.CurrentDomain.UnhandledException += UnhandledExceptionTrapper;
         string excelFilename = args[0];
         BirthdayModel birthdays = new BirthdayModel(ContactBuilder.ReadExcel(excelFilename), DateTime.Today);
         var config = new TemplateServiceConfiguration() {
            DisableTempFileLocking = true,
            TemplateManager = new ResolvePathTemplateManager(new string[] { AppDomain.CurrentDomain.BaseDirectory }),
            CachingProvider = new DefaultCachingProvider(t => { }),
         };
         using (var service = RazorEngineService.Create(config)) {
            service.Compile("MonthlyBirthdays", typeof(BirthdayModel));
            string template = service.Run("MonthlyBirthdays", modelType: typeof(BirthdayModel), model: birthdays);
            Console.WriteLine(template);
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
