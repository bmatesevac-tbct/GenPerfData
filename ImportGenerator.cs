using System;
using System.IO;


namespace ImportGenerator
{
   public class ImportGenerator
   {
      private Options _options;

      public ImportGenerator(Options options)
      {
         _options = options;
      }

      public void Run()
      {
         var ts = DateTime.Now;
         var fileName = String.Format("{0}-{1:00}{2:00}-{3:00}{4:00}{5:00}-F{6}-D{7}.xlsx", ts.Year, ts.Month, ts.Day, ts.Hour, ts.Minute, ts.Second, _options.NumFacilities, _options.NumDevices);
         var cwd = Directory.GetCurrentDirectory();
         var filePath = $"{cwd}/{fileName}";
         var templateFilePath = $"{cwd}/template.xlsx";


         Console.WriteLine($"Generating import document:");
         Console.WriteLine($" Devices:    {_options.NumDevices}");
         Console.WriteLine($" Facilities: {_options.NumFacilities}");
         Console.WriteLine($" S/N Prefix: {_options.SerialPrefix}");
         Console.WriteLine($" Duration:   {_options.Duration}");
         Console.WriteLine($" Output:     {filePath}");

         var targetObjects = new TargetObjects(_options);
         var document = new OutputDocument(templateFilePath);
         document.LoadTargetObjects(targetObjects);
         document.SaveAndClose(filePath);
      }
   }
}
