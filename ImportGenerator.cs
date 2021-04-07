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
         var baseFileName = String.Format("{0}-{1:00}{2:00}-{3:00}{4:00}{5:00}", ts.Year, ts.Month, ts.Day, ts.Hour, ts.Minute, ts.Second);
         var fileName = baseFileName;
         var cwd = Directory.GetCurrentDirectory();
         var templateFilePath = $"{cwd}/Template.xlsx";

         int numDevices = 0;
         int numFacilities = 0;

         Console.WriteLine($"Generating import document:");
         if (_options.FacilityGroupSpecifiers != null)
         {
            Console.WriteLine($" Facility Groups:");
            foreach (var facilityGroupSpec in _options.FacilityGroupSpecifiers)
            {
               numFacilities += facilityGroupSpec.NumFacilities;
               numDevices += facilityGroupSpec.NumDevices * facilityGroupSpec.NumFacilities;
               Console.WriteLine($"   {facilityGroupSpec.NumFacilities} @ {facilityGroupSpec.NumDevices}");
               fileName += $"-{facilityGroupSpec.NumFacilities}@{facilityGroupSpec.NumDevices}";
            }
         }
         else
         {
            fileName += $"-F{_options.NumFacilities}-D{_options.NumDevices}-R{_options.Duration}";
            numDevices = _options.NumDevices;
            numFacilities = _options.NumFacilities;
         }

         if (_options.OutputFileName != null)
         {
            fileName = _options.OutputFileName;
         }

         var filePath = $"{cwd}/{fileName}.xlsx";

         Console.WriteLine($" Devices:    {numDevices}");
         Console.WriteLine($" Facilities: {numFacilities}");
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
