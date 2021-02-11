using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using RandomNameGeneratorLibrary;
using CommandLine;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ImportGenerator
{

   class Facility
   {
      public string Id;
      public string Name;
      public string City;
      public string State;
      public string ZoneId;
      public bool Active;
      public FacilitySchedule Schedule;
   }

   class FacilityScheduledEvent
   {
      public string FacilityId;
      public int DayOfWeek;
      public string DataTransmissionStart;
      public int DurationHours;
   }

   class FacilitySchedule : List<FacilityScheduledEvent>
   {
   }

   class Device
   {
      public string Name;
      public string SerialNumber;
      public string FacilityId;
      public bool Active;
   }

   class DevicesWorksheet
   {
      private const int DeviceNameCell = 1;
      private const int SerialCell = 2;
      private const int FacilityIdCell = 3;
      private const int DeviceStatusCell = 4;

      private Worksheet _worksheet;
      public DevicesWorksheet(Worksheet worksheet)
      {
         _worksheet = worksheet;
      }

      public void Init()
      {
         _worksheet.Name = "Devices";
         _worksheet.Cells[1, DeviceNameCell] = "Device name";
         _worksheet.Cells[1, SerialCell] = "Serial number";
         _worksheet.Cells[1, FacilityIdCell] = "Facility ID";
         _worksheet.Cells[1, DeviceStatusCell] = "Device Active=1/Inactive=0";
      }

      public void AddDevices(List<Device> devices)
      {
         int row = 2;
         var cells = _worksheet.Cells;
         foreach (var device in devices)
         {
            cells[row, DeviceNameCell] = device.Name;
            cells[row, SerialCell] = device.SerialNumber;
            cells[row, FacilityIdCell] = device.FacilityId;
            cells[row, DeviceStatusCell] = device.Active ? "1" : "0";
            ++row;
         }
      }

   }

   class FacilitiesWorkSheet
   {
      private const int IdCell = 1;
      private const int NameCell = 2;
      private const int CityCell = 3;
      private const int StateCell = 4;
      private const int ZoneCell = 5;
      private const int ActiveCell = 6;

      private Worksheet _worksheet;

      public FacilitiesWorkSheet(Worksheet worksheet)
      {
         _worksheet = worksheet;
      }

      public void Init()
      {
         _worksheet.Name = "Facilities";
         _worksheet.Cells[1, IdCell] = "Facility ID";
         _worksheet.Cells[1, NameCell] = "Facility name";
         _worksheet.Cells[1, CityCell] = "Facility city";
         _worksheet.Cells[1, StateCell] = "Facility state";
         _worksheet.Cells[1, ZoneCell] = "IANA zone ID";
         _worksheet.Cells[1, ActiveCell] = "Facility Active=1/Inactive=0";
      }

      public void AddFacilities(List<Facility> facilities)
      {
         int row = 2;
         var cells = _worksheet.Cells;
         foreach (var facility in facilities)
         {

            cells[row, IdCell] = facility.Id;
            cells[row, NameCell] = facility.Name;
            cells[row, CityCell] = facility.City;
            cells[row, StateCell] = facility.State;
            cells[row, ZoneCell] = facility.ZoneId;
            cells[row, ActiveCell] = facility.Active ? "1" : "0";
            ++row;
         }
      }
   }

   class FacilitiesScheduleWorksheet
   {
      private const int IdCell = 1;
      private const int DayOfWeekCell = 2;
      private const int TxStartCell = 3;
      private const int DurationCell = 4;


      private Worksheet _worksheet;
      private int _insertRow = 2;

      public FacilitiesScheduleWorksheet(Worksheet worksheet)
      {
         _worksheet = worksheet;
      }

      public void Init()
      {
         _worksheet.Name = "Facilities schedule";
         _worksheet.Cells[1, IdCell] = "Facility ID";
         _worksheet.Cells[1, DayOfWeekCell] = "Day of week";
         _worksheet.Cells[1, TxStartCell] = "Data transmission start (24 hour clock)";
         _worksheet.Cells[1, DurationCell] = "Duration (Hours)";
      }

      public void AddSchedule(FacilitySchedule schedule)
      {
         foreach (var evnt in schedule)
         {
            _worksheet.Cells[_insertRow, IdCell] = evnt.FacilityId;
            _worksheet.Cells[_insertRow, DayOfWeekCell] = evnt.DayOfWeek.ToString();
            _worksheet.Cells[_insertRow, TxStartCell] = evnt.DataTransmissionStart;
            _worksheet.Cells[_insertRow, DurationCell] = evnt.DurationHours.ToString();
            ++_insertRow;
         }

      }
   }


   class Program
   {

      class Options
      {
         [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
         public bool Verbose { get; set; }
         [Option('d', "devices", Required = true, HelpText = "Specify number of devices.")]
         public int NumDevices { get; set; }
         [Option('f', "facilities", Required = true, HelpText = "Specify number of facilities.")]
         public int NumFacilities { get; set; }
         [Option('s', "serial", Required = false, HelpText = "Specify the serial number prefix.")]
         public string SerialPrefix { get; set; } = "1X";
      }


      static void Main(params string[] args)
      {
         Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(Run);
      }

      static void Run(Options options)
      {
         var program = new Program(options);
         program.Run();
      }

      _Application _excel = new _Excel.Application();
      Workbook _workbook;
      private Options _options;
      private FacilitiesWorkSheet _facilitiesWorksheet;
      private FacilitiesScheduleWorksheet _facilitiesScheduleWorksheet;
      private DevicesWorksheet _devicesWorksheet;

      private List<Facility> _facilities = new List<Facility>();
      private List<Device> _devices = new List<Device>();

      Program(Options options)
      {
         _options = options;
      }

      public void Run()
      {
         var ts = DateTime.Now;
         var fileName = String.Format("{0}-{1:00}-{2:00}-{3:00}-{4:00}-{5:00}.xlsx", ts.Year, ts.Month, ts.Day, ts.Hour, ts.Minute, ts.Second);
         var path = Directory.GetCurrentDirectory();
         var filePath = $"{path}/{fileName}";

         Console.WriteLine($"Generating import document:");
         Console.WriteLine($" Devices:    {_options.NumDevices}");
         Console.WriteLine($" Facilities: {_options.NumFacilities}");
         Console.WriteLine($" S/N Prefix: {_options.SerialPrefix}");
         Console.WriteLine($" Output:     {filePath}");


         CreateData();
         LoadWorksheetsFromTemplate();

         _devicesWorksheet.AddDevices(_devices);
         _facilitiesWorksheet.AddFacilities(_facilities);

         foreach (var facility in _facilities)
         {
            _facilitiesScheduleWorksheet.AddSchedule(facility.Schedule);
         }



         File.Delete(filePath);
         _workbook.SaveAs(filePath);

         // Closing the file
         _workbook.Close();
      }


      public void CreateData()
      {
         var states = SourceData.States;
         var timeZones = SourceData.TimeZones;

         var placeGenerator = new PlaceNameGenerator();


         int timeZoneIndex = 0;
         int stateIndex = 0;

         for (int nFacility = 1; nFacility <= _options.NumFacilities; ++nFacility)
         {
            var timeZone = timeZones[timeZoneIndex];
            timeZoneIndex = ++timeZoneIndex % timeZones.Length;

            var state = states[stateIndex];
            stateIndex = ++stateIndex % states.Length;

            var city = placeGenerator.GenerateRandomPlaceName();

            var facility = new Facility()
            {
               Id = String.Format("Site{0:000}", nFacility),
               Name = String.Format("Facility{0:000}", nFacility),
               City = city,
               State = state,
               ZoneId = timeZone,
               Active = true
            };
            _facilities.Add(facility);

            // create scheduled events
            var schedule = new FacilitySchedule();
            facility.Schedule = schedule;
            int start = 0;
            for (int day = 1; day <= 7; ++day)
            {
               var scheduledEvent = new FacilityScheduledEvent()
               {

                  FacilityId = facility.Id,
                  DayOfWeek = day,
                  DataTransmissionStart = String.Format("00:{0:00}", start++),
                  DurationHours = 1
               };
               schedule.Add(scheduledEvent);
            }
         }

         // create devices
         // alternate putting devices into the facilities

         var facilityIndex = 0;
         for (int nDevice = 1; nDevice <= _options.NumDevices; ++nDevice)
         {
            var facility = _facilities[facilityIndex];
            facilityIndex = ++facilityIndex % _facilities.Count();
            var device = new Device()
            {
               Name = String.Format("Device{0:000}", nDevice),
               SerialNumber = String.Format("{0}{1:00000}",_options.SerialPrefix, nDevice),
               FacilityId = facility.Id,
               Active = true,
            };
            _devices.Add(device);
         }
      }


      private void LoadWorksheetsFromTemplate()
      {
         Application excel = new Application();
         var cwd = Directory.GetCurrentDirectory();
         var fullPath = $"{cwd}\\template.xlsx";
         _workbook = excel.Workbooks.Open(fullPath);
         var worksheets = _workbook.Worksheets;

         foreach (var ws in worksheets)
         {
            Worksheet worksheet = (Worksheet)ws;
            if (worksheet.Name == "Facilities")
            {
               _facilitiesWorksheet = new FacilitiesWorkSheet(worksheet);
            }
            else if (worksheet.Name == "Devices")
            {
               _devicesWorksheet = new DevicesWorksheet(worksheet);
            }
            else if (worksheet.Name == "Facilities schedule")
            {
               _facilitiesScheduleWorksheet = new FacilitiesScheduleWorksheet(worksheet);
            }
            Debug.WriteLine($"{worksheet.Name}");
         }
      }

      private void CreateWorksheetsFromScratch()
      {
         _workbook = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

         // create worksheets
         var worksheet = _workbook.Worksheets[1];
         _facilitiesWorksheet = new FacilitiesWorkSheet(worksheet);
         _facilitiesWorksheet.Init();

         worksheet = _excel.Worksheets.Add(After: worksheet);
         _facilitiesScheduleWorksheet = new FacilitiesScheduleWorksheet(worksheet);
         _facilitiesScheduleWorksheet.Init();

         worksheet = _excel.Worksheets.Add(After: worksheet);
         _devicesWorksheet = new DevicesWorksheet(worksheet);
         _devicesWorksheet.Init();
      }

   }

}