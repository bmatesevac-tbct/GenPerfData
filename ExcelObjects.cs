using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using _Excel = Microsoft.Office.Interop.Excel;


namespace ImportGenerator
{

   public class DevicesWorksheet
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

   public class FacilitiesWorkSheet
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

   public class FacilitiesScheduleWorksheet
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


   public class OutputDocument
   {
      Workbook _workbook;
      private FacilitiesWorkSheet _facilitiesWorksheet;
      private FacilitiesScheduleWorksheet _facilitiesScheduleWorksheet;
      private DevicesWorksheet _devicesWorksheet;

      public OutputDocument(string templateFilePath = null)
      {
         if (templateFilePath != null)
         {
            CreateFromTemplate(templateFilePath);
         }
         else
         {
            CreateFromScratch();
         }
      }

      private void CreateFromScratch()
      {
         _Application excel = new _Excel.Application();
         _workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

         // create worksheets
         var worksheet = _workbook.Worksheets[1];
         _facilitiesWorksheet = new FacilitiesWorkSheet(worksheet);
         _facilitiesWorksheet.Init();

         worksheet = excel.Worksheets.Add(After: worksheet);
         _facilitiesScheduleWorksheet = new FacilitiesScheduleWorksheet(worksheet);
         _facilitiesScheduleWorksheet.Init();

         worksheet = excel.Worksheets.Add(After: worksheet);
         _devicesWorksheet = new DevicesWorksheet(worksheet);
         _devicesWorksheet.Init();
      }

      private void CreateFromTemplate(string templateFilePath)
      {
         Application excel = new Application();
         var cwd = Directory.GetCurrentDirectory();
         var fullPath = $"{cwd}\\template.xlsx";
         _workbook = excel.Workbooks.Open(fullPath);
         var worksheets = _workbook.Worksheets;

         foreach (var ws in worksheets)
         {
            Worksheet worksheet = (Worksheet) ws;
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

      public void LoadTargetObjects(TargetObjects targetObjects)
      {
         _devicesWorksheet.AddDevices(targetObjects.Devices);
         _facilitiesWorksheet.AddFacilities(targetObjects.Facilities);
         foreach (var facility in targetObjects.Facilities)
         {
            _facilitiesScheduleWorksheet.AddSchedule(facility.Schedule);
         }
      }

      public void SaveAndClose(string filePath)
      {
         File.Delete(filePath);
         _workbook.SaveAs(filePath);
         _workbook.Close();
      }


   }
}
