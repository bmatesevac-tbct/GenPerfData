using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

using _Excel = Microsoft.Office.Interop.Excel;


namespace GenPerfData
{


   class Program
   {
      static void Main(string[] args)
      {
         var program = new Program();
         program.Run(args);
      }




      _Application _excel = new _Excel.Application();
      Workbook _workbook;
      Worksheet _worksheet;

      public void NewFile()
      {
         _workbook = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
         _worksheet = _workbook.Worksheets[1];
         _worksheet.Name = "Worksheet0";
      }

      // Method adds a new worksheet to the existing workbook 
      public Worksheet NewSheet()
      {
         return _excel.Worksheets.Add(After: _worksheet);
      }


      public void Run(string[] args)
      {
         // Create a new workbook with a single sheet
         NewFile();

         // Add a new sheet to the workbook
         var sheet = NewSheet();
         sheet.Name = "MyWorksheet";
         sheet = NewSheet();
         sheet.Name = "MyWorksheetAgain";



         var path = Directory.GetCurrentDirectory();
         var filePath = $"{path}/MyTestData.xlsx";

         File.Delete(filePath);
         _workbook.SaveAs(filePath);

         // Closing the file
         _workbook.Close();
      }
   }

}