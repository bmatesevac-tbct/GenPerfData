using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

using _Excel = Microsoft.Office.Interop.Excel;


namespace Excel_with_C_Sharp
{
   class Excel
   {
      // Create an excel application object, workbook oject and worksheet object
      _Application excel = new _Excel.Application();
      Workbook workbook;
      Worksheet worksheet;

      // Method creates a new Excel file by creating a new Excel workbook with a single worksheet
      public void NewFile()
      {
         this.workbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
         this.worksheet = this.workbook.Worksheets[1];
      }

      // Method adds a new worksheet to the existing workbook 
      public Worksheet NewSheet()
      {
         return excel.Worksheets.Add(After: this.worksheet);
      }

      // Method saves workbook at a specified path
      public void SaveAs(string path)
      {
         workbook.SaveAs(path);
      }
   }


   class Program
   {

      _Application _excel = new _Excel.Application();
      Workbook _workbook;
      Worksheet _worksheet;


      static void Main(string[] args)
      {
         var program = new Program();
         program.Run(args);
      }


      public void NewFile()
      {
         _workbook = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
         _worksheet = _workbook.Worksheets[1];
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

         var path = Directory.GetCurrentDirectory();
         var filePath = $"{path}/MyTestData.xlsx";

         File.Delete(filePath);
         _workbook.SaveAs(filePath);

         // Closing the file
         _workbook.Close();
      }
   }

}