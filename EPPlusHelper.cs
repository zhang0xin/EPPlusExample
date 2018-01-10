using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;

namespace EPPlusExample
{
  class EPPlusHelper
  {
    string excelFile;
    public EPPlusHelper(string excelFile)
    {
      this.excelFile = excelFile;
    }
    public string GenerateCode()
    {
      StringBuilder code = new StringBuilder();
      using(ExcelPackage package = new ExcelPackage(new FileStream(excelFile, FileMode.Open)))
      {
        if (package != null)
        {
          code.AppendLine("ExcelWorksheet sheet");
          foreach(var sheet in package.Workbook.Worksheets)
          {
            string codeCreateSheet = "sheet = package.Workbook.Worksheets.Add(\"{0}\");";
            code.AppendLine(string.Format(codeCreateSheet, sheet.Name));
            foreach(string address in GetAddressList(sheet))
            {
              string codeSetValue = "sheet.Cells[\"{0}\"].Value = \"{1}\";";
              string cellValue = GetCellValue(sheet, address);
              if (!string.IsNullOrEmpty(cellValue))
              {
                if (sheet.MergedCells.Contains(address))
                {
                  string codeSetMerged = "sheet.Cells[\"{0}\"].Merge = true;";
                  code.AppendLine(string.Format(codeSetMerged, address));
                }
                code.AppendLine(string.Format(codeSetValue, address, cellValue));
              }
            }
          }
        }
      }
      return code.ToString();
    }
    public ICollection<string> GetAddressList(ExcelWorksheet sheet)
    {
      List<string> addressList = new List<string>();
      foreach(var address in sheet.MergedCells)
      {
        addressList.Add(address);
      }
      foreach(var range in sheet.Cells)
      {
        if (range.Merge) continue;
        addressList.Add(range.Address);
      }
      return addressList;
    }
    public string GetCellValue(ExcelWorksheet sheet, string address)
    {
      var val = sheet.Cells[address].Value;
      var arr = val as object[,];
      if (arr != null && arr.GetLength(0) > 0 && arr.GetLength(1) > 0)
        return arr[0,0] + "";
      else
        return val + "";
    }
  }
}
