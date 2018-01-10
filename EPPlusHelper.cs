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
          code.AppendLine("ExcelWorksheet sheet;");
          foreach(var sheet in package.Workbook.Worksheets)
          {
            string codeCreateSheet = "sheet = package.Workbook.Worksheets.Add(\"{0}\");";
            code.AppendLine(string.Format(codeCreateSheet, sheet.Name));
            foreach(string address in GetAddressList(sheet))
            {
              string codeSetValue = "sheet.Cells[\"{0}\"].Value = \"{1}\";";
              string cellValue = DistinctValue(sheet.Cells[address].Value)+"";
              code.AppendLine(GenerateStyleCodes(sheet.Cells[address]));
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
    public string GenerateStyleCodes(ExcelRange range)
    {
      StringBuilder codes = new StringBuilder();
      string codeFormat;
      codeFormat = "sheet.Cells[\"{0}\"].Style.Border.Left.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Border.Left.Style));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Border.Right.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Border.Right.Style));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Border.Top.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Border.Top.Style));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Border.Bottom.Style = "+
        " (ExcelBorderStyle) Enum.Parse(typeof(ExcelBorderStyle), \"{1}\");";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Border.Bottom.Style));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Numberformat.Format = \"{1}\";";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, EncodeCodeString(range.Style.Numberformat.Format)));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Font.Bold = {1};";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Font.Bold.ToString().ToLower()));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Font.Size = {1};";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Font.Size));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Font.Name = \"{1}\";";
      codes.AppendLine(string.Format(
        codeFormat, range.Address, range.Style.Font.Name));

      codeFormat = "sheet.Cells[\"{0}\"].Style.Font.Color.SetColor(Color.FromArgb({1}));";
      if (!string.IsNullOrEmpty(range.Style.Font.Color.Rgb))
      {
        codes.AppendLine(string.Format(
          codeFormat, range.Address, RgbToParameters(range.Style.Font.Color.Rgb)));
      }
      return codes.ToString();
    }
    public string RgbToParameters(string rgb)
    {
      //AARRGGBB 0xAA, 0xRR, 0xGG, 0xBB
      return rgb.Insert(6, ", 0x").Insert(4, ", 0x").Insert(2, ", 0x").Insert(0, "0x");
    }
    public string EncodeCodeString(string codes)
    {
      return codes.Replace(@"\", @"\\").Replace("\"", "\\\"");
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
    public object DistinctValue(object val)
    {
      var arr = val as object[,];
      if (arr != null && arr.GetLength(0) > 0 && arr.GetLength(1) > 0)
        return arr[0,0];
      else
        return val;
    }
  }
}
