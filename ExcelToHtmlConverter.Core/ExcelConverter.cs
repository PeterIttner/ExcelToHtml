using ExcelToHtmlConverter.Api;
using ExcelToHtmlConverter.Core.Model;
using Microsoft.CSharp;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using RazorEngine;
using RazorEngine.Templating;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToHtmlConverter.Core
{
    public class ExcelConverter : IExcelConverter
    {
        private IRenderer renderer;
        private const string DEFAULT_TEMPLATE = "Template\\template.html";

        public ExcelConverter() : this(new HtmlRenderer()) { }

        public ExcelConverter(IRenderer renderer)
        {
            this.renderer = renderer;
        }
        public string ConvertWorksheet(string filename)
        {
            return ConvertWorksheet(filename, DEFAULT_TEMPLATE);
        }

        public string ConvertWorksheet(string filename, string templateFile)
        {
            var workbook = new Workbook { Name = "TestWorkbook" };
            var fileinfo = new FileInfo(filename);

            if (!fileinfo.Exists)
            {
                throw new FileNotFoundException();
            }
            workbook.Name = fileinfo.Name;

            IWorkbook poiWorkbook;
            using (FileStream file = new FileStream(filename, FileMode.Open, FileAccess.Read))
            {
                if (filename.EndsWith(".xlsx"))
                {
                    poiWorkbook = new XSSFWorkbook(file);
                }
                else if (filename.EndsWith(".xls"))
                {
                    poiWorkbook = new HSSFWorkbook(file);
                }
                else
                {
                    throw new Exception("Not Supported File Format");
                }
            }

            var worksheets = new List<Worksheet>();
            for (int j = 0; j < poiWorkbook.NumberOfSheets; j++)
            {
                ISheet sheet = poiWorkbook.GetSheetAt(j);
                int rowCount = sheet.LastRowNum;
                IRow headerRow = sheet.GetRow(0);
                int colCount = 0;
                if (headerRow != null)
                {
                    colCount = headerRow.LastCellNum;
                    for (int rowNum = 0; rowNum < rowCount; rowNum++)
                    {
                        IRow hssfRow = sheet.GetRow(rowNum);
                        if (hssfRow != null)
                        {
                            colCount = hssfRow.LastCellNum > colCount ? hssfRow.LastCellNum : colCount;
                        }
                    }
                }

                var worksheet = new Worksheet();
                worksheet.Name = sheet.SheetName;
                var rows = new List<Row>();

                var sheetWidth = 0;
                for (int cellNum = 0; cellNum < colCount; cellNum++)
                {
                    sheetWidth += Convert.ToInt32(Math.Round(sheet.GetColumnWidthInPixels(cellNum)));
                }
                worksheet.WidthPixel = sheetWidth;

                for (int rowNum = 0; rowNum <= rowCount; rowNum++)
                {
                    var row = new Row();

                    var cells = new List<Cell>();

                    IRow hssfRow = sheet.GetRow(rowNum);

                    if (hssfRow != null)
                    {
                        for (int cellNum = 0; cellNum < colCount; cellNum++)
                        {
                            var cell = new Cell();
                            var hssfCell = hssfRow.GetCell(cellNum, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            if (hssfCell.CellStyle.FillPattern == FillPattern.SolidForeground)
                            {
                                byte[] cellBackground = hssfCell.CellStyle.FillForegroundColorColor.RGB;
                                var hexBackground = "#" + string.Concat(cellBackground.Select(b => b.ToString("X2")).ToArray());

                                cell.BackgroundColor = hexBackground;
                            }
                            if (hssfCell.CellStyle != null && hssfCell.CellStyle.GetFont(poiWorkbook) != null)
                            {
                                var foregroundColor = IndexedColors.ValueOf(hssfCell.CellStyle.GetFont(poiWorkbook).Color);
                                if (foregroundColor != null)
                                {
                                    cell.TextColor = foregroundColor.HexString;
                                }
                            }
                            if (hssfCell.CellType == CellType.Formula)
                            {
                                cell.Value = hssfCell.NumericCellValue.ToString();
                            }
                            else
                            {
                                cell.Value = hssfCell.ToString();
                            }
                            var cellWidthPixel = Convert.ToInt32(Math.Round(sheet.GetColumnWidthInPixels(cellNum)));
                            cell.WidthPixel = cellWidthPixel;

                            cells.Add(cell);
                        }
                    }

                    row.Cells = cells;
                    rows.Add(row);
                }
                worksheet.Rows = rows;
                worksheets.Add(worksheet);
            }

            workbook.Worksheets = worksheets;
            poiWorkbook = null;
            return renderer.RenderWorkbook(workbook, templateFile);
        }
    }
}
