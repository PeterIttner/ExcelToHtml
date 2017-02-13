using ExcelToHtmlConverter.Api;
using ExcelToHtmlConverter.Core.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelToHtmlConverter.Core
{
    public class ExcelConverter : IExcelConverter
    {
        #region Member

        private IRenderer renderer;
        private const string DEFAULT_TEMPLATE = "Template\\template.html";

        #endregion

        #region Ctor

        /// <summary>
        /// Creates a new instance of the <see cref="ExcelConverter"/> class.
        /// </summary>
        public ExcelConverter() : this(new HtmlRenderer()) { }

        /// <summary>
        /// Creates a new instance of the <see cref="ExcelConverter"/> class.
        /// </summary>
        /// <param name="renderer">The renderer that should be used for conversion</param>
        public ExcelConverter(IRenderer renderer)
        {
            this.renderer = renderer;
        }

        #endregion

        #region Public Interface

        /// <summary>
        /// Converts the worksheet from the given file to a string.
        /// </summary>
        /// <param name="filename">The path to the worksheet file</param>
        /// <returns>The converted worksheet as string</returns>
        public string ConvertWorksheet(string filename)
        {
            return ConvertWorksheet(filename, DEFAULT_TEMPLATE);
        }

        /// <summary>
        /// Converts the worksheet from the given file to a string.
        /// Applies the given template file for the conversion.
        /// </summary>
        /// <param name="filename">The path to the worksheet file</param>
        /// <param name="templateFile">The path to the template file that is used for the conversion</param>
        /// <returns>The converted worksheet as string</returns>
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

        #endregion

    }
}