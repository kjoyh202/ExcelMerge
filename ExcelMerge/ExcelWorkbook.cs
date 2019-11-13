using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelMerge
{
    public class ExcelWorkbook
    {
        public Dictionary<string, ExcelSheet> Sheets { get; private set; }
        public IWorkbook srcWb = null;

        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);

            var wb = new ExcelWorkbook();
            wb.srcWb = WorkbookFactory.Create(path);
            
            for (int i = 0; i < wb.srcWb.NumberOfSheets; i++)
            {
                var srcSheet = wb.srcWb.GetSheetAt(i);
                wb.Sheets.Add(srcSheet.SheetName, ExcelSheet.Create(srcSheet, config));
            }

            return wb;
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            if (Path.GetExtension(path) == ".csv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else if (Path.GetExtension(path) == ".tsv")
            {
                yield return System.IO.Path.GetFileName(path);
            }
            else
            {
                var wb = WorkbookFactory.Create(path);
                for (int i = 0; i < wb.NumberOfSheets; i++)
                    yield return wb.GetSheetAt(i).SheetName;
            }
        }

        private static ExcelWorkbook CreateFromCsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromCsv(path, config));

            return wb;
        }

        private static ExcelWorkbook CreateFromTsv(string path, ExcelSheetReadConfig config)
        {
            var wb = new ExcelWorkbook();
            wb.Sheets.Add(Path.GetFileName(path), ExcelSheet.CreateFromTsv(path, config));

            return wb;
        }

        private void SaveExcel(string path)
        {
            using(FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();

                foreach (KeyValuePair<string, ExcelSheet> kp in Sheets)
                {
                    var sheet = wb.CreateSheet(kp.Key);
                    var sheetData = kp.Value;

                    for (int i = 0; i < sheetData.Rows.Count; i++)
                    {
                        var rowData = sheetData.Rows[i];

                        for (int j = 0; j < rowData.Cells.Count; j++)
                        {
                            IRow row = sheet.CreateRow(i);
                            ICell cell = row.CreateCell(j);
                            cell.SetCellValue(rowData.Cells[j].Value);
                        }
                    }
                }

                wb.Write(stream);
            }    
        }
    }
}
