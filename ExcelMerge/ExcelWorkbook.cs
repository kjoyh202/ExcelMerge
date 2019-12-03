using System;
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

        public string FilePath = string.Empty;

  
        public ExcelWorkbook()
        {
            Sheets = new Dictionary<string, ExcelSheet>();
        }

        public static ExcelWorkbook Create(string path, ExcelSheetReadConfig config, bool enableEdit = false)
        {
            if (Path.GetExtension(path) == ".csv")
                return CreateFromCsv(path, config);

            if (Path.GetExtension(path) == ".tsv")
                return CreateFromTsv(path, config);


            var wb = new ExcelWorkbook();

            FileAccess access = enableEdit == true ? FileAccess.ReadWrite : FileAccess.Read;

            if (Path.GetExtension(path) == ".xlsm" || Path.GetExtension(path) == ".xlsx" || Path.GetExtension(path) == ".tmp")
            {
                FileStream stream = new FileStream(path, FileMode.OpenOrCreate, access, FileShare.ReadWrite);                
                wb.srcWb = new XSSFWorkbook(stream);
            }
            else
            {
                wb.srcWb = WorkbookFactory.Create(path);
            }           
            
            //wb.srcWb.Close();
            wb.FilePath = path;

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


        public void SetCellValue(string sheetName, int row, int column, string value)
        {
            var sheetIndex = srcWb.GetSheetIndex(sheetName);
            var srcSheet = srcWb.GetSheetAt(sheetIndex);

            if (srcSheet == null)
                return;

            var srcRow = srcSheet.GetRow(row);

            //create rows if not exist.
            if(srcRow == null)
            {
                for(int i = 0; i < row + 1; i++)
                {
                    if(srcSheet.GetRow(i) == null)
                    {
                        srcSheet.CreateRow(i);
                    }
                }

                srcRow = srcSheet.CreateRow(row);
            }

            var cell = srcRow.GetCell(column);

            //create empty cell if not exist.
            if(cell == null)
            {
                for (int i = 0; i < column + 1; i++)
                {
                    if (srcRow.GetCell(i) == null)
                    {
                        srcRow.CreateCell(i);
                    }
                }

                cell = srcRow.CreateCell(column);
            }            

            cell.SetCellValue(value);

            return;
        }

        public void SaveExcel(string path)
        {

            //            var stream = new FileStream(FilePath + ".xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite);
            using (var fs = new FileStream(FilePath + ".xlsm", FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {
                srcWb.Write(fs);
            }
            



            

            /*
            using (FileStream stream = new FileStream(path + "_TMP3.xlsx", FileMode.Create))
            {
                
                try
                {
                    srcWb.Write(stream);
                    stream.Close();
                    
                    var newWorkBook = new XSSFWorkbook();

                    for (int i = 0; i < srcWb.NumberOfSheets; i++)
                    {
                        var srcSheet = srcWb.GetSheetAt(i);
                        var newSheet = newWorkBook.CreateSheet(srcSheet.SheetName);

                        for(int rowIdx = 0; rowIdx < srcSheet.LastRowNum+1; rowIdx++)
                        {
                            if(srcSheet.GetRow(rowIdx) != null)
                            {
                                var row = srcSheet.GetRow(rowIdx);

                                var newRow = newSheet.CreateRow(rowIdx);

                                foreach(var cell in row.Cells)
                                {
                                    var newCell = newRow.CreateCell(cell.ColumnIndex);
                                    newCell.SetCellValue("test");
                                }
                            }
                        }
                    }                   

                    newWorkBook.Write(stream);
                    

                }
                catch(Exception e)
                {
                    Console.Error.WriteLine(e);
                }
                
                
            }
            */
        }
    }
}
