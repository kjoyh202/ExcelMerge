using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media;
using FastWpfGrid;

namespace ExcelMerge.GUI.Models
{
    public class SheetGridModel : FastGridModelBase
    {
        private int columnCount;
        private int rowCount;
        private Color? backgroundColor = null;
        private Color? decorationColor = null;
        private CellDecoration decoration = CellDecoration.None;
        private string toolTipText = string.Empty;
        private Dictionary<int, int> rowIndexMap = new Dictionary<int, int>();

        public override int ColumnCount
        {
            get { return columnCount; }
        }

        public override int RowCount
        {
            get { return rowCount; }
        }

        public override Color? BackgroundColor
        {
            get { return backgroundColor; }
        }

        public override Color? DecorationColor
        {
            get { return decorationColor; }
        }

        public override CellDecoration Decoration
        {
            get { return decoration; }
        }

        public override string ToolTipText
        {
            get { return toolTipText; }
        }

        public override TooltipVisibilityMode ToolTipVisibility
        {
            get { return TooltipVisibilityMode.OnlyWhenTrimmed; }
        }

        public int ColumnHeaderIndex { get; private set; }
        public int RowHeaderIndex { get; private set; } = -1;

        public ExcelSheet ExcelSheet { get; private set; }

        public SheetGridModel(ExcelSheet sheet) : base()
        {
            ExcelSheet = sheet;

            columnCount = ExcelSheet.Rows.Max(r => r.Value.Cells.Count);
            rowCount = ExcelSheet.Rows.Count();
            
            App.Instance.OnSettingUpdated += () => { InvalidateAll(); };
        }

        public override string GetColumnHeaderText(int column)
        {
            ExcelCell excelCell;
            if (TryGetCell(ColumnHeaderIndex, column, out excelCell))
                return GetCellText(excelCell);

            return string.Empty;
        }

        public override string GetRowHeaderText(int row)
        {
            if (RowHeaderIndex < 0)
                return base.GetRowHeaderText(row);

            ExcelCell excelCell;
            if (TryGetCell(row, RowHeaderIndex, out excelCell))
                return GetCellText(excelCell);

            return string.Empty;
        }

        public override string GetCellText(int row, int column)
        {
            return GetCellText(row, column, false);
        }

        public string GetCellText(int row, int column, bool direct)
        {
            ExcelCell excelCell;
            if (TryGetCell(row, column, out excelCell, direct))
                return GetCellText(excelCell);

            return string.Empty;
        }

        public string GetCellText(FastGridCellAddress address, bool direct)
        {
            if (address.IsEmpty)
                return string.Empty;

            return GetCellText(address.Row.Value, address.Column.Value, true);
        }

        
        private bool TryGetCell(int row, int column, out ExcelCell cell, bool direct = false)
        {
            ExcelCell excelCell = null;

            ExcelRow excelRow;
            if (TryGetRow(row, out excelRow, direct))
            {
                if(column < excelRow.Cells.Count)
                {
                    excelCell = excelRow.Cells[column];
                }                
            }

            if(excelCell == null)
            {
                cell = null;
                return false;
            }

            cell = excelCell;
            return true;
        }

        public override void SetCellText(int row, int column, string value)
        {
      
            ExcelCell newCell = new ExcelCell(value, column, row);
            ExcelSheet.ReplaceCell(row, column, newCell);
        }


        private bool TryGetRow(int row, out ExcelRow excelRow, bool direct = false)
        {
            if (direct)
                row = rowIndexMap.ContainsKey(row) ? rowIndexMap[row] : row;

            return ExcelSheet.Rows.TryGetValue(row, out excelRow);
        }
        

        private string GetCellText(ExcelCell excellCell)
        {
            return excellCell.Value;
        }

        private Color? GetColor(ExcelCellStatus status)
        {
            switch (status)
            {
                case ExcelCellStatus.Modified:
                    return App.Instance.Setting.ModifiedColor;
                case ExcelCellStatus.Added:
                    return App.Instance.Setting.AddedColor;
                case ExcelCellStatus.Removed:
                    return App.Instance.Setting.RemovedColor;
            }

            return null;
        }
               

        public override IFastGridCell GetRowHeader(IFastGridView view, int row)
        {
            var header = base.GetRowHeader(view, row) as SheetGridModel;
            if (header == null)
                return header;

            header.backgroundColor = App.Instance.Setting.RowHeaderColor;

            return header;
        }

        public override IFastGridCell GetColumnHeader(IFastGridView view, int column)
        {
            var header = base.GetColumnHeader(view, column) as SheetGridModel;
            if (header == null)
                return header;

            header.backgroundColor = App.Instance.Setting.ColumnHeaderColor;

            return header;
        }

        public IFastGridCell GetCell(IFastGridView view, int row, int column, bool direct)
        {
            toolTipText = GetCellText(row, column, direct);

            var cell = base.GetCell(view, row, column) as SheetGridModel;
            if (cell == null)
                return cell;

            ExcelCell excelCell;
            var status = ExcelCellStatus.None;
            
            if (TryGetCell(row, column, out excelCell, direct))
            {
                //do nothing.
            }            

            cell.backgroundColor = null;
            cell.backgroundColor = GetColor(status) ?? cell.backgroundColor;
            
            return cell;
        }

        public override IFastGridCell GetCell(IFastGridView view, int row, int column)
        {
            return GetCell(view, row, column, false);
        }

        private FastGridCellAddress GetVisualCellAddress(FastGridCellAddress realCellAddress)
        {
            if (realCellAddress.IsEmpty)
                return FastGridCellAddress.Empty;

            var swapped = rowIndexMap.ToDictionary(i => i.Value, i => i.Key);
            int visualRow;
            if (swapped.TryGetValue(realCellAddress.Row.Value, out visualRow))
                return new FastGridCellAddress(visualRow, realCellAddress.Column.Value, realCellAddress.IsGridHeader);

            return rowIndexMap.Any() ? FastGridCellAddress.Empty : realCellAddress;
        }

        private FastGridCellAddress GetRealCellAddress(FastGridCellAddress visualCellAddress)
        {
            if (visualCellAddress.IsEmpty)
                return FastGridCellAddress.Empty;

            int realRow;
            if (rowIndexMap.TryGetValue(visualCellAddress.Row.Value, out realRow))
                return new FastGridCellAddress(realRow, visualCellAddress.Column.Value, visualCellAddress.IsGridHeader);

            return rowIndexMap.Any() ? FastGridCellAddress.Empty : visualCellAddress;
        }

        private bool IsMatch(ExcelCellDiff cell, string text, bool exactMatch, bool caseSensitive, bool useRegex)
        {
            var srcValue = caseSensitive ? cell.SrcCell.Value : cell.SrcCell.Value.ToUpper();
            var dstValue = caseSensitive ? cell.DstCell.Value : cell.DstCell.Value.ToUpper();
            var target = caseSensitive ? text : text.ToUpper();

            if (useRegex)
            {
                var regex = new Regex(target);
                return regex.IsMatch(srcValue) || regex.IsMatch(srcValue);
            }

            if (exactMatch)
                return target == srcValue || target == dstValue;

            return srcValue.Contains(target) || dstValue.Contains(target);
        }


        public FastGridCellAddress GetNextCell(
                FastGridCellAddress current, Func<int, int> rowSelector, Func<int, List<ExcelCell>, FastGridCellAddress> selector)
        {
            if (current.IsEmpty)
                return current;

            var rowIndex = current.Row.Value;
            while (true)
            {
                ExcelRow excelRow;
                if (!TryGetRow(rowIndex, out excelRow, false))
                    break;

                var next = selector(rowIndex, excelRow.Cells);
                if (!next.IsEmpty)
                    return next;

                rowIndex = rowSelector(rowIndex);
            }

            return FastGridCellAddress.Empty;
        }

        public void SetRowHeader(int? column)
        {
            if (column.HasValue)
                RowHeaderIndex = column.Value;
        }

        public void SetRowHeader(string columnHeaderName)
        {
            for (int i = 0; i < columnCount; i++)
            {
                if (columnHeaderName == GetColumnHeaderText(i))
                {
                    SetRowHeader(i);
                    return;
                }
            }
        }

        public void SetColumnHeader(int? row)
        {
            if (row.HasValue)
                ColumnHeaderIndex = row.Value;
        }

        public void ShowEqualRows()
        {
            SetRowArrange(new HashSet<int>(), new HashSet<int>());

            rowIndexMap.Clear();
        }
    }
}
