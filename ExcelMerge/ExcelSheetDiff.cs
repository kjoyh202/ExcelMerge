using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    [Serializable]
    public class ExcelSheetDiff
    {
        public SortedDictionary<int, ExcelRowDiff> Rows { get; private set; }

        public ExcelSheetDiff()
        {
            Rows = new SortedDictionary<int, ExcelRowDiff>();
        }


        public ExcelRowDiff CreateRow()
        {
            var row = new ExcelRowDiff(Rows.Any() ? Rows.Keys.Last() + 1 : 0);
            Rows.Add(row.Index, row);

            return row;
        }

        public ExcelSheetDiffSummary CreateSummary()
        {
            var addedRowCount = 0;
            var removedRowCount = 0;
            var modifiedRowCount = 0;
            var modifiedCellCount = 0;
            foreach (var row in Rows)
            {
                if (row.Value.IsAdded())
                    addedRowCount++;
                else if (row.Value.IsRemoved())
                    removedRowCount++;

                if (row.Value.IsModified())
                    modifiedRowCount++;

                modifiedCellCount += row.Value.ModifiedCellCount;
            }

            return new ExcelSheetDiffSummary
            {
                AddedRowCount = addedRowCount,
                RemovedRowCount = removedRowCount,
                ModifiedRowCount = modifiedRowCount,
                ModifiedCellCount = modifiedCellCount,
            };
        }


        public void ReplaceCell(int row, int column, ExcelCell cell)
        {
            if (Rows.Count < row)
                return;

            foreach(var _cell in Rows[row].Cells)
            {
                if(_cell.Value.ColumnIndex == column)
                {
                    _cell.Value.DstCell.SetValue(cell.Value);
                    return;
                }
            }            
        }


    }
}
