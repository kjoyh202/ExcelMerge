﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge
{
    public class ExcelRow : IEquatable<ExcelRow>
    {
        public int Index { get; private set; }
        public List<ExcelCell> Cells { get; private set; }

        public ExcelRow(int index, IEnumerable<ExcelCell> cells)
        {
            Index = index;
            Cells = cells.ToList();
        }

        public override bool Equals(object obj)
        {
            var other = obj as ExcelRow;

            return Equals(other);
        }

        public override int GetHashCode()
        {
            var hash = 7;
            foreach (var cell in Cells)
            {
                hash = hash * 13 + cell.Value.GetHashCode();
            }

            return hash;
        }

        public bool Equals(ExcelRow other)
        {
            if (other == null)
                return false;

            return GetHashCode() == other.GetHashCode();
        }

        public bool IsBlank()
        {
            return Cells.All(c => string.IsNullOrEmpty(c.Value));
        }

        public void UpdateCells(IEnumerable<ExcelCell> cells)
        {
            Cells = cells.ToList();
        }

        public void AddCell(ExcelCell newCell)
        {
            for (int i = 0; i < Cells.Count; i++)
            {
                if (Cells[i].OriginalColumnIndex == newCell.OriginalColumnIndex)
                {
                    Cells[i].SetValue(newCell.Value);
                    return;
                }
            }

            if (Cells.Count <= newCell.OriginalColumnIndex)
            {
                for (int i = 0; i < newCell.OriginalColumnIndex + 1; i++)
                {
                    if(i >= Cells.Count)
                    {
                        Cells.Add(new ExcelCell(string.Empty, i, newCell.OriginalRowIndex));
                    }                    

                    if(i == newCell.OriginalColumnIndex)
                    {
                        Cells[i].SetValue(newCell.Value);
                        return;
                    }
                }
            }
        }

        public ExcelCell GetCellWithOriginalColum(int column)
        {
            foreach(var cell in Cells)
            {
                if(cell.OriginalColumnIndex == column)
                {
                    return cell;
                }
            }

            return null;
        }
    }

    internal class RowComparer : IEqualityComparer<ExcelRow>
    {
        public HashSet<int> IgnoreColumns { get; private set; }

        public RowComparer(HashSet<int> ignoreColumns)
        {
            IgnoreColumns = ignoreColumns;
        }

        public bool Equals(ExcelRow x, ExcelRow y)
        {
            return GetHashCode(x).Equals(GetHashCode(y));
        }

        public int GetHashCode(ExcelRow obj)
        {
            var hash = 7;
            var index = 0;
            foreach (var cell in obj.Cells)
            {
                if (IgnoreColumns.Contains(index))
                    continue;

                hash = hash * 13 + cell.Value.GetHashCode();

                index++;
            }

            return hash;
        }
    }
}
