namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; private set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }
        public ExcelCellStatus Status { get; private set; }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex)
        {
            Value = value;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;

            Status = ExcelCellStatus.Empty;

            if(value != null && value.Length > 0)
            {
                Status = ExcelCellStatus.Filled;
            }
        }

        public void SetValue(string value)
        {
            Value = value;
        }

        public void SetStatus(ExcelCellStatus status)
        {
            Status = status;
        }
    }
}
