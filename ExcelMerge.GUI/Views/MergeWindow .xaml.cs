using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelMerge.GUI.Models;
using ExcelMerge.GUI.Settings;
using ExcelMerge.GUI.Shell;
using FastWpfGrid;
using Microsoft.Practices.Unity;

namespace ExcelMerge.GUI.Views
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MergeWindow : Window
    {
        private ExcelSheetDiffConfig diffConfig = new ExcelSheetDiffConfig();

        private GridLength previousConsoleHeight = new GridLength(0);

        private IUnityContainer container;

        private FastGridControl copyTargetGrid;

        private const string mergeKey = "merge";

        private DiffView _diffView = null;

        private string _dstFilePath = string.Empty;
        private string _mergeFilePath = string.Empty;

        private string _currentSheetName = string.Empty;

        private ExcelWorkbook excelWorkBook = null;
        private ExcelSheet excelSheet = null;

        public MergeWindow(DiffView diffView, string dstFilePath, string mergeFilePath, string currentSheetName, FileSetting fileSetting)
        {
            InitializeComponent();
            InitializeContainer();

            if(diffView == null)
            {
                this.Close();
            }

            _diffView = diffView;
            _dstFilePath = dstFilePath;
            _mergeFilePath = mergeFilePath;
            _currentSheetName = currentSheetName;

            var args = new DiffViewEventArgs<FastGridControl>(null, container, TargetType.First);
            DataGridEventDispatcher.Instance.DispatchPreExecuteDiffEvent(args);

            ReadWorkBook();

            MergeDataGrid.Model = new SheetGridModel(excelSheet);

            args = new DiffViewEventArgs<FastGridControl>(MergeDataGrid, container);
            DataGridEventDispatcher.Instance.DispatchFileSettingUpdateEvent(args, fileSetting);
            DataGridEventDispatcher.Instance.DispatchPostExecuteDiffEvent(args);

            InitCurrentCell();
        }

        private void InitializeContainer()
        {
            container = new UnityContainer();
            container
                .RegisterInstance(mergeKey, MergeDataGrid);
        }

        private void ReadWorkBook()
        {
            ProgressWindow.DoWorkWithModal(progress =>
            {
                progress.Report(Properties.Resources.Msg_ReadingFiles);

                var config = CreateReadConfig();
                excelWorkBook = ExcelWorkbook.Create(_mergeFilePath, config);
                excelSheet = excelWorkBook.Sheets[_currentSheetName];                
            });
        }

        private ExcelSheetReadConfig CreateReadConfig()
        {
            var setting = ((App)Application.Current).Setting;

            return new ExcelSheetReadConfig()
            {
                TrimFirstBlankRows = setting.SkipFirstBlankRows,
                TrimFirstBlankColumns = setting.SkipFirstBlankColumns,
                TrimLastBlankRows = setting.TrimLastBlankRows,
                TrimLastBlankColumns = setting.TrimLastBlankColumns,
            };
        }

        private void MenuItem_Loaded(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem == null)
                return;

            var binding = menuItem.GetBindingExpression(MenuItem.IsEnabledProperty);
            if (binding == null)
                return;

            binding.UpdateTarget();
        }

        private void LocationGrid_MouseDown(object sender, MouseEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
        }

        private void LocationGrid_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
                LocationGridEventDispatcher.Instance.DispatchMouseDownEvent(args, e);
            }
        }

        private void LocationGrid_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            var args = new DiffViewEventArgs<Grid>(sender as Grid, container);
            LocationGridEventDispatcher.Instance.DispatchMouseWheelEvent(args, e);
        }

        private void DataGrid_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSizeChangeEvent(args, e);
        }

        public void SetCurrentCell(int? row, int? column)
        {
            if (MergeDataGrid.CurrentCell.Row == row && MergeDataGrid.CurrentCell.Column == column)
                return;

            MergeDataGrid.CurrentCell = new FastGridCellAddress(row, column);
        }

        public void DataGrid_SelectedCellsChanged(object sender, FastWpfGrid.SelectionChangedEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;

            copyTargetGrid = grid;

            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSelectedCellChangeEvent(args);

            
            if (!MergeDataGrid.CurrentCell.Row.HasValue)
                return;

            if (!MergeDataGrid.CurrentCell.Column.HasValue)
                return;

            if (MergeDataGrid.Model == null)
                return;

            var srcValue =
                (MergeDataGrid.Model as SheetGridModel).GetCellText(MergeDataGrid.CurrentCell.Row.Value, MergeDataGrid.CurrentCell.Column.Value, true);            

         //   UpdateValue(srcValue, dstValue);
            
            if (App.Instance.Setting.AlwaysExpandCellDiff)
            {
                var a = new DiffViewEventArgs<RichTextBox>(null, container, TargetType.First);
                ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(a);
            }
            
            if(_diffView != null)
            {
                _diffView.SetCurrentCell(MergeDataGrid.CurrentCell.Row, MergeDataGrid.CurrentCell.Column);
            }
        }

        private void UpdateValue(string srcValue)
        {
            /*
            SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Clear();

            var srcLines = srcValue.Split('\n').Select(s => s.TrimEnd());
            var dstLines = dstValue.Split('\n').Select(s => s.TrimEnd());

            var lineDiffResults = DiffCellValue(srcLines, dstLines).ToList();

            var srcRange = new List<Tuple<string, Color?>>();
            var dstRange = new List<Tuple<string, Color?>>();
            foreach (var lineDiffResult in lineDiffResults)
            {
                if (lineDiffResult.Status == DiffStatus.Equal)
                {
                    DiffEqualLine(lineDiffResult, srcRange);
                    DiffEqualLine(lineDiffResult, dstRange);
                }
                else if (lineDiffResult.Status == DiffStatus.Modified)
                {
                    var charDiffResults = DiffUtil.Diff(lineDiffResult.Obj1, lineDiffResult.Obj2);
                    charDiffResults = DiffUtil.Order(charDiffResults, DiffOrderType.LazyDeleteFirst);
                    charDiffResults = DiffUtil.OptimizeCaseDeletedFirst(charDiffResults);

                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Inserted), srcRange, true);
                    DiffModifiedLine(charDiffResults.Where(r => r.Status != DiffStatus.Deleted), dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Deleted)
                {
                    DiffDeletedLine(lineDiffResult, srcRange, true);
                    DiffDeletedLine(lineDiffResult, dstRange, false);
                }
                else if (lineDiffResult.Status == DiffStatus.Inserted)
                {
                    DiffInsertedLine(lineDiffResult, srcRange, true);
                    DiffInsertedLine(lineDiffResult, dstRange, false);
                }
            }

            foreach (var r in srcRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                SrcValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }

            foreach (var r in dstRange)
            {
                var bc = r.Item2.HasValue ? new SolidColorBrush(r.Item2.Value) : new SolidColorBrush();
                DstValueTextBox.Document.Blocks.First().ContentStart.Paragraph.Inlines.Add(new Run(r.Item1) { Background = bc });
            }
            */
        }


        private void SetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderChagneEvent(args);
                }
            }
        }

        private void ResetRowHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchRowHeaderResetEvent(args);
                }
            }
        }

        private void SetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderChangeEvent(args);
                }
            }
        }

        private void ResetColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            var menuItem = sender as MenuItem;
            if (menuItem != null)
            {
                var dataGrid = ((ContextMenu)menuItem.Parent).PlacementTarget as FastGridControl;
                if (dataGrid != null)
                {
                    var args = new DiffViewEventArgs<FastGridControl>(dataGrid, container, TargetType.First);
                    DataGridEventDispatcher.Instance.DispatchColumnHeaderResetEvent(args);
                }
            }
        }

        private void DiffByHeaderSrc_Click(object sender, RoutedEventArgs e)
        {
            var headerIndex = MergeDataGrid.CurrentCell.Row.HasValue ? MergeDataGrid.CurrentCell.Row.Value : -1;

            diffConfig.SrcHeaderIndex = headerIndex;

        //    ExecuteDiff();
        }

        private void DiffByHeaderDst_Click(object sender, RoutedEventArgs e)
        {
       //     var headerIndex = DstDataGrid.CurrentCell.Row.HasValue ? DstDataGrid.CurrentCell.Row.Value : -1;

        //    diffConfig.DstSheetIndex = headerIndex;

       //     ExecuteDiff();
        }

        private void BuildCellBaseLog_Click(object sender, RoutedEventArgs e)
        {
         //   ShowLog();
        }

        private void CopyAsTsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells("\t");
        }

        private void CopyAsCsv_Click(object sender, RoutedEventArgs e)
        {
            CopyToClipboardSelectedCells(",");
        }

        private void CopyToClipboardSelectedCells(string separator)
        {
            if (copyTargetGrid == null)
                return;

            
            var model = copyTargetGrid.Model as DiffGridModel;
            if (model == null)
                return;
                
            var tsv = string.Join(Environment.NewLine,
               copyTargetGrid.SelectedCells
              .GroupBy(c => c.Row.Value)
              .OrderBy(g => g.Key)
              .Select(g => string.Join(separator, g.Select(c => model.GetCellText(c, true)))));

            Clipboard.SetDataObject(tsv);
        }


        private void InitCurrentCell()
        {
           if(MergeDataGrid.CurrentCell == null)
            {
                MergeDataGrid.CurrentCell = FastGridCellAddress.Zero;
            }

        }

        private bool ValidateDataGrids()
        {
            return MergeDataGrid.Model != null;
        }


    }
}
