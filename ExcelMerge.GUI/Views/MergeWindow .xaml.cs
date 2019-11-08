using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ExcelMerge.GUI.Models;
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

        private const string srcKey = "src";

        private DiffView _diffView = null;

        public MergeWindow(DiffView diffView)
        {
            InitializeComponent();
            InitializeContainer();

            if(diffView == null)
            {
                this.Close();
            }

            _diffView = diffView;
        }

        private void InitializeContainer()
        {
            container = new UnityContainer();
            container
                .RegisterInstance(srcKey, SrcDataGrid);
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

        private void DataGrid_SelectedCellsChanged(object sender, FastWpfGrid.SelectionChangedEventArgs e)
        {
            var grid = copyTargetGrid = sender as FastGridControl;
            if (grid == null)
                return;

            copyTargetGrid = grid;

            var args = new DiffViewEventArgs<FastGridControl>(sender as FastGridControl, container);
            DataGridEventDispatcher.Instance.DispatchSelectedCellChangeEvent(args);

            /*
            if (!SrcDataGrid.CurrentCell.Row.HasValue || !DstDataGrid.CurrentCell.Row.HasValue)
                return;

            if (!SrcDataGrid.CurrentCell.Column.HasValue || !DstDataGrid.CurrentCell.Column.HasValue)
                return;

            if (SrcDataGrid.Model == null || DstDataGrid.Model == null)
                return;

            var srcValue =
                (SrcDataGrid.Model as DiffGridModel).GetCellText(SrcDataGrid.CurrentCell.Row.Value, SrcDataGrid.CurrentCell.Column.Value, true);
            var dstValue =
                (DstDataGrid.Model as DiffGridModel).GetCellText(DstDataGrid.CurrentCell.Row.Value, DstDataGrid.CurrentCell.Column.Value, true);

            UpdateValueDiff(srcValue, dstValue);
            */
            if (App.Instance.Setting.AlwaysExpandCellDiff)
            {
                var a = new DiffViewEventArgs<RichTextBox>(null, container, TargetType.First);
                ValueTextBoxEventDispatcher.Instance.DispatchGotFocusEvent(a);
            }
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
            var headerIndex = SrcDataGrid.CurrentCell.Row.HasValue ? SrcDataGrid.CurrentCell.Row.Value : -1;

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


    }
}
