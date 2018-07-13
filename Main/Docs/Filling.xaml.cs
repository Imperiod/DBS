using DBSolom;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Main.Docs
{
    public partial class Filling : Window
    {
        #region "Variables"

        public string type = "";
        public string prop = "";
        public object value = null;
        public int counterForDGMColumns = 0;

        public List<Label> GetLabels = new List<Label>();
        public Dictionary<string, ComboBox> dict_cmb = new Dictionary<string, ComboBox>();
        public Dictionary<string, TextBox> dict_txt = new Dictionary<string, TextBox>();
        List<Filters> GetFilters = new List<Filters>();
        public List<ToggleButton> CheckBoxes = new List<ToggleButton>();

        #endregion

        public Filling()
        {
            InitializeComponent();

            #region "Load entity"

            foreach (var item in Func.GetDB.Fillings.Local.ToList())
            {
                switch (Func.GetDB.Entry(item).State)
                {
                    case EntityState.Detached:
                        break;
                    case EntityState.Unchanged:
                        break;
                    case EntityState.Added:
                        Func.GetDB.Fillings.Local.Remove(item);
                        break;
                    case EntityState.Deleted:
                        break;
                    case EntityState.Modified:
                        Func.GetDB.Entry(item).Reload();
                        break;
                    default:
                        break;
                }
            }

            Func.GetDB.Fillings

                        .Include(i => i.Головний_розпорядник).Include(i => i.Змінив).Include(i => i.КДБ)
                        .Include(i => i.КЕКВ).Include(i => i.КФК).Include(i => i.Створив).Include(i => i.Фонд)

                        .OrderBy(o=>o.Проведено).ThenBy(tb=>tb.Фонд).ThenBy(tb=>tb.КФК)
                        .ThenBy(tb=>tb.Головний_розпорядник).ThenBy(tb=>tb.КДБ).ThenBy(tb=>tb.КЕКВ)

                        .Load();

            #endregion

            ((CollectionViewSource)FindResource("cvs")).Source = Func.GetDB.Fillings.Local;

            ((CollectionViewSource)FindResource("cvs")).Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Accept.Click += BTN_Accept_Click;
            BTN_Reset.Click += BTN_Reset_Click;
            BTN_ResetGroup.Click += BTN_ResetGroup_Click;
            BTN_Save.Click += BTN_Save_Click;
            BTN_ExportToExcel.Click += BTN_ExportToExcel_Click;

            EVAL.MouseEnter += Func.Expander_MouseEnter;
            EVAL.MouseLeave += Func.Expander_MouseLeave;
            EXPMAESTRO.MouseEnter += Func.Expander_MouseEnter;
            EXPMAESTRO.MouseLeave += Func.Expander_MouseLeave;

            var t = 0;
            foreach (var item in ((IItemProperties)DGM.Items).ItemProperties)
            {
                Func.GetFilters(EXPGRO, t, item, ref dict_cmb, ref dict_txt, ref GetLabels);

                Func.GetGroups(t, item, ref CheckBoxes, ref EXPGRT);

                Func.GetVisibilityOfColumns(t, item, ref EXPHDN);

                t++;
            }
        }

        #region "BUTTONS"

        public void BTN_Accept_Click(object sender, RoutedEventArgs e)
        {
            Filters filters = new Filters();
            string str = "";
            bool first = true;

            for (int i = 0; i < GetLabels.Count; i++)
            {
                if (dict_txt[GetLabels[i].Content.ToString()].Text != "")
                {
                    type = dict_cmb[GetLabels[i].Content.ToString()].Text;
                    prop = GetLabels[i].Content.ToString();
                    value = dict_txt[GetLabels[i].Content.ToString()].Text;
                    filters.GetFilters.Add(new Dictionary<string, dynamic>() { { "type", type }, { "prop", prop }, { "value", value } });

                    str += first ? prop + " " + type + " " + value : " & " + prop + " " + type + " " + value;
                    first = false;
                }
            }

            LBFilters.Items.Add(str);

            for (int i = 0; i < dict_cmb.Count; i++)
            {
                dict_cmb.Select(s => s.Value).ToList()[i].SelectedValue = null;
                dict_txt.Select(s => s.Value).ToList()[i].Text = null;
            }
            type = "";
            prop = "";
            value = null;

            GetFilters.Add(filters);

            CollectionViewSource.GetDefaultView(DGM.ItemsSource).Refresh();
        }
        public void BTN_Reset_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < dict_cmb.Count; i++)
            {
                dict_cmb.Select(s => s.Value).ToList()[i].SelectedValue = null;
                dict_txt.Select(s => s.Value).ToList()[i].Text = null;
            }
            type = "";
            prop = "";
            value = null;

            LBFilters.Items.Clear();
            GetFilters.Clear();
            CollectionViewSource.GetDefaultView(DGM.ItemsSource).Refresh();
        }
        public void BTN_ResetGroup_Click(object sender, RoutedEventArgs e)
        {
            ICollectionView cvTasks = CollectionViewSource.GetDefaultView(DGM.ItemsSource);
            CheckBoxes.ForEach(a => a.IsChecked = false);
            cvTasks.GroupDescriptions.Clear();
        }
        public void BTN_Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Func.GetDB.SaveChanges();
                MessageBox.Show("Зміни збережено!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            DGM.Items.Refresh();
        }
        private void BTN_ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            bool IsExist = false;
            if (DGM.SelectedCells.Count > 0)
            {
                List<DBSolom.Filling> fillings = new List<DBSolom.Filling>();

                foreach (var item in DGM.SelectedCells)
                {
                    
                    if (item.Item.ToString() != "{NewItemPlaceholder}" && fillings.FirstOrDefault(f => f.Id == ((DBSolom.Filling)item.Item).Id) is null)
                    {
                        fillings.Add(((DBSolom.Filling)item.Item));
                    }
                }

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                if (openFileDialog.ShowDialog() == true)
                {
                    PB.Minimum = 0;
                    PB.Maximum = fillings.Count;
                    PB.Value = 1;

                    Action action = () => { PB.Value++; };
                    var Task = new Task(() =>
                    {
                        Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                        application.AskToUpdateLinks = false;
                        application.DisplayAlerts = false;
                        Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Open(openFileDialog.FileName);
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

                        foreach (Microsoft.Office.Interop.Excel.Worksheet item in workbook.Worksheets)
                        {
                            if (item.Name == "Maestro_Data")
                            {
                                IsExist = true;
                                worksheet = item;
                                break;
                            }
                        }

                        if (IsExist)
                        {
                            worksheet.Cells.Clear();
                            if (worksheet.ListObjects.Count != 0)
                            {
                                for (int i = 0; i < worksheet.ListObjects.Count; i++)
                                {
                                    if (worksheet.ListObjects[i].Name == "Maestro_DataTable")
                                    {
                                        worksheet.ListObjects.Item[i].Delete();
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            worksheet = workbook.Worksheets.Add();
                            worksheet.Name = "Maestro_Data";
                        }

                        int r = 1;
                        worksheet.Cells[r, 1] = "Id";
                        worksheet.Cells[r, 2] = "Видалено";
                        worksheet.Cells[r, 3] = "Створив";
                        worksheet.Cells[r, 4] = "Створино";
                        worksheet.Cells[r, 5] = "Змінив";
                        worksheet.Cells[r, 6] = "Змінено";
                        worksheet.Cells[r, 7] = "Проведено";
                        worksheet.Cells[r, 8] = "Підписано";
                        worksheet.Cells[r, 9] = "Головний_розпорядник";
                        worksheet.Cells[r, 10] = "КФК";
                        worksheet.Cells[r, 11] = "Фонд";
                        worksheet.Cells[r, 12] = "КДБ";
                        worksheet.Cells[r, 13] = "КЕКВ";
                        worksheet.Cells[r, 14] = "Січень";
                        worksheet.Cells[r, 15] = "Лютий";
                        worksheet.Cells[r, 16] = "Березень";
                        worksheet.Cells[r, 17] = "Квітень";
                        worksheet.Cells[r, 18] = "Травень";
                        worksheet.Cells[r, 19] = "Червень";
                        worksheet.Cells[r, 20] = "Липень";
                        worksheet.Cells[r, 21] = "Серпень";
                        worksheet.Cells[r, 22] = "Вересень";
                        worksheet.Cells[r, 23] = "Жовтень";
                        worksheet.Cells[r, 24] = "Листопад";
                        worksheet.Cells[r, 25] = "Грудень";
                        r++;

                        worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[fillings.Count, 25]], Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "Maestro_DataTable";

                        foreach (var item in fillings)
                        {
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 1] = item.Id;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 2] = item.Видалено;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 3] = item.Створив.Логін;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 4] = item.Створино.ToShortDateString();
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 5] = item.Змінив.Логін;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 6] = item.Змінено.ToShortDateString();
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 7] = item.Проведено.ToShortDateString();
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 8] = item.Підписано;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 9] = item.Головний_розпорядник.Найменування;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 10] = item.КФК.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 11] = item.Фонд.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 12] = item.КДБ.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 13] = item.КЕКВ.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 14] = item.Січень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 15] = item.Лютий;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 16] = item.Березень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 17] = item.Квітень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 18] = item.Травень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 19] = item.Червень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 20] = item.Липень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 21] = item.Серпень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 22] = item.Вересень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 23] = item.Жовтень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 24] = item.Листопад;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 25] = item.Грудень;
                            PB.Dispatcher.Invoke(action);
                            r++;
                        }

                        MessageBox.Show("Виконано!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Information);
                        application.Visible = true;
                        openFileDialog = null;
                        application = null;
                        workbook = null;
                        worksheet = null;
                        PB.Dispatcher.Invoke(() => PB.Value = 0);
                    });

                    Task.Start();
                }
            }
            else
            {
                MessageBox.Show("Виділіть всі строки, які будуть експортовані!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }
        #endregion

        private void DGM_Loaded(object sender, RoutedEventArgs e)
        {
            if (Func.GetDB.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Filling == true) is null)
            {
                DGM.IsReadOnly = true;
            }
        }

        private void DGM_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                if (DGM.SelectedCells.Count > 0)
                {
                    if (DGM.SelectedCells.Count == 1)
                    {
                        double d;
                        if (double.TryParse(DGM.SelectedCells.First().Item.GetType().GetProperty(DGM.SelectedCells.FirstOrDefault().Column.Header.ToString()).GetValue(DGM.SelectedCells.First().Item).ToString(), out d))
                        {
                            GRPBElm.Content = "1";
                            GRPBSum.Content = d.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                            GRPBSred.Content = "";
                            GRPBMin.Content = "";
                            GRPBMax.Content = "";
                        }
                    }
                    else
                    {
                        double sum = 0;
                        int counter = 0;
                        double min = double.MaxValue;
                        double max = 0;
                        foreach (var item in DGM.SelectedCells)
                        {
                            double d;
                            if (double.TryParse(item.Item.GetType().GetProperty(item.Column.Header.ToString()).GetValue(item.Item).ToString(), out d))
                            {
                                if (d > max)
                                {
                                    max = d;
                                }
                                if (d < min)
                                {
                                    min = d;
                                }
                                counter++;
                                sum += d;
                            }
                        }
                        GRPBElm.Content = counter.ToString("N0", CultureInfo.CreateSpecificCulture("ru-RU"));
                        GRPBSum.Content = sum.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                        GRPBSred.Content = (sum == 0 ? 0 : sum / counter).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                        GRPBMin.Content = min.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                        GRPBMax.Content = max.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private void DGM_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            IsAllowEditing(e);
        }

        private void IsAllowEditing(DataGridBeginningEditEventArgs e)
        {
            if (((DBSolom.Filling)e.Row.Item).Id != 0 && ((DBSolom.Filling)e.Row.Item).Підписано)
            {
                if (Func.Login == "LeXX" || ((DBSolom.Correction)e.Row.Item).Змінив.Логін == Func.Login)
                {
                    ((DBSolom.Filling)e.Row.Item).Підписано = false;
                    var cellContent = DGM.Columns.First(f => f.Header.ToString() == "Підписано").GetCellContent(e.Row);
                    if (cellContent is CheckBox)
                    {
                        ((CheckBox)cellContent).IsChecked = false;
                    }
                    e.Cancel = false;
                }
                else if (MessageBox.Show("Ви маєте пароль для редагування?\n\t(пр. пароль це пароль користувача який вніс зміни)", "Редагування заблоковано", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    string pass = Microsoft.VisualBasic.Interaction.InputBox("Пароль: ", "");
                    if (((DBSolom.Filling)e.Row.Item).Змінив.Пароль == pass)
                    {
                        ((DBSolom.Filling)e.Row.Item).Підписано = false;
                        var cellContent = DGM.Columns.First(f => f.Header.ToString() == "Підписано").GetCellContent(e.Row);
                        if (cellContent is CheckBox)
                        {
                            ((CheckBox)cellContent).IsChecked = false;
                        }
                        e.Cancel = false;
                    }
                    else
                    {
                        e.Cancel = true;
                    }
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }

        private void DGM_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            Func.GenerateColumnForDataGrid(ref counterForDGMColumns, e);
        }

        private void DGM_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (((DBSolom.Filling)e.Row.Item).Id == 0)
            {
                ((DBSolom.Filling)e.Row.Item).Створив = Func.GetDB.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            }
            ((DBSolom.Filling)e.Row.Item).Змінив = Func.GetDB.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            ((DBSolom.Filling)e.Row.Item).Змінено = DateTime.Now;
        }
    }

    #region "Converters"

    public class FillingGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Total: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Січень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Лютий).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Березень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Квітень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Травень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Червень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Липень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Серпень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Вересень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Жовтень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Листопад).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Грудень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Січень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Лютий).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Березень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Квітень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Травень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Червень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Липень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Серпень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Вересень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Жовтень).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Листопад).Sum();
                sum += (from x in items
                        select ((DBSolom.Filling)x).Грудень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Січень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Січень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Січень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupTwoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Лютий: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Лютий).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Лютий).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupThreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Березень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Березень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Березень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupFourConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Квітень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Квітень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Квітень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupFiveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Травень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Травень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Травень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupSixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Червень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Червень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Червень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupSevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Липень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Липень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Липень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupEightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Серпень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Серпень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Серпень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupNineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Вересень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Вересень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Вересень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupTenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Жовтень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Жовтень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Жовтень).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupElevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Листопад: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Листопад).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Листопад).Sum();
                return sum;
            }
        }
    }

    public class FillingGroupTwelveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Грудень: " + CheckedFillingItems(items).ToString("N2",
                  CultureInfo.CreateSpecificCulture("ru-RU"));
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
        private static double CheckedFillingItems(ReadOnlyObservableCollection<object> items)
        {
            var delta = items.FirstOrDefault(s => s.GetType().GetProperties().Select(p => p.Name).ToList().Contains("Items"));
            double sum = 0;
            if (delta != null)
            {
                long counter = items.Count;
                for (int i = 0; i < items.Count; i++)
                {
                    if (((CollectionViewGroup)items[i]).Items.FirstOrDefault(f => f.GetType().GetProperties().Select(s => s.Name).ToList().Contains("Items")) != null)
                    {
                        sum += CheckedFillingItems(((CollectionViewGroup)items[i]).Items);
                    }
                    else
                    {
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Filling)x).Грудень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Filling)x).Грудень).Sum();
                return sum;
            }
        }
    }

    #endregion
}
