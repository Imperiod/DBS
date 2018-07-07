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
    public partial class Microfilling : Window
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

        public Microfilling()
        {
            InitializeComponent();

            #region "Load entity"

            foreach (var item in Func.GetDB.Microfillings.Local.ToList())
            {
                switch (Func.GetDB.Entry(item).State)
                {
                    case EntityState.Detached:
                        break;
                    case EntityState.Unchanged:
                        break;
                    case EntityState.Added:
                        Func.GetDB.Microfillings.Local.Remove(item);
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

            Func.GetDB.Microfillings

                        .Include(i => i.Головний_розпорядник).Include(i => i.Змінив).Include(i => i.КДБ)
                        .Include(i => i.КЕКВ).Include(i => i.КФК).Include(i => i.Мікрофонд).Include(i => i.Створив)

                        .OrderBy(o=>o.Проведено).ThenBy(tb => tb.Мікрофонд).ThenBy(tb=>tb.КФК)
                        .ThenBy(tb => tb.Головний_розпорядник.Найменування).ThenBy(tb=>tb.КДБ).ThenBy(tb=>tb.КЕКВ)

                        .Load();

            #endregion

            ((CollectionViewSource)FindResource("cvs")).Source = Func.GetDB.Microfillings.Local;

            ((CollectionViewSource)FindResource("cvs")).Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Accept.Click += BTN_Accept_Click;
            BTN_Reset.Click += BTN_Reset_Click;
            BTN_ResetGroup.Click += BTN_ResetGroup_Click;
            BTN_Save.Click += BTN_Save_Click;

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
                List<MicroFilling> microFillings = new List<MicroFilling>();

                foreach (var item in DGM.SelectedCells)
                {
                    if (item.Item.ToString() != "{NewItemPlaceholder}" && microFillings.FirstOrDefault(f => f.Id == ((MicroFilling)item.Item).Id) is null)
                    {
                        microFillings.Add(((MicroFilling)item.Item));
                    }
                }

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                if (openFileDialog.ShowDialog() == true)
                {
                    PB.Minimum = 0;
                    PB.Maximum = microFillings.Count;
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
                        worksheet.Cells[r, 12] = "Мікрофонд";
                        worksheet.Cells[r, 13] = "КДБ";
                        worksheet.Cells[r, 14] = "КЕКВ";
                        worksheet.Cells[r, 15] = "Січень";
                        worksheet.Cells[r, 16] = "Лютий";
                        worksheet.Cells[r, 17] = "Березень";
                        worksheet.Cells[r, 18] = "Квітень";
                        worksheet.Cells[r, 19] = "Травень";
                        worksheet.Cells[r, 20] = "Червень";
                        worksheet.Cells[r, 21] = "Липень";
                        worksheet.Cells[r, 22] = "Серпень";
                        worksheet.Cells[r, 23] = "Вересень";
                        worksheet.Cells[r, 24] = "Жовтень";
                        worksheet.Cells[r, 25] = "Листопад";
                        worksheet.Cells[r, 26] = "Грудень";
                        r++;

                        worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[microFillings.Count, 26]], Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "Maestro_DataTable";

                        foreach (var item in microFillings)
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
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 11] = item.Мікрофонд.Фонд.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 12] = item.Мікрофонд.Повністю;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 13] = item.КДБ.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 14] = item.КЕКВ.Код;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 15] = item.Січень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 16] = item.Лютий;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 17] = item.Березень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 18] = item.Квітень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 19] = item.Травень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 20] = item.Червень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 21] = item.Липень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 22] = item.Серпень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 23] = item.Вересень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 24] = item.Жовтень;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 25] = item.Листопад;
                            worksheet.ListObjects["Maestro_DataTable"].Range.Cells[r, 26] = item.Грудень;
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
            if (Func.GetDB.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Microfilling == true) is null)
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
                        if (Func.GetDB.names_months.Contains(e.AddedCells[0].Column.Header.ToString()) &&
                            ((MicroFilling)e.AddedCells[0].Item).Головний_розпорядник != null &&
                            ((MicroFilling)e.AddedCells[0].Item).КДБ != null &&
                            ((MicroFilling)e.AddedCells[0].Item).КЕКВ != null &&
                            ((MicroFilling)e.AddedCells[0].Item).КФК != null &&
                            ((MicroFilling)e.AddedCells[0].Item).Мікрофонд != null)
                        {
                            #region "Fiels of Cell"
                            DateTime date = new DateTime();
                            date = ((MicroFilling)e.AddedCells[0].Item).Проведено;
                            var KFK = ((MicroFilling)e.AddedCells[0].Item).КФК;
                            var Main_manager = ((MicroFilling)e.AddedCells[0].Item).Головний_розпорядник;
                            var KDB = ((MicroFilling)e.AddedCells[0].Item).КДБ;
                            var KEKB = ((MicroFilling)e.AddedCells[0].Item).КЕКВ;
                            var MicroFond = ((MicroFilling)e.AddedCells[0].Item).Мікрофонд;
                            #endregion

                            #region "Current"
                            //Заполнение////////////////////////////////////////////////////////////////////////////////////

                            var qfill = Func.GetDB.Fillings
                                .Include(i => i.Головний_розпорядник)
                                .Include(i => i.КДБ)
                                .Include(i => i.КЕКВ)
                                .Include(i => i.КФК)
                                .Include(i => i.Фонд)
                                .Where(w => w.Видалено == false &&
                                            w.Головний_розпорядник.Id == Main_manager.Id &&
                                            w.КДБ.Id == KDB.Id &&
                                            w.КЕКВ.Id == KEKB.Id &&
                                            w.КФК.Id == KFK.Id &&
                                            w.Фонд.Id == MicroFond.Фонд.Id &&
                                            w.Проведено.Year == date.Year).ToList();

                            double fill = qfill.Select(s => (double)(s.GetType().GetProperty(e.AddedCells[0].Column.Header.ToString()).GetValue(s))).Sum();
                            ////////////////////////////////////////////////////////////////////////////////////////////////
                            //Мікрозаполнение///////////////////////////////////////////////////////////////////////////////
                            var qcurr = Func.GetDB.Microfillings.Local
                                .Where(w => w.Видалено == false &&
                                            w.Головний_розпорядник.Id == Main_manager.Id &&
                                            w.Проведено.Year == date.Year &&
                                            w.КДБ.Id == KDB.Id &&
                                            w.КЕКВ.Id == KEKB.Id &&
                                            w.КФК.Id == KFK.Id &&
                                            w.Мікрофонд.Фонд.Id == MicroFond.Фонд.Id).ToList();

                            double curr = qcurr.Select(s => (double)(s.GetType().GetProperty(e.AddedCells[0].Column.Header.ToString()).GetValue(s))).Sum();
                            ////////////////////////////////////////////////////////////////////////////////////////////////


                            GRPBCurr.Content = (fill - curr).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));

                            #endregion

                            #region "All"
                            //Заполнение////////////////////////////////////////////////////////////////////////////////////
                            qfill = Func.GetDB.Fillings
                                        .Include(i => i.Головний_розпорядник)
                                        .Include(i => i.КФК)
                                        .Include(i => i.Фонд)
                                    .Where(w => w.Видалено == false &&
                                                w.Головний_розпорядник.Id == Main_manager.Id &&
                                                w.КФК.Id == KFK.Id &&
                                                w.Фонд.Id == MicroFond.Фонд.Id &&
                                                w.Проведено.Year == date.Year).ToList();

                            fill = qfill.Select(s => (double)(s.GetType().GetProperty(e.AddedCells[0].Column.Header.ToString()).GetValue(s))).Sum();
                            ////////////////////////////////////////////////////////////////////////////////////////////////
                            //Мікрозаполнение///////////////////////////////////////////////////////////////////////////////
                            qcurr = Func.GetDB.Microfillings.Local
                                .Where(w => w.Видалено == false &&
                                            w.Головний_розпорядник.Id == Main_manager.Id &&
                                            w.Проведено.Year == date.Year &&
                                            w.КФК.Id == KFK.Id &&
                                            w.Мікрофонд.Фонд.Id == MicroFond.Фонд.Id).ToList();
                            curr = qcurr.Select(s => (double)(s.GetType().GetProperty(e.AddedCells[0].Column.Header.ToString()).GetValue(s))).Sum();
                            ////////////////////////////////////////////////////////////////////////////////////////////////

                            GRPBAll.Content = (fill - curr).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));

                            #endregion

                            double d;
                            double.TryParse(DGM.SelectedCells.First().Item.GetType().GetProperty(DGM.SelectedCells.FirstOrDefault().Column.Header.ToString()).GetValue(DGM.SelectedCells.First().Item).ToString(), out d);

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
                            if (double.TryParse(item.Item.GetType().GetProperty(item.Column.Header.ToString()).GetValue(item.Item)?.ToString(), out d))
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
                        GRPBAll.Content = "Не более 1 ячейки.";
                        GRPBCurr.Content = "Не более 1 ячейки.";
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

        private void DGM_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            Func.GenerateColumnForDataGrid(ref counterForDGMColumns, e);
        }

        private void IsAllowEditing(DataGridBeginningEditEventArgs e)
        {
            if (((MicroFilling)e.Row.Item).Id != 0 && ((MicroFilling)e.Row.Item).Підписано)
            {
                if (Func.Login == "LeXX")
                {
                    ((MicroFilling)e.Row.Item).Підписано = false;
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
                    if (((MicroFilling)e.Row.Item).Змінив.Пароль == pass)
                    {
                        ((MicroFilling)e.Row.Item).Підписано = false;
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

        private void DGM_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (((MicroFilling)e.Row.Item).Id == 0)
            {
                ((MicroFilling)e.Row.Item).Створив = Func.GetDB.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            }
            ((MicroFilling)e.Row.Item).Змінив = Func.GetDB.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            ((MicroFilling)e.Row.Item).Змінено = DateTime.Now;
        }
    }
    #region "Converters"

    public class MicroFillingGroupTotalConverter : IValueConverter
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
                                select ((MicroFilling)x).Січень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Лютий).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Березень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Квітень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Травень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Червень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Липень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Серпень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Вересень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Жовтень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Листопад).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((MicroFilling)x).Грудень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Січень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Лютий).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Березень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Квітень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Травень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Червень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Липень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Серпень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Вересень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Жовтень).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Листопад).Sum();
                sum += (from x in items
                        select ((MicroFilling)x).Грудень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupOneConverter : IValueConverter
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
                                select ((MicroFilling)x).Січень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Січень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupTwoConverter : IValueConverter
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
                                select ((MicroFilling)x).Лютий).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Лютий).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupThreeConverter : IValueConverter
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
                                select ((MicroFilling)x).Березень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Березень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupFourConverter : IValueConverter
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
                                select ((MicroFilling)x).Квітень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Квітень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupFiveConverter : IValueConverter
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
                                select ((MicroFilling)x).Травень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Травень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupSixConverter : IValueConverter
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
                                select ((MicroFilling)x).Червень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Червень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupSevenConverter : IValueConverter
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
                                select ((MicroFilling)x).Липень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Липень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupEightConverter : IValueConverter
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
                                select ((MicroFilling)x).Серпень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Серпень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupNineConverter : IValueConverter
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
                                select ((MicroFilling)x).Вересень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Вересень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupTenConverter : IValueConverter
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
                                select ((MicroFilling)x).Жовтень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Жовтень).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupElevenConverter : IValueConverter
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
                                select ((MicroFilling)x).Листопад).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Листопад).Sum();
                return sum;
            }
        }
    }

    public class MicroFillingGroupTwelveConverter : IValueConverter
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
                                select ((MicroFilling)x).Грудень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((MicroFilling)x).Грудень).Sum();
                return sum;
            }
        }
    }

    #endregion
}
