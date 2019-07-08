using DBSolom;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
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
    public partial class Financing : Window
    {

        #region "Variables"

        string currentColumnNameForD_D { get; set; }
        DataGridCell DataGridCell { get; set; }
        public string type = "";
        public string prop = "";
        public object value = null;
        public int counterForDGMColumns = 0;

        public List<Label> GetLabels = new List<Label>();
        public Dictionary<string, ComboBox> dict_cmb = new Dictionary<string, ComboBox>();
        public Dictionary<string, TextBox> dict_txt = new Dictionary<string, TextBox>();
        List<Filters> GetFilters = new List<Filters>();
        public List<ToggleButton> CheckBoxes = new List<ToggleButton>();
        bool IsInitialization = true;
        CollectionViewSource CollectionViewSource { get; set; }

        DBSolom.Db db = new Db(Func.GetConnectionString);

        #endregion

        public Financing()
        {
            InitializeComponent();

            CollectionViewSource = ((CollectionViewSource)FindResource("cvs"));

            CollectionViewSource.Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Accept.Click += BTN_Accept_Click;
            BTN_Reset.Click += BTN_Reset_Click;
            BTN_ResetGroup.Click += BTN_ResetGroup_Click;
            BTN_Save.Click += BTN_Save_Click;
            BTN_ExportToExcel.Click += Func.BTN_ExportToExcel_Click;
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
                db.SaveChanges();
                MessageBox.Show("Зміни збережено!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            DGM.Items.Refresh();
        }
        #endregion

        private void DGM_Loaded(object sender, RoutedEventArgs e)
        {
            if (db.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Financing == true) is null)
            {
                DGM.IsReadOnly = true;
            }
        }

        private void DGM_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                if (DGM.SelectedCells.Count > 0 && EXPCALC.IsExpanded)
                {
                    if (DGM.SelectedCells.Count == 1)
                    {
                        if (DGM.CurrentColumn.Header.ToString() == "Сума" &&
                            ((DBSolom.Financing)e.AddedCells[0].Item).Проведено != null &&
                            ((DBSolom.Financing)e.AddedCells[0].Item).Головний_розпорядник != null &&
                            ((DBSolom.Financing)e.AddedCells[0].Item).КЕКВ != null &&
                            ((DBSolom.Financing)e.AddedCells[0].Item).КФК != null &&
                            ((DBSolom.Financing)e.AddedCells[0].Item).Мікрофонд != null)
                        {
                            List<string> vs = Func.names_months;
                            double yearSumFilling = 0;
                            double yearSumCorrection = 0;
                            double yearSumFinancing = 0;
                            double periodSumFilling = 0;
                            double periodSumCorrection = 0;
                            double periodSumFinancing = 0;

                            #region "Fiels of Cell"

                            DateTime date = ((DBSolom.Financing)e.AddedCells[0].Item).Проведено;
                            KFK KFK = ((DBSolom.Financing)e.AddedCells[0].Item).КФК;
                            Main_manager Main_manager = ((DBSolom.Financing)e.AddedCells[0].Item).Головний_розпорядник;
                            KEKB KEKB = ((DBSolom.Financing)e.AddedCells[0].Item).КЕКВ;
                            Foundation FOND = ((DBSolom.Financing)e.AddedCells[0].Item).Мікрофонд.Фонд;
                            MicroFoundation MicroFond = ((DBSolom.Financing)e.AddedCells[0].Item).Мікрофонд;

                            #endregion

                            DBSolom.Db mdb = new Db(Func.GetConnectionString);

                            #region "Foundation"

                            #region "Filling"

                            List<DBSolom.Filling> qfil = mdb.Fillings
                                            .Include(i => i.Головний_розпорядник)
                                            .Include(i => i.КФБ)
                                            .Include(i => i.КДБ)
                                            .Include(i => i.КЕКВ)
                                            .Include(i => i.КФК)
                                            .Include(i => i.Фонд)
                                            .Where(w => w.Видалено == false &&
                                                        w.Головний_розпорядник.Найменування == Main_manager.Найменування &&
                                                        w.Проведено.Year == date.Year &&
                                                        w.КЕКВ.Код == KEKB.Код &&
                                                        w.КФК.Код == KFK.Код &&
                                                        w.Фонд.Код == FOND.Код).ToList();

                            for (int j = 0; j < date.Month; j++)
                            {
                                periodSumFilling += qfil.Select(s => (double)s.GetType().GetProperty(vs[j]).GetValue(s)).Sum();
                            }
                            for (int j = 0; j < 11; j++)
                            {
                                yearSumFilling += qfil.Select(s => (double)s.GetType().GetProperty(vs[j]).GetValue(s)).Sum();
                            }

                            #endregion

                            #region "Correction"

                            List<DBSolom.Correction> qcorr = mdb.Corrections
                                                .Include(i => i.Головний_розпорядник)
                                                .Include(i => i.КФБ)
                                                .Include(i => i.КДБ)
                                                .Include(i => i.КЕКВ)
                                                .Include(i => i.КФК)
                                                .Include(i => i.Мікрофонд)
                                        .Where(w => w.Видалено == false &&
                                                    w.Головний_розпорядник.Найменування == Main_manager.Найменування &&
                                                    w.Проведено.Year == date.Year &&
                                                    w.КЕКВ.Код == KEKB.Код &&
                                                    w.КФК.Код == KFK.Код &&
                                                    w.Мікрофонд.Фонд.Код == FOND.Код).ToList();

                            for (int j = 0; j < date.Month; j++)
                            {
                                periodSumCorrection += qcorr.Select(s => (double)s.GetType().GetProperty(vs[j]).GetValue(s)).Sum();
                            }
                            for (int j = 0; j < 11; j++)
                            {
                                yearSumCorrection += qcorr.Select(s => (double)s.GetType().GetProperty(vs[j]).GetValue(s)).Sum();
                            }

                            #endregion

                            #region "Financing"

                            List<DBSolom.Financing> qfin = mdb.Financings.Where(w => w.Видалено == false &&
                                                            w.Головний_розпорядник.Найменування == Main_manager.Найменування &&
                                                            w.Проведено.Year == date.Year &&
                                                            w.КЕКВ.Код == KEKB.Код &&
                                                            w.КФК.Код == KFK.Код &&
                                                            w.Мікрофонд.Фонд.Код == FOND.Код).ToList();

                            db.Financings.Local.Where(w => w.Видалено == false &&
                                                            w.Головний_розпорядник.Найменування == Main_manager.Найменування &&
                                                            w.Проведено.Year == date.Year &&
                                                            w.КЕКВ.Код == KEKB.Код &&
                                                            w.КФК.Код == KFK.Код &&
                                                            w.Мікрофонд.Фонд.Код == FOND.Код)
                                                            .ToList()
                                                            .ForEach(item =>
                                                            {
                                                                if (db.Entry(item).State != EntityState.Unchanged)
                                                                {
                                                                    qfin.Add(item);
                                                                }
                                                            });

                            yearSumFinancing = qfin.Select(s => s.Сума).Sum();
                            periodSumFinancing = qfin.Where(w => w.Проведено <= date).Select(s => s.Сума).Sum();

                            #endregion

                            GRPBYearFond.Content = (yearSumFilling + yearSumCorrection - yearSumFinancing).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                            GRPBPeriodFond.Content = (periodSumFilling + periodSumCorrection - periodSumFinancing).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
                            #endregion
                        }


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

        private void DGM_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (((DBSolom.Financing)e.Row.Item).Id == 0)
            {
                ((DBSolom.Financing)e.Row.Item).Створив = db.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            }
            ((DBSolom.Financing)e.Row.Item).Змінив = db.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            ((DBSolom.Financing)e.Row.Item).Змінено = DateTime.Now;
            if (e.EditAction != DataGridEditAction.Cancel)
            {
                if (((DBSolom.Financing)e.EditingElement.DataContext).Головний_розпорядник != null &&
                    ((DBSolom.Financing)e.EditingElement.DataContext).КЕКВ != null &&
                    ((DBSolom.Financing)e.EditingElement.DataContext).КФК != null &&
                    ((DBSolom.Financing)e.EditingElement.DataContext).Мікрофонд != null)
                {
                    #region "Fields of Cell"
                    DateTime date = ((DBSolom.Financing)e.Row.Item).Проведено;
                    KFK KFK = ((DBSolom.Financing)e.Row.Item).КФК;
                    Main_manager Main_manager = ((DBSolom.Financing)e.Row.Item).Головний_розпорядник;
                    KEKB KEKB = ((DBSolom.Financing)e.Row.Item).КЕКВ;
                    Foundation FOND = ((DBSolom.Financing)e.Row.Item).Мікрофонд.Фонд;
                    MicroFoundation MicroFond = ((DBSolom.Financing)e.Row.Item).Мікрофонд;
                    #endregion

                    var x = Func.GetCurrentPlanAndRemainderFromDBPerMonth(db, date.Year, KFK, Main_manager, KEKB, FOND);

                    if ((x[TypeOfFinanceData.Remainders][date.Month - 1] < 0))
                    {
                        DGM.CancelEdit(DataGridEditingUnit.Cell);

                        MessageBox.Show(($"[Дата: {date.ToShortDateString()}] [Фонд: {FOND.Код}] [КПБ: {KFK.Код}]" +
                            $" [Головний розпорядник: {Main_manager.Найменування}] [КЕКВ: {KEKB.Код}]" +
                            $" [Місяць: {Func.names_months[date.Month - 1]}] [Остаток:{x[TypeOfFinanceData.Remainders][date.Month - 1]}]"),
                            "Maestro", MessageBoxButton.OK, MessageBoxImage.Error);

                        return;
                    }
                }
            }
            else
            {
                ((DBSolom.Financing)e.Row.Item).Сума = 0;
            }
        }

        private void DGM_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            IsAllowEditing(e);
        }

        private void IsAllowEditing(DataGridBeginningEditEventArgs e)
        {
            if (((DBSolom.Financing)e.Row.Item).Id != 0 && ((DBSolom.Financing)e.Row.Item).Підписано)
            {
                if (Func.Login == "LeXX" || ((DBSolom.Financing)e.Row.Item).Змінив.Логін == Func.Login)
                {
                    ((DBSolom.Financing)e.Row.Item).Підписано = false;
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
                    if (((DBSolom.Financing)e.Row.Item).Змінив.Пароль == pass)
                    {
                        ((DBSolom.Financing)e.Row.Item).Підписано = false;

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
            Func.GenerateColumnForDataGrid(db, ref counterForDGMColumns, e);
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SDate.SelectedDate != null && EDate.SelectedDate != null)
            {
                #region "Load entity"

                DateTime FinalDate = (EDate.SelectedDate.Value.AddDays(1) - TimeSpan.FromSeconds(1));

                db.Financings
                            .Include(i => i.Головний_розпорядник)
                            .Include(i => i.Змінив)
                            .Include(i => i.КЕКВ)
                            .Include(i => i.КФК)
                            .Include(i => i.Мікрофонд)
                            .Include(i => i.Створив)
                            .Where(w => w.Проведено >= SDate.SelectedDate && w.Проведено <= FinalDate)
                            .Load();

                #endregion

                if (IsInitialization)
                {
                    CollectionViewSource.Source = db.Financings.Local;

                    DGM.ItemsSource = CollectionViewSource.View;

                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Проведено", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Мікрофонд.Повністю", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("КФК.Код", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Головний_розпорядник.Найменування", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("КЕКВ.Код", ListSortDirection.Ascending));

                    int t = 0;
                    foreach (ItemPropertyInfo item in ((IItemProperties)DGM.Items).ItemProperties)
                    {
                        Func.GetFilters(EXPGRO, t, item, ref dict_cmb, ref dict_txt, ref GetLabels);

                        Func.GetGroups(t, item, ref CheckBoxes, ref EXPGRT);

                        Func.GetVisibilityOfColumns(t, item, ref EXPHDN);

                        t++;
                    }

                    IsInitialization = false;
                }

                CollectionViewSource.GetDefaultView(DGM.ItemsSource).Refresh();
            }
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F9)
            {
                CopyEntityInTable();
            }
        }

        private void CopyEntityInTable()
        {
            if (DGM.CurrentItem is null || DGM.SelectedCells?.Count > 1)
            {
                MessageBox.Show("Встаньте на 1 ячейку необходимой строки!", "Maestro: [Коригування]", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else
            {
                DBSolom.Financing row = (DBSolom.Financing)DGM.CurrentItem;
                if (row.Головний_розпорядник is null || row.КЕКВ is null || row.КФК is null || row.Мікрофонд is null)
                {
                    MessageBox.Show("Заполните все данные в строке!", "Maestro: [Коригування]", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
                else
                {
                    DBSolom.Financing financing = new DBSolom.Financing()
                    {
                        Головний_розпорядник = row.Головний_розпорядник,
                        КФК = row.КФК,
                        КЕКВ = row.КЕКВ,
                        Мікрофонд = row.Мікрофонд,
                        Змінив = db.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login),
                        Створив = db.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login)
                    };
                    db.Financings.Local.Add(financing);

                    DGM.CommitEdit();
                    DGM.CommitEdit();

                    CollectionViewSource.View.Refresh();
                    DGM.CurrentItem = financing;
                    DGM.CurrentCell = new DataGridCellInfo(financing, DGM.Columns.First(f => f.Header.ToString() == "Головний_розпорядник"));
                    DGM.BeginEdit();
                }
            }
        }

        private void BTN_Copy_Click(object sender, RoutedEventArgs e)
        {
            CopyEntityInTable();
        }
    }

    public class FinancingGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Сума: " + CheckedFillingItems(items).ToString("N2",
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
                                select ((DBSolom.Financing)x).Сума).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Financing)x).Сума).Sum();
                return sum;
            }
        }
    }
}
