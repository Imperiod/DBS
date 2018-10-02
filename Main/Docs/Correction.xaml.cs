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
    public partial class Correction : Window
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

        public Correction()
        {
            InitializeComponent();

            #region "Load entity"

            //foreach (var item in Func.GetDB.Fillings.Local)
            //{
            //    switch (Func.GetDB.Entry(item).State)
            //    {
            //        case EntityState.Detached:
            //            break;
            //        case EntityState.Unchanged:
            //            break;
            //        case EntityState.Added:
            //            Func.GetDB.Fillings.Remove(item);
            //            break;
            //        case EntityState.Deleted:
            //            break;
            //        case EntityState.Modified:
            //            Func.GetDB.Entry(item).Reload();
            //            break;
            //        default:
            //            break;
            //    }
            //}

            //Func.GetDB.Fillings

            //        .Include(i => i.Головний_розпорядник).Include(i => i.Змінив).Include(i => i.КДБ)
            //        .Include(i => i.КЕКВ).Include(i => i.КФК).Include(i => i.Створив).Include(i => i.Фонд.Макрофонд)

            //        .OrderBy(o => o.Проведено).ThenBy(tb => tb.Фонд).ThenBy(tb => tb.КФК)
            //        .ThenBy(tb => tb.Головний_розпорядник).ThenBy(tb => tb.КДБ).ThenBy(tb => tb.КЕКВ)

            //        .Load();

            foreach (var item in Func.GetDB.Corrections.Local.ToList())
            {
                switch (Func.GetDB.Entry(item).State)
                {
                    case EntityState.Detached:
                        break;
                    case EntityState.Unchanged:
                        break;
                    case EntityState.Added:
                        Func.GetDB.Corrections.Local.Remove(item);
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

            Func.GetDB.Corrections
                .Include(i => i.Головний_розпорядник)
                .Include(i => i.Змінив)
                .Include(i => i.КДБ)
                .Include(i => i.КЕКВ)
                .Include(i => i.КФК)
                .Include(i => i.Створив)
                .Include(i => i.Мікрофонд)
                .Include(i => i.Мікрофонд.Фонд)
                .Include(i => i.Мікрофонд.Фонд.Макрофонд)
                .Include(i => i.Статус)
                .Load();

            #endregion

            ((CollectionViewSource)FindResource("cvs")).Source = Func.GetDB.Corrections.Local;

            ((CollectionViewSource)FindResource("cvs")).Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Accept.Click += BTN_Accept_Click;
            BTN_Reset.Click += BTN_Reset_Click;
            BTN_ResetGroup.Click += BTN_ResetGroup_Click;
            BTN_Save.Click += BTN_Save_Click;
            BTN_ExportToExcel.Click += Func.BTN_ExportToExcel_Click;

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
        #endregion

        private void DGM_Loaded(object sender, RoutedEventArgs e)
        {
            if (Func.GetDB.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Correction == true) is null)
            {
                DGM.IsReadOnly = true;
            }
        }

        private void DGM_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                if (DGM.SelectedCells.Count > 0 && EXPEVAL.IsExpanded)
                {
                    if (DGM.SelectedCells.Count == 1)
                    {
                        if (Func.GetDB.names_months.Contains(e.AddedCells[0].Column.Header.ToString()) &&
                            ((DBSolom.Correction)e.AddedCells[0].Item).Головний_розпорядник != null &&
                            ((DBSolom.Correction)e.AddedCells[0].Item).КДБ != null &&
                            ((DBSolom.Correction)e.AddedCells[0].Item).КЕКВ != null &&
                            ((DBSolom.Correction)e.AddedCells[0].Item).КФК != null &&
                            ((DBSolom.Correction)e.AddedCells[0].Item).Мікрофонд != null &&
                            ((DBSolom.Correction)e.AddedCells[0].Item).Статус != null)
                        {
                            //Рассчёты, вычисления годового плана и уточненного

                            #region "Fiels of Cell"
                            DateTime date = new DateTime();
                            date = ((DBSolom.Correction)e.AddedCells[0].Item).Проведено;
                            var KFK = ((DBSolom.Correction)e.AddedCells[0].Item).КФК;
                            var Main_manager = ((DBSolom.Correction)e.AddedCells[0].Item).Головний_розпорядник;
                            var KDB = ((DBSolom.Correction)e.AddedCells[0].Item).КДБ;
                            var KEKB = ((DBSolom.Correction)e.AddedCells[0].Item).КЕКВ;
                            var FOND = ((DBSolom.Correction)e.AddedCells[0].Item).Мікрофонд.Фонд;
                            #endregion

                            #region "Годовой план"
                            double plan = 0;
                            var qplan = Func.GetDB.Fillings.FirstOrDefault(w =>
                                            w.Видалено == false &&
                                            w.Головний_розпорядник.Id == Main_manager.Id &&
                                            w.Проведено.Year == date.Year &&
                                            w.КДБ.Id == KDB.Id &&
                                            w.КЕКВ.Id == KEKB.Id &&
                                            w.КФК.Id == KFK.Id &&
                                            w.Фонд.Id == FOND.Id);
                            if (qplan != null)
                            {
                                plan = (double)qplan.GetType().GetProperty(e.AddedCells[0].Column.Header.ToString()).GetValue(qplan);
                            }
                            #endregion

                            #region "Довідки"
                            double corrections = 0;
                            var qcorr = Func.GetDB.Corrections.Local.Where(w => ((w?.Видалено ?? true) == false) &&
                                                                               ((w.Головний_розпорядник?.Id ?? 0) == Main_manager.Id) &&
                                                                               w.Проведено.Year == date.Year &&
                                                                               ((w.КДБ?.Id ?? 0) == KDB.Id) &&
                                                                               ((w.КЕКВ?.Id ?? 0) == KEKB.Id) &&
                                                                               ((w.КФК?.Id ?? 0) == KFK.Id) &&
                                                                               ((w.Мікрофонд?.Фонд?.Id ?? 0) == FOND.Id)).ToList();
                            if (qcorr.Count != 0)
                            {
                                corrections = qcorr.Select(s => (double)s.GetType().GetProperty(e.AddedCells[0].Column.Header.ToString()).GetValue(s)).Sum();
                            }

                            #endregion

                            GRPBPlan.Content = plan.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));

                            GRPBCorr.Content = corrections.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));

                            GRPBNow.Content = (plan + corrections).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));

                        }


                        //Рассчёты
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
                        GRPBPlan.Content = "Лише 1 комірка";
                        GRPBCorr.Content = "Лише 1 комірка";
                        GRPBNow.Content = "Лише 1 комірка";

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
            if (((DBSolom.Correction)e.Row.Item).Id == 0)
            {
                ((DBSolom.Correction)e.Row.Item).Створив = Func.GetDB.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            }
            ((DBSolom.Correction)e.Row.Item).Змінив = Func.GetDB.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            ((DBSolom.Correction)e.Row.Item).Змінено = DateTime.Now;


            if (((DBSolom.Correction)e.EditingElement.DataContext).Головний_розпорядник != null &&
                ((DBSolom.Correction)e.EditingElement.DataContext).КДБ != null &&
                ((DBSolom.Correction)e.EditingElement.DataContext).КЕКВ != null &&
                ((DBSolom.Correction)e.EditingElement.DataContext).КФК != null &&
                ((DBSolom.Correction)e.EditingElement.DataContext).Мікрофонд != null &&
                ((DBSolom.Correction)e.EditingElement.DataContext).Статус != null)
            {
                #region "Fiels of Cell"
                DateTime date = new DateTime();
                date = ((DBSolom.Correction)e.EditingElement.DataContext).Проведено;
                var KFK = ((DBSolom.Correction)e.EditingElement.DataContext).КФК;
                var Main_manager = ((DBSolom.Correction)e.EditingElement.DataContext).Головний_розпорядник;
                var KDB = ((DBSolom.Correction)e.EditingElement.DataContext).КДБ;
                var KEKB = ((DBSolom.Correction)e.EditingElement.DataContext).КЕКВ;
                var FOND = ((DBSolom.Correction)e.EditingElement.DataContext).Мікрофонд.Фонд;
                #endregion

                #region "Year_plan"

                var qplan = Func.GetDB.Fillings.FirstOrDefault(w => w.Видалено == false &&
                                                             w.Головний_розпорядник.Id == Main_manager.Id &&
                                                             w.Проведено.Year == date.Year &&
                                                             w.КДБ.Id == KDB.Id &&
                                                             w.КЕКВ.Id == KEKB.Id &&
                                                             w.КФК.Id == KFK.Id &&
                                                             w.Фонд.Id == FOND.Id);
                double plan = 0;
                if (qplan != null)
                {
                    foreach (var item in Func.GetDB.names_months)
                    {
                        plan += (double)qplan.GetType().GetProperty(item).GetValue(qplan);
                    }
                }

                #endregion

                #region "Corrections"
                double corrections = 0;
                var qcorr = Func.GetDB.Corrections.Local.Where(w => ((w?.Видалено ?? true) == false) &&
                                                                   ((w.Головний_розпорядник?.Id ?? 0) == Main_manager.Id) &&
                                                                   w.Проведено.Year == date.Year &&
                                                                   ((w.КДБ?.Id ?? 0) == KDB.Id) &&
                                                                   ((w.КЕКВ?.Id ?? 0) == KEKB.Id) &&
                                                                   ((w.КФК?.Id ?? 0) == KFK.Id) &&
                                                                   ((w.Мікрофонд?.Фонд?.Id ?? 0) == FOND.Id)).ToList();
                if (qcorr.Count != 0)
                {
                    foreach (var item in Func.GetDB.names_months)
                    {
                        corrections = qcorr.Select(s => (double)s.GetType().GetProperty(item).GetValue(s)).Sum();
                    }
                }
                #endregion

                PropertyInfo k = null;

                if (plan + corrections < 0)
                {
                    e.Cancel = true;
                    if (Func.GetDB.names_months.Contains(e.Column.Header.ToString()))
                    {
                        k = ((DBSolom.Correction)e.Row.Item).GetType().GetProperty(e.Column.Header.ToString());
                        k.SetValue(e.Row.Item, 0);
                        ((TextBox)e.EditingElement).Text = "0";
                    }
                    
                    MessageBox.Show("Недостатньо коштів! Річний план: " + plan + "; Корегування: " + (corrections) + "; Уточнення: " + (plan - corrections));
                    return;
                }
            }

        }

        private void DGM_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            Func.GenerateColumnForDataGrid(ref counterForDGMColumns, e);
        }

        private void DGM_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            IsAllowEditing(e);
        }

        private void IsAllowEditing(DataGridBeginningEditEventArgs e)
        {
            if (((DBSolom.Correction)e.Row.Item).Id != 0 && ((DBSolom.Correction)e.Row.Item).Статус?.Повністю != "Зареєстровано")
            {
                if (Func.Login == "LeXX" || ((DBSolom.Correction)e.Row.Item).Змінив.Логін == Func.Login)
                {
                    ((DBSolom.Correction)e.Row.Item).Статус = Func.GetDB.DocStatuses.Include(i => i.Змінив).Include(i => i.Створив).FirstOrDefault(f => f.Видалено == false && f.Повністю == "Зареєстровано");
                    var cellContent = DGM.Columns.First(f => f.Header.ToString() == "Статус").GetCellContent(e.Row);
                    if (cellContent is ComboBox)
                    {
                        ((ComboBox)cellContent).Text = "Зареєстровано";
                    }
                    e.Cancel = false;
                }
                else if (MessageBox.Show("Ви маєте пароль для редагування?\n\t(пр. пароль це пароль користувача який вніс зміни)", "Редагування заблоковано", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    string pass = Microsoft.VisualBasic.Interaction.InputBox("Пароль: ", "");
                    if (((DBSolom.Correction)e.Row.Item).Змінив.Пароль == pass)
                    {
                        ((DBSolom.Correction)e.Row.Item).Статус = Func.GetDB.DocStatuses.Include(i => i.Змінив).Include(i => i.Створив).FirstOrDefault(f => f.Видалено == false && f.Повністю == "Зареєстровано");
                        var cellContent = DGM.Columns.First(f => f.Header.ToString() == "Статус").GetCellContent(e.Row);
                        if (cellContent is ComboBox)
                        {
                            ((ComboBox)cellContent).Text = "Зареєстровано";
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
    }

    #region "Converters"

    public class CorrectionGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return "Ʃ: " + CheckedFillingItems(items).ToString("N2",
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
                                select ((DBSolom.Correction)x).Січень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Лютий).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Березень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Квітень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Травень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Червень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Липень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Серпень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Вересень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Жовтень).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Листопад).Sum();
                        sum += (from x in ((CollectionViewGroup)items[i]).Items
                                select ((DBSolom.Correction)x).Грудень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Січень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Лютий).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Березень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Квітень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Травень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Червень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Липень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Серпень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Вересень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Жовтень).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Листопад).Sum();
                sum += (from x in items
                        select ((DBSolom.Correction)x).Грудень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupOneConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Січень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Січень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupTwoConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Лютий).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Лютий).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupThreeConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Березень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Березень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupFourConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Квітень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Квітень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupFiveConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Травень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Травень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupSixConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Червень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Червень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupSevenConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Липень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Липень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupEightConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Серпень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Серпень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupNineConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Вересень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Вересень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupTenConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Жовтень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Жовтень).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupElevenConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Листопад).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Листопад).Sum();
                return sum;
            }
        }
    }

    public class CorrectionGroupTwelveConverter : IValueConverter
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
                                select ((DBSolom.Correction)x).Грудень).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (from x in items
                        select ((DBSolom.Correction)x).Грудень).Sum();
                return sum;
            }
        }
    }

    #endregion
}
