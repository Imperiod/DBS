using DBSolom;
using Main;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;

namespace Maestro.Functional
{
    public partial class Summary : Window
    {
        #region "Variables"

        public string type = "";
        public string prop = "";
        public object value = null;
        public ForWidth GetWidth { get; set; } = new ForWidth();
        public SummaryEntities GetSummaryEntities { get; set; }
        public List<Label> GetLabels = new List<Label>();
        public Dictionary<string, ComboBox> dict_cmb = new Dictionary<string, ComboBox>();
        public Dictionary<string, TextBox> dict_txt = new Dictionary<string, TextBox>();
        List<Filters> GetFilters = new List<Filters>();
        public List<ToggleButton> CheckBoxes = new List<ToggleButton>();
        bool IsInitialization = true;
        CollectionViewSource CollectionViewSource { get; set; }

        DBSolom.Db db { get; set; }

        #endregion

        public Summary()
        {
            InitializeComponent();
            CollectionViewSource = ((CollectionViewSource)FindResource("cvs"));

            CollectionViewSource.Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Accept.Click += BTN_Accept_Click;
            BTN_Reset.Click += BTN_Reset_Click;
            BTN_ResetGroup.Click += BTN_ResetGroup_Click;
            BTN_ExportToExcel.Click += Func.BTN_ExportToExcel_Click;
        }

        private void DatePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EDate.SelectedDate != null)
            {
                FillDate(EDate.SelectedDate.Value);
            }
        }

        public void FillDate(DateTime D)
        {
            db = new Db(Func.GetConnectionString);

            #region "Clear filters, groups and visibility"
            if (IsInitialization == false)
            {
                dict_cmb.Values.ToList().ForEach(cmb => cmb.SelectedValue = null);
                dict_txt.Values.ToList().ForEach(txt => txt.Text = null);
                CheckBoxes.ForEach(a => a.IsChecked = false);
                EXPHDN.Children.Cast<ToggleButton>().ToList().ForEach(tgb => tgb.IsChecked = true);

                type = "";
                prop = "";
                value = null;

                LBFilters.Items.Clear();
                GetFilters.Clear();
            }
            #endregion

            GetSummaryEntities = new SummaryEntities(db, new DateTime(D.Year, 1, 1, 0, 0, 0, 0), D);
            CollectionViewSource.Source = GetSummaryEntities.SummaryEntitiesSource;
            DGM.ItemsSource = CollectionViewSource.View;

            if (IsInitialization)
            {
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
            dict_cmb.Values.ToList().ForEach(cmb => cmb.SelectedValue = null);
            dict_txt.Values.ToList().ForEach(txt => txt.Text = null);

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
        #endregion

        private void DGM_SelectedCellsChanged(object sender, SelectedCellsChangedEventArgs e)
        {
            try
            {
                if (DGM.SelectedCells.Count > 0 && EVAL.IsExpanded)
                {
                    if (DGM.SelectedCells.Count == 1)
                    {
                        if (Func.names_months.Concat(new List<string>() { "Рік", "Період" }).Contains(e.AddedCells[0].Column.Header.ToString()))
                        {
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

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            GetWidth.Width = e.NewSize.Width-80;
        }
    }

    public class ForWidth : INotifyPropertyChanged
    {
        double width;
        public double Width
        {
            get { return width; }
            set
            {
                width = value;
                OnPropertyChanged("Width");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([System.Runtime.CompilerServices.CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }

    public class SummaryEntities
    {
        public List<SummaryEntity> SummaryEntitiesSource { get; set; }

        DBSolom.Db db { get; set; }

        public SummaryEntities(DBSolom.Db db, DateTime start, DateTime end)
        {
            this.db = db;
            SummaryEntitiesSource = new List<SummaryEntity>();
            GetSummaryEntities(start, end);
        }

        private void GetSummaryEntities(DateTime s, DateTime e)
        {
            long Id = 1;
            DateTime endDateTime = e + TimeSpan.FromDays(1) - TimeSpan.FromSeconds(1);

            #region Queries

            var QueryToMicrofillings = db.Microfillings
                            .Include(i => i.Змінив)
                            .Include(i => i.КДБ)
                            .Include(i => i.КФБ)
                            .Include(i => i.Створив)
                            .Include(i => i.Головний_розпорядник)
                            .Include(i => i.КЕКВ)
                            .Include(i => i.КФК)
                            .Include(i => i.Мікрофонд)
                            .Include(i => i.Мікрофонд.Фонд)
                    .Where(w => w.Видалено == false && w.Проведено.Year >= s.Year && w.Проведено.Year <= e.Year).ToList();

            var QueryToCorrections = db.Corrections
                            .Include(i => i.Змінив)
                            .Include(i => i.КДБ)
                            .Include(i => i.КФБ)
                            .Include(i => i.Створив)
                            .Include(i => i.Головний_розпорядник)
                            .Include(i => i.КЕКВ)
                            .Include(i => i.КФК)
                            .Include(i => i.Мікрофонд)
                            .Include(i => i.Мікрофонд.Фонд)
                            .Include(i => i.Статус)
                    .Where(w => w.Видалено == false && w.Проведено >= s && w.Проведено <= endDateTime).ToList();

            var QueryToFinancing = db.Financings
                            .Include(i => i.Змінив)
                            .Include(i => i.Створив)
                            .Include(i => i.Головний_розпорядник)
                            .Include(i => i.КЕКВ)
                            .Include(i => i.КФК)
                            .Include(i => i.Мікрофонд)
                            .Include(i => i.Мікрофонд.Фонд)
                    .Where(w => w.Видалено == false && w.Проведено >= s && w.Проведено <= endDateTime).ToList();

            #endregion

            //Add annual plans
            foreach (var item in QueryToMicrofillings)
            {
                SummaryEntity summaryEntityMicrofilling = new SummaryEntity()
                {
                    SortIndex = 0,
                    Id = Id,
                    Тип = "План",
                    Створив = item.Створив,
                    Створино = item.Створино,
                    Змінив = item.Змінив,
                    Змінено = item.Змінено,
                    Проведено = item.Проведено,
                    Підписано = item.Підписано,
                    Внутрішній_номер = null,
                    Підстава = null,
                    Статус = null,
                    Головний_розпорядник = item.Головний_розпорядник,
                    КФК = item.КФК,
                    Фонд = item.Мікрофонд.Фонд,
                    Мікрофонд = item.Мікрофонд,
                    КФБ = item.КФБ,
                    КДБ = item.КДБ,
                    КЕКВ = item.КЕКВ,
                    Сума = item.Січень + item.Лютий + item.Березень + item.Квітень + item.Травень + item.Червень + item.Липень + item.Серпень + item.Вересень + item.Жовтень + item.Листопад + item.Грудень,
                    Січень = item.Січень,
                    Лютий = item.Лютий,
                    Березень = item.Березень,
                    Квітень = item.Квітень,
                    Травень = item.Травень,
                    Червень = item.Червень,
                    Липень = item.Липень,
                    Серпень = item.Серпень,
                    Вересень = item.Вересень,
                    Жовтень = item.Жовтень,
                    Листопад = item.Листопад,
                    Грудень = item.Грудень
                };

                SummaryEntitiesSource.Add(summaryEntityMicrofilling);
                Id++;
            }

            //Add corrections
            foreach (var item in QueryToCorrections)
            {
                SummaryEntity summaryEntityMicrofilling = new SummaryEntity()
                {
                    SortIndex = 1,
                    Id = Id,
                    Тип = "Зміни",
                    Створив = item.Створив,
                    Створино = item.Створино,
                    Змінив = item.Змінив,
                    Змінено = item.Змінено,
                    Проведено = item.Проведено,
                    Підписано = true,
                    Внутрішній_номер = item.Внутрішній_номер,
                    Підстава = item.Підстава,
                    Статус = item.Статус,
                    Головний_розпорядник = item.Головний_розпорядник,
                    КФК = item.КФК,
                    Фонд = item.Мікрофонд.Фонд,
                    Мікрофонд = item.Мікрофонд,
                    КФБ = item.КФБ,
                    КДБ = item.КДБ,
                    КЕКВ = item.КЕКВ,
                    Сума = item.Січень + item.Лютий + item.Березень + item.Квітень + item.Травень + item.Червень + item.Липень + item.Серпень + item.Вересень + item.Жовтень + item.Листопад + item.Грудень,
                    Січень = item.Січень,
                    Лютий = item.Лютий,
                    Березень = item.Березень,
                    Квітень = item.Квітень,
                    Травень = item.Травень,
                    Червень = item.Червень,
                    Липень = item.Липень,
                    Серпень = item.Серпень,
                    Вересень = item.Вересень,
                    Жовтень = item.Жовтень,
                    Листопад = item.Листопад,
                    Грудень = item.Грудень
                };

                SummaryEntitiesSource.Add(summaryEntityMicrofilling);
                Id++;
            }

            var Q = SummaryEntitiesSource.ToList();

            //Add new annual plans like corrected plans
            foreach (var item in QueryToCorrections)
            {
                var concret = Q.Where(w =>
                    w.Головний_розпорядник.Найменування == item.Головний_розпорядник.Найменування &&
                    w.КДБ.Код == item.КДБ.Код &&
                    w.КФБ.Код == item.КФБ.Код &&
                    w.Проведено <= item.Проведено &&
                    w.КЕКВ.Код == item.КЕКВ.Код &&
                    w.КФК.Код == item.КФК.Код &&
                    w.Мікрофонд.Повністю == item.Мікрофонд.Повністю).ToList();

                SummaryEntity summaryEntityMicrofilling = new SummaryEntity()
                {
                    SortIndex = 2,
                    Id = Id,
                    Тип = "Уточнення",
                    Створив = item.Створив,
                    Створино = item.Створино,
                    Змінив = item.Змінив,
                    Змінено = item.Змінено,
                    Проведено = item.Проведено,
                    Підписано = true,
                    Внутрішній_номер = item.Внутрішній_номер,
                    Підстава = item.Підстава,
                    Статус = item.Статус,
                    Головний_розпорядник = item.Головний_розпорядник,
                    КФК = item.КФК,
                    Фонд = item.Мікрофонд.Фонд,
                    Мікрофонд = item.Мікрофонд,
                    КФБ = item.КФБ,
                    КДБ = item.КДБ,
                    КЕКВ = item.КЕКВ,
                    Сума = concret.Sum(ss => ss.Сума),
                    Січень = concret.Sum(ss => ss.Січень),
                    Лютий = concret.Sum(ss => ss.Лютий),
                    Березень = concret.Sum(ss => ss.Березень),
                    Квітень = concret.Sum(ss => ss.Квітень),
                    Травень = concret.Sum(ss => ss.Травень),
                    Червень = concret.Sum(ss => ss.Червень),
                    Липень = concret.Sum(ss => ss.Липень),
                    Серпень = concret.Sum(ss => ss.Серпень),
                    Вересень = concret.Sum(ss => ss.Вересень),
                    Жовтень = concret.Sum(ss => ss.Жовтень),
                    Листопад = concret.Sum(ss => ss.Листопад),
                    Грудень = concret.Sum(ss => ss.Грудень)
                };

                SummaryEntitiesSource.Add(summaryEntityMicrofilling);
                Id++;
            }

            //Add financing
            foreach (var item in QueryToFinancing)
            {
                SummaryEntity summaryEntityMicrofilling = new SummaryEntity()
                {
                    SortIndex = 3,
                    Id = Id,
                    Тип = "Фінансування",
                    Створив = item.Створив,
                    Створино = item.Створино,
                    Змінив = item.Змінив,
                    Змінено = item.Змінено,
                    Проведено = item.Проведено,
                    Підписано = item.Підписано,
                    Внутрішній_номер = null,
                    Підстава = null,
                    Статус = null,
                    Головний_розпорядник = item.Головний_розпорядник,
                    КФК = item.КФК,
                    Фонд = item.Мікрофонд.Фонд,
                    Мікрофонд = item.Мікрофонд,
                    КФБ = null,
                    КДБ = null,
                    КЕКВ = item.КЕКВ,
                    Сума = item.Сума * -1,
                    Січень = item.Проведено.Month == 1 ? item.Сума * -1 : 0,
                    Лютий = item.Проведено.Month == 2 ? item.Сума * -1 : 0,
                    Березень = item.Проведено.Month == 3 ? item.Сума * -1 : 0,
                    Квітень = item.Проведено.Month == 4 ? item.Сума * -1 : 0,
                    Травень = item.Проведено.Month == 5 ? item.Сума * -1 : 0,
                    Червень = item.Проведено.Month == 6 ? item.Сума * -1 : 0,
                    Липень = item.Проведено.Month == 7 ? item.Сума * -1 : 0,
                    Серпень = item.Проведено.Month == 8 ? item.Сума * -1 : 0,
                    Вересень = item.Проведено.Month == 9 ? item.Сума * -1 : 0,
                    Жовтень = item.Проведено.Month == 10 ? item.Сума * -1 : 0,
                    Листопад = item.Проведено.Month == 11 ? item.Сума * -1 : 0,
                    Грудень = item.Проведено.Month == 12 ? item.Сума * -1 : 0
                };

                SummaryEntitiesSource.Add(summaryEntityMicrofilling);
                Id++;
            }

            //Add remainder per financing
            foreach (var item in QueryToFinancing)
            {
                var concret = SummaryEntitiesSource.Where(w =>
                    (w.Тип == "Фінансування" || w.Тип == "План" || w.Тип == "Зміни") &&
                    w.Головний_розпорядник.Найменування == item.Головний_розпорядник.Найменування &&
                    w.Проведено <= item.Проведено &&
                    w.КЕКВ.Код == item.КЕКВ.Код &&
                    w.КФК.Код == item.КФК.Код &&
                    w.Мікрофонд.Повністю == item.Мікрофонд.Повністю).ToList();

                SummaryEntity summaryEntityMicrofilling = new SummaryEntity()
                {
                    SortIndex = 4,
                    Id = Id,
                    Тип = "Залишок",
                    Створив = item.Створив,
                    Створино = item.Створино,
                    Змінив = item.Змінив,
                    Змінено = item.Змінено,
                    Проведено = item.Проведено,
                    Підписано = true,
                    Внутрішній_номер = null,
                    Підстава = null,
                    Статус = null,
                    Головний_розпорядник = item.Головний_розпорядник,
                    КФК = item.КФК,
                    Фонд = item.Мікрофонд.Фонд,
                    Мікрофонд = item.Мікрофонд,
                    КФБ = null,
                    КДБ = null,
                    КЕКВ = item.КЕКВ,
                    Сума = concret.Sum(ss => ss.Сума),
                    Січень = 0,
                    Лютий = 0,
                    Березень = 0,
                    Квітень = 0,
                    Травень = 0,
                    Червень = 0,
                    Липень = 0,
                    Серпень = 0,
                    Вересень = 0,
                    Жовтень = 0,
                    Листопад = 0,
                    Грудень = 0
                };

                //per month
                for (int i = 0; i < item.Проведено.Month; i++)
                {
                    double d = 0;
                    for (int m = 0; m <= i; m++)
                    {
                        d += concret.Sum(ss => (double)ss.GetType().GetProperty(Func.names_months[m]).GetValue(ss));
                    }
                    summaryEntityMicrofilling.GetType().GetProperty(Func.names_months[i]).SetValue(summaryEntityMicrofilling, d);
                }

                SummaryEntitiesSource.Add(summaryEntityMicrofilling);
                Id++;
            }

            SummaryEntitiesSource = SummaryEntitiesSource.OrderBy(o => o.Проведено).ThenBy(t => t.SortIndex).ToList();
        }
    }

    public class SummaryEntity
    {
        public byte SortIndex { get; set; }

        public long Id { get; set; }

        public string Тип { get; set; }

        public User Створив { get; set; }

        public DateTime Створино { get; set; } = DateTime.Now;

        public User Змінив { get; set; }

        public DateTime Змінено { get; set; } = DateTime.Now;

        public DateTime Проведено { get; set; } = DateTime.Now;

        public bool Підписано { get; set; } = false;

        public string Внутрішній_номер { get; set; }

        public string Підстава { get; set; }

        public DocStatus Статус { get; set; }

        public Main_manager Головний_розпорядник { get; set; }

        public KFK КФК { get; set; }

        public Foundation Фонд { get; set; }

        public MicroFoundation Мікрофонд { get; set; }

        public KFB КФБ { get; set; }

        public KDB КДБ { get; set; }

        public KEKB КЕКВ { get; set; }

        public double Сума { get; set; }

        public double Січень { get; set; }

        public double Лютий { get; set; }

        public double Березень { get; set; }

        public double Квітень { get; set; }

        public double Травень { get; set; }

        public double Червень { get; set; }

        public double Липень { get; set; }

        public double Серпень { get; set; }

        public double Вересень { get; set; }

        public double Жовтень { get; set; }

        public double Листопад { get; set; }

        public double Грудень { get; set; }
    }

    #region Converters

    public class SummaryDigitConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double d;
            if (double.TryParse(value.ToString(), out d))
            {
                return d.ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
            }
            else
            {
                return value;
            }
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {

            value = value.ToString().Replace('.', ',');
            if (value.ToString().Where(s => s == ',').Count() > 1)
            {
                string s = "";
                for (int i = 0; i < value.ToString().Length; i++)
                {
                    if (value.ToString()[i] != ',' || i == value.ToString().LastIndexOf(','))
                    {
                        s += value.ToString()[i];
                    }
                }
                value = s;
            }
            else if (value.ToString().Length > 4 && value.ToString().Where(s => s == ',').Count() == 0)
            {
                value = value.ToString().Insert(value.ToString().Length - 2, ",");
            }

            double d;
            if (double.TryParse((string)value, out d))
            {
                return double.Parse(value.ToString(), CultureInfo.CreateSpecificCulture("ru-RU"));
            }
            else
            {

                return value;
            }
        }
    }

    public class SummaryWidthConverterForColumnHeaderOfGroupFirst : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (175.0 / 1880.0) * (double)value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    public class SummaryWidthConverterForColumnHeaderOfGroupSecond : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (145.0 / 1880.0) * (double)value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    public class SummaryWidthConverterForColumnHeaderOfGroupMonths : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (130.0 / 1880.0) * (double)value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    public class SummaryConverterForNameOfGroup : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string s = "";

            while (value != null && value.GetType().GetProperty("Name").GetValue(value).ToString() != "Root")
            {
                s = $" [{value.GetType().GetProperty("Name").GetValue(value).ToString()}] " + s;
                value = ((System.Reflection.PropertyInfo[])((System.Reflection.TypeInfo)value.GetType()).DeclaredProperties).FirstOrDefault(w => w.Name == "Parent")?.GetValue(value);
            }
            
            return s;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    #region AnnualPlanConverters
    public class SummaryAPlanGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((SummaryEntity)w).Тип == "План").Select(s => ((SummaryEntity)s).Сума).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((SummaryEntity)w).Тип == "План").Select(s => ((SummaryEntity)s).Сума).Sum();
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupTwoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupThreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupFourConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupFiveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupSixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupSevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupEightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupNineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupTenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupElevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryAPlanGroupTwelveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План")
                                    .Select(s => ((SummaryEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План")
                            .Select(s => ((SummaryEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }
    #endregion

    #region CorrectionConverters
    public class SummaryCorrGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((SummaryEntity)w).Тип == "Зміни").Select(s => ((SummaryEntity)s).Сума).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((SummaryEntity)w).Тип == "Зміни").Select(s => ((SummaryEntity)s).Сума).Sum();
                return sum;
            }
        }
    }
    public class SummaryCorrGroupOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupTwoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupThreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupFourConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupFiveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupSixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupSevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupEightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupNineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupTenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupElevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCorrGroupTwelveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }
    #endregion

    #region CorrectedPlansConverters
    public class SummaryCPlanGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни").Select(s => ((SummaryEntity)s).Сума).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни").Select(s => ((SummaryEntity)s).Сума).Sum();
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupTwoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupThreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupFourConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupFiveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupSixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupSevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupEightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupNineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupTenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupElevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryCPlanGroupTwelveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                                    .Select(s => ((SummaryEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "План" || ((SummaryEntity)w).Тип == "Зміни")
                            .Select(s => ((SummaryEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }
    #endregion

    #region FinancingConverters
    public class SummaryFinGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((SummaryEntity)w).Тип == "Фінансування").Select(s => ((SummaryEntity)s).Сума).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((SummaryEntity)w).Тип == "Фінансування").Select(s => ((SummaryEntity)s).Сума).Sum();
                return sum;
            }
        }
    }
    public class SummaryFinGroupOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupTwoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupThreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupFourConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupFiveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupSixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupSevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupEightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupNineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupTenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupElevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryFinGroupTwelveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                                    .Select(s => ((SummaryEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип == "Фінансування")
                            .Select(s => ((SummaryEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }
    #endregion

    #region RemainderConverters
    /// <summary>
    /// These converters are calculated, I'm mean that remainder = annual plan + changes (corrections) + financing
    /// So in code type != remainder && type != Corrected plan
    /// </summary>
    public class SummaryRemGroupTotalConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Сума)
                            .Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Сума)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupOneConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupTwoConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupThreeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupFourConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupFiveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupSixConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupSevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupEightConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupNineConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupTenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupElevenConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class SummaryRemGroupTwelveConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (null == value)
            {
                return "null";
            }

            ReadOnlyObservableCollection<object> items = (ReadOnlyObservableCollection<object>)value;

            return CheckedFillingItems(items).ToString("N2", CultureInfo.CreateSpecificCulture("ru-RU"));
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
                        sum += (((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                                    .Select(s => ((SummaryEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((SummaryEntity)w).Тип != "Залишок" && ((SummaryEntity)w).Тип != "Уточнення")
                            .Select(s => ((SummaryEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }
    #endregion

    #endregion
}
