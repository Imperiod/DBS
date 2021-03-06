﻿using DBSolom;
using Main;
using Microsoft.CSharp;
using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;

namespace Maestro.Functional
{
    public partial class Remainders : Window
    {
        #region "Variables"

        CurrPlanEntities GetCurrPlanEntities { get; set; }

        public string type = "";
        public string prop = "";
        public object value = null;

        public RemainderForWidth GetWidth { get; set; } = new RemainderForWidth();

        public List<Label> GetLabels = new List<Label>();
        public Dictionary<string, ComboBox> dict_cmb = new Dictionary<string, ComboBox>();
        public Dictionary<string, TextBox> dict_txt = new Dictionary<string, TextBox>();
        List<Filters> GetFilters = new List<Filters>();
        public List<ToggleButton> CheckBoxes = new List<ToggleButton>();
        bool IsInitialization = true;
        CollectionViewSource CollectionViewSource { get; set; }

        DBSolom.Db db { get; set; }

        #endregion

        public Remainders()
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
            if (SDate.SelectedDate != null && EDate.SelectedDate != null)
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

                GetCurrPlanEntities = new CurrPlanEntities(db, (DateTime)SDate.SelectedDate, (DateTime)EDate.SelectedDate);
                CollectionViewSource.Source = GetCurrPlanEntities.GetEntities;
                DGM.ItemsSource = CollectionViewSource.View;

                if (IsInitialization)
                {
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("КФК.Код", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Головний_розпорядник.Найменування", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Фонд.Код", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Мікрофонд.Повністю", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("КЕКВ.Код", ListSortDirection.Ascending));
                    CollectionViewSource.SortDescriptions.Add(new SortDescription("Дані", ListSortDirection.Ascending));

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
            GetWidth.Width = e.NewSize.Width - 80;
        }
    }

    public class RemainderForWidth : INotifyPropertyChanged
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

    class CurrPlanEntities
    {
        DBSolom.Db db { get; set; }

        public List<CurrPlanEntity> GetEntities { get; set; }

        public CurrPlanEntities(DBSolom.Db db, DateTime start, DateTime end)
        {
            this.db = db;
            GetEntities = new List<CurrPlanEntity>();
            FillEntities(start, end);
        }

        private void FillEntities(DateTime s, DateTime e)
        {
            long Id = 1;
            DateTime endDateTime = e.AddDays(1) - TimeSpan.FromSeconds(1);

            #region Queries

            var QueryToAnnualPlans = db.Microfillings
                            .Include(i => i.Головний_розпорядник)
                            .Include(i => i.КЕКВ)
                            .Include(i => i.КФК)
                            .Include(i => i.Мікрофонд.Фонд)
                    .Where(w => w.Видалено == false && w.Проведено.Year >= s.Year && w.Проведено.Year <= e.Year)
                    .Select(ss => new
                    {
                        Фонд = ss.Мікрофонд.Фонд,
                        Мікрофонд = ss.Мікрофонд,
                        КФК = ss.КФК,
                        Головний_розпорядник = ss.Головний_розпорядник,
                        КЕКВ = ss.КЕКВ,
                        ss.Січень,
                        ss.Лютий,
                        ss.Березень,
                        ss.Квітень,
                        ss.Травень,
                        ss.Червень,
                        ss.Липень,
                        ss.Серпень,
                        ss.Вересень,
                        ss.Жовтень,
                        ss.Листопад,
                        ss.Грудень
                    })
                    .ToList();

            var QueryToCorrections = db.Corrections
                    .Include(i => i.Головний_розпорядник)
                    .Include(i => i.КЕКВ)
                    .Include(i => i.КФК)
                    .Include(i => i.Мікрофонд.Фонд)
                    .Where(w => w.Видалено == false && w.Проведено >= s && w.Проведено <= endDateTime)
                    .Select(ss => new
                    {
                        Фонд = ss.Мікрофонд.Фонд,
                        Мікрофонд = ss.Мікрофонд,
                        КФК = ss.КФК,
                        Головний_розпорядник = ss.Головний_розпорядник,
                        КЕКВ = ss.КЕКВ,
                        ss.Січень,
                        ss.Лютий,
                        ss.Березень,
                        ss.Квітень,
                        ss.Травень,
                        ss.Червень,
                        ss.Липень,
                        ss.Серпень,
                        ss.Вересень,
                        ss.Жовтень,
                        ss.Листопад,
                        ss.Грудень
                    })
                    .ToList();

            var QueryToFinancings = db.Financings
                    .Include(i => i.Головний_розпорядник)
                    .Include(i => i.КЕКВ)
                    .Include(i => i.КФК)
                    .Include(i => i.Мікрофонд.Фонд)
                    .Where(w => w.Видалено == false && w.Проведено >= s && w.Проведено <= endDateTime)
                                        .Select(ss => new
                                        {
                                            Місяць = ss.Проведено.Month,
                                            Фонд = ss.Мікрофонд.Фонд,
                                            Мікрофонд = ss.Мікрофонд,
                                            КФК = ss.КФК,
                                            Головний_розпорядник = ss.Головний_розпорядник,
                                            КЕКВ = ss.КЕКВ,
                                            ss.Сума
                                        })
                                        .GroupBy(g => new { g.Фонд, g.Мікрофонд, g.КФК, g.Головний_розпорядник, g.КЕКВ, g.Місяць })
                                        .ToList();

            #endregion

            QueryToAnnualPlans.AddRange(QueryToCorrections);
            var CorrectedPlans = QueryToAnnualPlans.GroupBy(g => new { g.Фонд, g.Мікрофонд, g.КФК, g.Головний_розпорядник, g.КЕКВ }).ToList();

            for (int i = 0; i < CorrectedPlans.Count; i++)
            {
                var item = CorrectedPlans[i];
                CurrPlanEntity currPlanEntity = new CurrPlanEntity(
                    Id, e.Month, item.Key.Головний_розпорядник, item.Key.Фонд, item.Key.Мікрофонд, item.Key.КФК, item.Key.КЕКВ, "План",
                    item.Sum(ss => ss.Січень),
                    item.Sum(ss => ss.Лютий),
                    item.Sum(ss => ss.Березень),
                    item.Sum(ss => ss.Квітень),
                    item.Sum(ss => ss.Травень),
                    item.Sum(ss => ss.Червень),
                    item.Sum(ss => ss.Липень),
                    item.Sum(ss => ss.Серпень),
                    item.Sum(ss => ss.Вересень),
                    item.Sum(ss => ss.Жовтень),
                    item.Sum(ss => ss.Листопад),
                    item.Sum(ss => ss.Грудень));
                GetEntities.Add(currPlanEntity);
                Id++;
            }

            for (int i = 0; i < QueryToFinancings.Count; i++)
            {
                var item = QueryToFinancings[i];

                  CurrPlanEntity currPlanEntity = new CurrPlanEntity(
                        Id, item.Key.Місяць, item.Key.Головний_розпорядник, item.Key.Фонд, item.Key.Мікрофонд, item.Key.КФК, item.Key.КЕКВ, "Факт",
                                            item.Key.Місяць == 1 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 2 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 3 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 4 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 5 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 6 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 7 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 8 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 9 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 10 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 11 ? item.Sum(ss => ss.Сума) : 0,
                                            item.Key.Місяць == 12 ? item.Sum(ss => ss.Сума) : 0);

                    GetEntities.Add(currPlanEntity);
                    Id++;
                
            }

            var QueryToCorrectedPlans = GetEntities.Where(w => w.Дані == "План").ToList();

            foreach (var item in QueryToCorrectedPlans)
            {
                var fact = GetEntities.Where(f =>
                f.Головний_розпорядник == item.Головний_розпорядник &&
                f.КЕКВ == item.КЕКВ &&
                f.КФК == item.КФК &&
                f.Мікрофонд == item.Мікрофонд &&
                f.Дані == "Факт");

                CurrPlanEntity currPlanEntity = new CurrPlanEntity(Id, e.Month, item.Головний_розпорядник, item.Фонд, item.Мікрофонд, item.КФК, item.КЕКВ, "Н_Залишок",
                                            fact == null ? item.Січень : item.Січень - fact.Sum(ss=>ss.Січень),
                                            fact == null ? item.Лютий : item.Лютий - fact.Sum(ss => ss.Лютий),
                                            fact == null ? item.Березень : item.Березень - fact.Sum(ss => ss.Березень),
                                            fact == null ? item.Квітень : item.Квітень - fact.Sum(ss=>ss.Квітень),
                                            fact == null ? item.Травень : item.Травень - fact.Sum(ss=>ss.Травень),
                                            fact == null ? item.Червень : item.Червень - fact.Sum(ss => ss.Червень),
                                            fact == null ? item.Липень : item.Липень - fact.Sum(ss => ss.Липень),
                                            fact == null ? item.Серпень : item.Серпень - fact.Sum(ss => ss.Серпень),
                                            fact == null ? item.Вересень : item.Вересень - fact.Sum(ss => ss.Вересень),
                                            fact == null ? item.Жовтень : item.Жовтень - fact.Sum(ss => ss.Жовтень),
                                            fact == null ? item.Листопад : item.Листопад - fact.Sum(ss => ss.Листопад),
                                            fact == null ? item.Грудень : item.Грудень - fact.Sum(ss => ss.Грудень));
                GetEntities.Add(currPlanEntity);
                Id++;

                CurrPlanEntity currPlanEntityTwo = new CurrPlanEntity(Id, e.Month, item.Головний_розпорядник, item.Фонд, item.Мікрофонд, item.КФК, item.КЕКВ, "М_Залишок",
                                            fact == null ? item.Січень : item.Січень - fact.Sum(ss => ss.Січень),
                                            fact == null ? item.Лютий : item.Лютий - fact.Sum(ss => ss.Лютий),
                                            fact == null ? item.Березень : item.Березень - fact.Sum(ss => ss.Березень),
                                            fact == null ? item.Квітень : item.Квітень - fact.Sum(ss => ss.Квітень),
                                            fact == null ? item.Травень : item.Травень - fact.Sum(ss => ss.Травень),
                                            fact == null ? item.Червень : item.Червень - fact.Sum(ss => ss.Червень),
                                            fact == null ? item.Липень : item.Липень - fact.Sum(ss => ss.Липень),
                                            fact == null ? item.Серпень : item.Серпень - fact.Sum(ss => ss.Серпень),
                                            fact == null ? item.Вересень : item.Вересень - fact.Sum(ss => ss.Вересень),
                                            fact == null ? item.Жовтень : item.Жовтень - fact.Sum(ss => ss.Жовтень),
                                            fact == null ? item.Листопад : item.Листопад - fact.Sum(ss => ss.Листопад),
                                            fact == null ? item.Грудень : item.Грудень - fact.Sum(ss => ss.Грудень));
                GetEntities.Add(currPlanEntityTwo);
                Id++;
            }
        }
    }

    class CurrPlanEntity
    {
        public CurrPlanEntity(long Id, int Month, Main_manager main_Manager, Foundation Fond, MicroFoundation microFoundation, KFK kfk, KEKB kekb, string type, double one, double two, double three, double four, double five, double six, double seven, double eight, double nine, double ten, double eleven, double twelve)
        {
            this.Id = Id;
            Головний_розпорядник = main_Manager;
            Фонд = Fond;
            Мікрофонд = microFoundation;
            КФК = kfk;
            КЕКВ = kekb;
            Дані = type;

            if (type == "Н_Залишок")
            {
                Рік = one + two + three + four + five + six + seven + eight + nine + ten + eleven + twelve;
                Січень = one;
                Лютий = one + two;
                Березень = one + two + three;
                Квітень = one + two + three + four;
                Травень = one + two + three + four + five;
                Червень = one + two + three + four + five + six;
                Липень = one + two + three + four + five + six + seven;
                Серпень = one + two + three + four + five + six + seven + eight;
                Вересень = one + two + three + four + five + six + seven + eight + nine;
                Жовтень = one + two + three + four + five + six + seven + eight + nine + ten;
                Листопад = one + two + three + four + five + six + seven + eight + nine + ten + eleven;
                Грудень = one + two + three + four + five + six + seven + eight + nine + ten + eleven + twelve;

                Період = (double)GetType().GetProperty(Func.names_months[Month - 1].ToString()).GetValue(this);
            }
            else if (type == "М_Залишок")
            {
                Рік = one + two + three + four + five + six + seven + eight + nine + ten + eleven + twelve;
                Січень = one;
                Лютий = two;
                Березень = three;
                Квітень = four;
                Травень = five;
                Червень = six;
                Липень = seven;
                Серпень = eight;
                Вересень = nine;
                Жовтень = ten;
                Листопад = eleven;
                Грудень = twelve;

                for (int i = 0; i < Month; i++)
                {
                    Період += (double)GetType().GetProperty(Func.names_months[i].ToString()).GetValue(this);
                }
            }
            else
            {
                Рік = one + two + three + four + five + six + seven + eight + nine + ten + eleven + twelve;
                Січень = one;
                Лютий = two;
                Березень = three;
                Квітень = four;
                Травень = five;
                Червень = six;
                Липень = seven;
                Серпень = eight;
                Вересень = nine;
                Жовтень = ten;
                Листопад = eleven;
                Грудень = twelve;

                for (int i = 0; i < Month; i++)
                {
                    Період += (double)GetType().GetProperty(Func.names_months[i].ToString()).GetValue(this);
                }
            }
        }

        public long Id { get; set; }
        public Foundation Фонд { get; set; }
        public MicroFoundation Мікрофонд { get; set; }
        public KFK КФК { get; set; }
        public Main_manager Головний_розпорядник { get; set; }
        public KEKB КЕКВ { get; set; }
        public string Дані { get; set; }
        public double Рік { get; set; }
        public double Період { get; set; }
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

    #region "Converters"

    public class RemainderDigitConverter : IValueConverter
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

    public class RemainderWidthConverterForColumnHeaderOfGroupFirst : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (170.0 / 1880.0) * (double)value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    public class RemainderWidthConverterForColumnHeaderOfGroupSecond : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (135.0 / 1880.0) * (double)value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    public class RemainderWidthConverterForColumnHeaderOfGroupMonths : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (120.0 / 1880.0) * (double)value;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }

    public class RemainderConverterForNameOfGroup : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string s = "";

            while (value != null && value.GetType().GetProperty("Name").GetValue(value).ToString() != "Root")
            {
                s = $" [{value.GetType().GetProperty("Name").GetValue(value).ToString()}] " + s;
                value = ((PropertyInfo[])((TypeInfo)value.GetType()).DeclaredProperties).FirstOrDefault(w => w.Name == "Parent")?.GetValue(value);
            }

            return s;
        }
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return null;
        }
    }


    public class AccumulativelyCurrBalanceGroupTotalConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Рік).Sum() -
                               ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Рік).Sum() -
                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupCurrentConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Період).Sum() -
                               ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Період).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Період).Sum() -
                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Період).Sum();
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupOneConverter : IValueConverter
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
            string month = "Січень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupTwoConverter : IValueConverter
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
            string month = "Лютий";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupThreeConverter : IValueConverter
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
            string month = "Березень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupFourConverter : IValueConverter
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
            string month = "Квітень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupFiveConverter : IValueConverter
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
            string month = "Травень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupSixConverter : IValueConverter
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
            string month = "Червень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupSevenConverter : IValueConverter
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
            string month = "Липень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupEightConverter : IValueConverter
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
            string month = "Серпень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupNineConverter : IValueConverter
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
            string month = "Вересень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupTenConverter : IValueConverter
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
            string month = "Жовтень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupElevenConverter : IValueConverter
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
            string month = "Листопад";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }
    public class AccumulativelyCurrBalanceGroupTwelveConverter : IValueConverter
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
            string month = "Грудень";
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
                        for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                        {
                            sum += (((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                            .Sum() -

                                   ((CollectionViewGroup)items[i]).Items
                                        .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                        .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                        .Sum());
                        }
                    }
                }
                return sum;
            }
            else
            {
                for (int m = 0; m <= Func.names_months.IndexOf(month); m++)
                {
                    sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                    .Sum() -

                           items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                .Select(s => (double)s.GetType().GetProperty(Func.names_months[m]).GetValue(s))
                                .Sum());
                }
                return sum;
            }
        }
    }

    public class MonthCurrBalanceGroupTotalConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Рік).Sum() -
                               ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Рік).Sum() -
                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupCurrentConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Період).Sum() -
                               ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Період).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Період).Sum() -
                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Період).Sum();
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupOneConverter : IValueConverter
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
            string month = "Січень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupTwoConverter : IValueConverter
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
            string month = "Лютий";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupThreeConverter : IValueConverter
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
            string month = "Березень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupFourConverter : IValueConverter
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
            string month = "Квітень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupFiveConverter : IValueConverter
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
            string month = "Травень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupSixConverter : IValueConverter
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
            string month = "Червень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupSevenConverter : IValueConverter
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
            string month = "Липень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupEightConverter : IValueConverter
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
            string month = "Серпень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupNineConverter : IValueConverter
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
            string month = "Вересень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupTenConverter : IValueConverter
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
            string month = "Жовтень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupElevenConverter : IValueConverter
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
            string month = "Листопад";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }
    public class MonthCurrBalanceGroupTwelveConverter : IValueConverter
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
            string month = "Грудень";
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                        .Sum() -

                               ((CollectionViewGroup)items[i]).Items
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                                .Sum() -

                       items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => (double)s.GetType().GetProperty(month).GetValue(s))
                            .Sum());
                return sum;
            }
        }
    }

    public class CurrPlanGroupTotalConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                return sum;
            }
        }
    }
    public class CurrPlanGroupCurrentConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Період).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "План").Select(s => ((CurrPlanEntity)s).Період).Sum();
                return sum;
            }
        }
    }
    public class CurrPlanGroupOneConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupTwoConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupThreeConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupFourConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupFiveConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupSixConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupSevenConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupEightConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupNineConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupTenConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupElevenConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPlanGroupTwelveConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "План")
                                    .Select(s => ((CurrPlanEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "План")
                            .Select(s => ((CurrPlanEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }

    public class CurrPayGroupTotalConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Рік).Sum();
                return sum;
            }
        }
    }
    public class CurrPayGroupCurrentConverter : IValueConverter
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
                        sum += ((CollectionViewGroup)items[i]).Items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Період).Sum();
                    }
                }
                return sum;
            }
            else
            {
                sum += items.Where(w => ((CurrPlanEntity)w).Дані == "Факт").Select(s => ((CurrPlanEntity)s).Період).Sum();
                return sum;
            }
        }
    }
    public class CurrPayGroupOneConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Січень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Січень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupTwoConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Лютий)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Лютий)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupThreeConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Березень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Березень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupFourConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Квітень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Квітень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupFiveConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Травень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Травень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupSixConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Червень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Червень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupSevenConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Липень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Липень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupEightConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Серпень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Серпень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupNineConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Вересень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Вересень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupTenConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Жовтень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Жовтень)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupElevenConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Листопад)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Листопад)
                            .Sum());
                return sum;
            }
        }
    }
    public class CurrPayGroupTwelveConverter : IValueConverter
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
                                    .Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                                    .Select(s => ((CurrPlanEntity)s).Грудень)
                                    .Sum());
                    }
                }
                return sum;
            }
            else
            {
                sum += (items.Where(w => ((CurrPlanEntity)w).Дані == "Факт")
                            .Select(s => ((CurrPlanEntity)s).Грудень)
                            .Sum());
                return sum;
            }
        }
    }

    #endregion
}
