using DBSolom;
using ClosedXML.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Entity;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Data;
using System.Collections.ObjectModel;

namespace Main
{
    public static class Func
    {
        public static string Login { get; set; }

        public static readonly List<string> names_months = new List<string>() { "Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень", "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень" };

        static readonly DBSolom.Db db = new Db(GetConnectionString);

        private static string ConnectionString { get; set; }

        /// <summary>
        /// Предоставляет строку подключения к базе данных в зависимости от того в каком состоянии проект - дебагинг/релиз
        /// </summary>
        public static string GetConnectionString
        {
            get
            {
                string resultString = "";
                if (ConnectionString is null || ConnectionString == "")
                {
#if DEBUG
                    var file = System.IO.File.OpenText(Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.LastIndexOf("Main")) + "Main\\Connection.imperiod");
#else
                    var file = System.IO.File.OpenText(Environment.CurrentDirectory + "\\Connection.imperiod");
#endif
                    resultString = file.ReadLine();
                    file.Close();
                }
                else
                {
                    resultString = ConnectionString;
                }
                return resultString;
            }
        }

        /// <summary>
        /// Генерирует столбцы для всех DataGrid
        /// </summary>
        /// <param name="counterForDGMColumns">Счётчик</param>
        /// <param name="e">Стандартный аргумент события</param>
        static public void GenerateColumnForDataGrid(DBSolom.Db db, ref int counterForDGMColumns, DataGridAutoGeneratingColumnEventArgs e)
        {
            CultureInfo cultureInfo = new CultureInfo("ru-RU", true);
            string headerString = e.Column.Header.ToString();
            object header = e.Column.Header;

            switch (headerString)
            {
                case "Id":
                case "SortIndex":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = header,
                        Visibility = Visibility.Hidden,
                        Binding = new Binding(headerString) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Видалено":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = header,
                        Visibility = Visibility.Hidden,
                        Binding = new Binding(headerString) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Створив":
                case "Змінив":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = header,
                        Visibility = Visibility.Hidden,

                        ItemsSource = db.Users
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Логін).ToList(),

                        DisplayMemberPath = "Логін",
                        SelectedValueBinding = new Binding(headerString) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = headerString + ".Логін"
                    };
                    break;
                case "Створино":
                case "Змінено":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = header,
                        Visibility = Visibility.Hidden,
                        Binding = new Binding(headerString) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged, StringFormat = "dd.MM.yyyy HH:mm" },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Правовласник":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = header,

                        ItemsSource = db.Users
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Логін).ToList(),

                        DisplayMemberPath = "Логін",
                        SelectedValueBinding = new Binding(headerString) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = false,
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = headerString + ".Логін"
                    };
                    break;
                case "Контакти":
                case "Логін":
                case "Тип":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding(headerString) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Проведено":
                    #region "DatePicker"

                    Binding dateBind = new Binding(headerString) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged, StringFormat = "dd.MM.yyyy" };

                    FrameworkElementFactory datePickerFactoryElem = new FrameworkElementFactory(typeof(DatePicker));
                    datePickerFactoryElem.SetValue(DatePicker.SelectedDateProperty, dateBind);
                    datePickerFactoryElem.SetValue(DatePicker.DisplayDateProperty, dateBind);

                    FrameworkElementFactory frameworkElementFactory = new FrameworkElementFactory(typeof(TextBlock));
                    frameworkElementFactory.SetValue(TextBlock.TextProperty, dateBind);

                    DataTemplate cellTemplate = new DataTemplate() { VisualTree = datePickerFactoryElem };
                    DataTemplate dataTemplate = new DataTemplate() { VisualTree = frameworkElementFactory };

                    DataGridTemplateColumn templateColumn = new DataGridTemplateColumn()
                    {
                        Header = header,
                        CellEditingTemplate = cellTemplate,
                        CellTemplate = dataTemplate,
                        DisplayIndex = counterForDGMColumns
                    };

                    e.Column = templateColumn;

                    #endregion
                    break;
                case "Підписано":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = header,
                        Binding = new Binding(headerString) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Внутрішній_номер":
                case "Підстава":
                case "Повністю":
                case "Найменування":
                case "Код":
                case "КПОЛ":
                case "Код_ГУДКСУ":
                case "Код_УДКСУ":
                case "ЕГРПОУ":
                case "Рівень_розпорядника":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = header,
                        Binding = new Binding(headerString) { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Статус":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = header,
                        Width = DataGridLength.Auto,

                        ItemsSource = db.DocStatuses
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Повністю).ToList(),

                        DisplayMemberPath = "Повністю",
                        SelectedValueBinding = new Binding(headerString) { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = headerString + ".Повністю"
                    };
                    break;
                case "Розпорядник":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,

                        ItemsSource = db.Managers
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Include(i => i.Головний_розпорядник)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Найменування).ToList(),

                        DisplayMemberPath = "Найменування",
                        SelectedValueBinding = new Binding("Розпорядник") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Розпорядник.Найменування"
                    };
                    break;
                case "Головний_розпорядник":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.Main_Managers
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Найменування).ToList(),

                        DisplayMemberPath = "Найменування",
                        SelectedValueBinding = new Binding("Головний_розпорядник") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Головний_розпорядник.Найменування"
                    };
                    break;
                case "КФК":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.KFKs
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Код).ToList(),

                        DisplayMemberPath = "Код",
                        SelectedValueBinding = new Binding("КФК") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "КФК.Код"
                    };
                    break;
                case "Макрофонд":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,

                        ItemsSource = db.MacroFoundations
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Повністю).ToList(),

                        DisplayMemberPath = "Повністю",
                        SelectedValueBinding = new Binding("Макрофонд") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Макрофонд.Повністю"
                    };
                    break;
                case "Фонд":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.Foundations
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Include(i => i.Макрофонд)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Повністю).ToList(),

                        DisplayMemberPath = "Код",
                        SelectedValueBinding = new Binding("Фонд") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Фонд.Код"
                    };
                    break;
                case "Мікрофонд":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.MicroFoundations
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Include(i => i.Фонд)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Повністю).ToList(),

                        DisplayMemberPath = "Повністю",
                        SelectedValueBinding = new Binding("Мікрофонд") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Мікрофонд.Повністю"
                    };
                    break;
                case "КФБ":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.KFBs
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Код).ToList(),

                        DisplayMemberPath = "Код",
                        SelectedValueBinding = new Binding("КФБ") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "КФБ.Код"
                    };
                    break;
                case "КДБ":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.KDBs
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Код).ToList(),

                        DisplayMemberPath = "Код",
                        SelectedValueBinding = new Binding("КДБ") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "КДБ.Код"
                    };
                    break;
                case "КЕКВ":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = db.KEKBs
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Код).ToList(),

                        DisplayMemberPath = "Код",
                        SelectedValueBinding = new Binding("КЕКВ") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "КЕКВ.Код"
                    };
                    break;
                case "Дані":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = DataGridLength.Auto,
                        ItemsSource = new List<string> { "План", "Факт", "Н_Залишок", "М_Залишок" },
                        SelectedValueBinding = new Binding("Дані")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Значення":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = 80,
                        ItemsSource = Func.names_months.Concat(new List<string> { "Рік", "Період" }),
                        SelectedValueBinding = new Binding("Значення")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Операція":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = 70,
                        ItemsSource = new List<string> { "==", "!=", ">=", "<=", ">", "<", "+", "-", "/", "*" },
                        SelectedValueBinding = new Binding("Операція")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Рік":
                case "Період":
                case "Січень":
                case "Лютий":
                case "Березень":
                case "Квітень":
                case "Травень":
                case "Червень":
                case "Липень":
                case "Серпень":
                case "Вересень":
                case "Жовтень":
                case "Листопад":
                case "Грудень":
                case "Сума":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = header,
                        Width = DataGridLength.Auto,
                        Binding = new Binding(headerString)
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Уточнений_план":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = header,
                        IsReadOnly = true,
                        Binding = new Binding(headerString)
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "План":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        IsReadOnly = true,
                        Binding = new Binding("План")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Профінансовано":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        IsReadOnly = true,
                        Binding = new Binding("Профінансовано")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Залишок":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        IsReadOnly = true,
                        Binding = new Binding("Залишок")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "DocStatus":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("DocStatus") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Macrofoundation":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Macrofoundation") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Foundation":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Foundation") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Microfoundation":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Microfoundation") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "KDB":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("KDB") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "KEKB":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("KEKB") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "KFK":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("KFK") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Main_manager":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Main_manager") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Manager":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Manager") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Correction":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Correction") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Filling":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Filling") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Microfilling":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Microfilling") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Financing":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Financing") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "User":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("User") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Lowt":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Lowt") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                default:
                    e.Cancel = true;
                    counterForDGMColumns--;
                    break;
            }
            counterForDGMColumns++;
        }

        /// <summary>
        /// Создаёт представление подменю "Фильтр"
        /// </summary>
        /// <param name="EXPGRO">Грид экспендера тот что подменю "Фильтр"</param>
        /// <param name="t">Счётчик</param>
        /// <param name="item">Сущность с именами столбцов</param>
        /// <param name="dict_cmb">Словарь с ComboBox - тип сравнение</param>
        /// <param name="dict_txt">Словарь с TextBox - сравнитель</param>
        /// <param name="GetLabels">Список labels</param>
        public static void GetFilters(Grid EXPGRO, int t, ItemPropertyInfo item,
            ref Dictionary<string, ComboBox> dict_cmb, ref Dictionary<string, TextBox> dict_txt, ref List<Label> GetLabels)
        {
            //Filter name
            Label label = new Label()
            {
                Content = item.Name,
                Margin = new Thickness(2),
                HorizontalContentAlignment = HorizontalAlignment.Stretch
            };
            Grid.SetRow(label, 0);
            Grid.SetColumn(label, t);

            GetLabels.Add(label);

            //Filter type
            ComboBox comboBox = new ComboBox()
            {
                ItemsSource = db.list,
                Margin = new Thickness(2),
                HorizontalContentAlignment = HorizontalAlignment.Stretch
            };
            Grid.SetRow(comboBox, 1);
            Grid.SetColumn(comboBox, t);

            dict_cmb.Add(item.Name, comboBox);

            //Filter value
            TextBox textBox = new TextBox()
            {
                Margin = new Thickness(2),
                HorizontalContentAlignment = HorizontalAlignment.Stretch
            };
            Grid.SetRow(textBox, 2);
            Grid.SetColumn(textBox, t);

            dict_txt.Add(item.Name, textBox);

            //Add filters
            EXPGRO.Children.Add(label);
            EXPGRO.Children.Add(comboBox);
            EXPGRO.Children.Add(textBox);
        }

        /// <summary>
        /// Creat view for context menu and attach his to top grid of context menu
        /// </summary>
        /// <param name="topGrid">Top grid of context menu</param>
        /// <param name="correctionCntxMenuList">List of correction context menu entity</param>
        public static void GetContextMenuView(Grid topGrid, List<CorrectionCntxMenu> correctionCntxMenuList)
        {
            topGrid.Children.Clear();
            foreach (var correctionCntxMenu in correctionCntxMenuList)
            {
                Grid grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });

                Grid.SetRow(grid, correctionCntxMenuList.IndexOf(correctionCntxMenu));
                Grid.SetColumn(grid, 0);

                foreach (var lbl in correctionCntxMenu.cntx_dict_lbl)
                {
                    int indexColumn = correctionCntxMenu.cntx_dict_lbl.IndexOf(lbl);

                    Grid.SetRow(lbl, 0);
                    Grid.SetColumn(lbl, indexColumn);
                    grid.Children.Add(lbl);
                }
                foreach (var cmb in correctionCntxMenu.cntx_dict_cmb)
                {
                    int indexColumn = correctionCntxMenu.cntx_dict_cmb.IndexOf(cmb);

                    Grid.SetRow(cmb, 1);
                    Grid.SetColumn(cmb, indexColumn);
                    grid.Children.Add(cmb);
                }
                foreach (var txt in correctionCntxMenu.cntx_dict_txt)
                {
                    int indexColumn = correctionCntxMenu.cntx_dict_txt.IndexOf(txt);

                    Grid.SetRow(txt, 2);
                    Grid.SetColumn(txt, indexColumn);
                    grid.Children.Add(txt);
                }
                if (grid.Children.Count > 0)
                {
                    topGrid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
                    topGrid.Children.Add(grid);
                }
            }

            Button button = new Button() { Content = "Консолідувати" };
            Grid.SetRow(button, topGrid.RowDefinitions.Count);
            Grid.SetColumn(button, 0);
            button.Click += Maestro_T_Btn_Click;

            topGrid.RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
            topGrid.Children.Add(button);

            void Maestro_T_Btn_Click(object sender, RoutedEventArgs e)
            {
                Maestro.Functional.Summary summary = new Maestro.Functional.Summary();
                summary.Dispatcher.Invoke(new Action(() =>
                {
                    summary.FillDate(DateTime.Now);
                }));

                foreach (var correctionCntxMenu in correctionCntxMenuList)
                {
                    if (correctionCntxMenu.cntx_dict_cmb[0].SelectedValue != null)
                    {
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_cmb["Фонд"].SelectedValue = correctionCntxMenu.cntx_dict_cmb[0].SelectedValue;
                        }));
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_txt["Фонд"].Text = correctionCntxMenu.cntx_dict_txt[0].Text;
                        }));
                    }

                    if (correctionCntxMenu.cntx_dict_cmb[1].SelectedValue != null)
                    {
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_cmb["Головний_розпорядник"].SelectedValue = correctionCntxMenu.cntx_dict_cmb[1].SelectedValue;
                        }));
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_txt["Головний_розпорядник"].Text = correctionCntxMenu.cntx_dict_txt[1].Text;
                        }));
                    }

                    if (correctionCntxMenu.cntx_dict_cmb[2].SelectedValue != null)
                    {
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_cmb["КФК"].SelectedValue = correctionCntxMenu.cntx_dict_cmb[2].SelectedValue;
                        }));
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_txt["КФК"].Text = correctionCntxMenu.cntx_dict_txt[2].Text;
                        }));
                    }

                    if (correctionCntxMenu.cntx_dict_cmb[3].SelectedValue != null)
                    {
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_cmb["КЕКВ"].SelectedValue = correctionCntxMenu.cntx_dict_cmb[3].SelectedValue;
                        }));
                        summary.Dispatcher.Invoke(new Action(() =>
                        {
                            summary.dict_txt["КЕКВ"].Text = correctionCntxMenu.cntx_dict_txt[3].Text;
                        }));
                    }

                    summary.Dispatcher.Invoke(new Action(() =>
                    {
                        summary.BTN_Accept_Click(null, null);
                    }));
                }
                summary.Show();
            }
        }

        /// <summary>
        /// Обработчик группировки
        /// </summary>
        /// <param name="sender">Стандартный аргумент события</param>
        /// <param name="e">Стандартный аргумент события</param>
        public static void GroupButton_Click(object sender, RoutedEventArgs e)
        {
            var query = ((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)e.OriginalSource).Parent).Parent).Parent).Parent).Parent).Parent).Parent.GetType().GetRuntimeFields();
            DataGrid DGM = (DataGrid)query.First(w => w.Name == "DGM").GetValue(((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)e.OriginalSource).Parent).Parent).Parent).Parent).Parent).Parent).Parent);

            object q = null;
            List<string> x = new List<string>();
            bool zeta = false;

            ICollectionView cvTasks = CollectionViewSource.GetDefaultView(DGM.ItemsSource);

            if (cvTasks != null && cvTasks.CanGroup == true)
            {
                if (((ToggleButton)sender).IsChecked.Value)
                {
                    try
                    {
                        zeta = ((ListCollectionView)cvTasks).ItemProperties.FirstOrDefault(k => k.Name == ((ToggleButton)sender).Content.ToString()).PropertyType.FullName.Contains("DBSolom");
                        if (zeta)
                        {
                            q = ((ListCollectionView)cvTasks).ItemProperties.FirstOrDefault(k => k.Name == ((ToggleButton)sender).Content.ToString());

                            x = ((TypeInfo)((ItemPropertyInfo)q).PropertyType).DeclaredProperties.Select(k => k.Name).ToList();
                            if (x.Contains("Код"))
                            {
                                cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + "." + "Код"));
                            }
                            else if (x.Contains("Найменування"))
                            {
                                cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + "." + "Найменування"));
                            }
                            else if (x.Contains("Повністю"))
                            {
                                cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + "." + "Повністю"));
                            }
                            else if (x.Contains("Логін"))
                            {
                                cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + "." + "Логін"));
                            }
                        }
                        else
                        {

                            if (DGM.Items[0].GetType().GetProperty(((ToggleButton)sender).Content.ToString()).PropertyType == typeof(DateTime))
                            {
                                switch (Microsoft.VisualBasic.Interaction.InputBox("Групувати за: \nДатою = 0\nРоком = 1\nМісяцем = 2\nДнем = 3"))
                                {
                                    case "0":
                                        cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Date"));
                                        break;
                                    case "1":
                                        cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Year"));
                                        break;
                                    case "2":
                                        cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Month"));
                                        break;
                                    case "3":
                                        cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Day"));
                                        break;
                                    default:
                                        MessageBox.Show("Групування не виконано! Хибне значення! Спробуйте ще.");
                                        break;
                                }
                            }
                            else if (((ToggleButton)sender).Content.ToString() == "Фонд_Мікрофонд")
                            {
                                cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Повністю"));
                            }
                            else
                            {
                                cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString()));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else if (DGM.Items[0].GetType().GetProperty(((ToggleButton)sender).Content.ToString()).PropertyType == typeof(DateTime))
                {
                    switch (Microsoft.VisualBasic.Interaction.InputBox("Групувати за: \nДатою = 0\nРоком = 1\nМісяцем = 2\nДнем = 3"))
                    {
                        case "0":
                            cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Date"));
                            break;
                        case "1":
                            cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Year"));
                            break;
                        case "2":
                            cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Month"));
                            break;
                        case "3":
                            cvTasks.GroupDescriptions.Add(new PropertyGroupDescription(((ToggleButton)sender).Content.ToString() + ".Day"));
                            break;
                        default:
                            MessageBox.Show("Групування не виконано! Хибне значення! Спробуйте ще.");
                            break;
                    }
                }
                else
                {
                    try
                    {
                        cvTasks.GroupDescriptions.Remove(cvTasks.GroupDescriptions.FirstOrDefault(k => ((PropertyGroupDescription)k).PropertyName.Contains(((ToggleButton)sender).Content.ToString())));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            try
            {
                cvTasks.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Создаёт представление подменю "Группы"
        /// </summary>
        /// <param name="t">Счётчик</param>
        /// <param name="item">Сущность с именами столбцов</param>
        /// <param name="CheckBoxes">Список togglebuttons</param>
        /// <param name="EXPGRT">Грид экспендера тот что подменю "Группы"</param>
        public static void GetGroups(int t, ItemPropertyInfo item, ref List<ToggleButton> CheckBoxes, ref Grid EXPGRT)
        {
            var w = ((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)EXPGRT.Parent).Parent).Parent).Parent).Parent).Parent;
            Style st = (Style)((FrameworkElement)w).Resources["Style"];

            ToggleButton toggleButton = new ToggleButton()
            {
                Content = item.Name,
                IsThreeState = false,
                Style = st,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                IsChecked = false
            };

            Grid.SetColumn(toggleButton, t);

            toggleButton.Checked += GroupButton_Click;
            toggleButton.Unchecked += GroupButton_Click;
            CheckBoxes.Add(toggleButton);
            EXPGRT.Children.Add(toggleButton);
        }

        /// <summary>
        /// Создаёт представление подменю "Видимость"
        /// </summary>
        /// <param name="t">Счётчик</param>
        /// <param name="item">Сущность с именами столбцов</param>
        /// <param name="EXPHDN">Грид экспендера(Тот что меню "Видимость")</param>
        public static void GetVisibilityOfColumns(int t, ItemPropertyInfo item, ref Grid EXPHDN)
        {
            var w = ((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)EXPHDN.Parent).Parent).Parent).Parent).Parent).Parent;
            Style st = (Style)((FrameworkElement)w).Resources["Style"];

            ToggleButton toggleButton = new ToggleButton()
            {
                Content = item.Name,
                IsThreeState = false,
                IsChecked = new List<string>() { "SortIndex", "Id", "Видалено", "Створив", "Створино", "Змінив", "Змінено" }.Contains(item.Name) ? false : true,
                Style = st,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };

            Grid.SetColumn(toggleButton, t);

            toggleButton.Checked += HiddenUnhiddenColumn;
            toggleButton.Unchecked += HiddenUnhiddenColumn;
            EXPHDN.Children.Add(toggleButton);
        }

        /// <summary>
        /// Скрывает или показывает столбец
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void HiddenUnhiddenColumn(object sender, RoutedEventArgs e)
        {
            var query = ((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)e.OriginalSource).Parent).Parent).Parent).Parent).Parent).Parent).Parent.GetType().GetRuntimeFields();
            DataGrid DGM = (DataGrid)query.First(w => w.Name == "DGM").GetValue(((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)e.OriginalSource).Parent).Parent).Parent).Parent).Parent).Parent).Parent);

            var o = DGM.Columns.FirstOrDefault(w => w.Header.ToString() == ((ToggleButton)sender).Content.ToString());
            int i = o.DisplayIndex;
            if (((ToggleButton)sender).IsChecked.Value)
            {
                DGM.Columns[i].Visibility = Visibility.Visible;
            }
            else
            {
                DGM.Columns[i].Visibility = Visibility.Hidden;
            }
        }

        /// <summary>
        /// Унифицированый фильтр сущностей
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void CollectionView_Filter(object sender, FilterEventArgs e)
        {
            Window active_window = (Window)((TypeInfo)sender.GetType()).DeclaredProperties.FirstOrDefault(f => f.Name == "InheritanceContext").GetValue(sender);
            List<Filters> filters = (List<Filters>)((TypeInfo)active_window.GetType()).DeclaredFields.First(f => f.Name == "GetFilters").GetValue(active_window);

            if (filters.Count > 0)
            {
                if ((e.Item.GetType().GetProperties().FirstOrDefault(f => f.Name == "Id") is null) || ((e.Item.GetType().GetProperties().FirstOrDefault(f => f.Name == "Id") != null) && (long)e.Item.GetType().GetProperty("Id").GetValue(e.Item) != 0))
                {
                    List<bool> resultList = new List<bool>();
                    try
                    {
                        foreach (Filters item in filters) //Перебор всех фильтров по типу ИЛИ
                        {
                            resultList.Add(CheckOneEntity(item));
                        }
                        e.Accepted = resultList.Where(w => w == true).Count() > 0 ? true : false;
                    }
                    catch
                    {
                        e.Accepted = false;
                    }
                }
            }

            bool CheckOneEntity(Filters filter)
            {
                object OriginalValue = null;
                dynamic RealValue = null;
                string PropertyName = null;
                string typeValue = null;
                List<bool> resultOfEquels = new List<bool>();

                foreach (var micro_item in filter.GetFilters) //Перебор всех фильтров по типу И (если хоть один не проходит тогда False)
                {
                    //Определение главных переменных - оригинальное значение, тип значения и сравниваемое значение
                    if (e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item) is null)
                    {
                        typeValue = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item)?.GetType().Name;
                        RealValue = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item);
                    }
                    else if (e.Item.GetType().GetProperty(micro_item["prop"]).PropertyType.FullName.Contains("DBSolom"))
                    {

                        var ListPropertysOfEntity = ((PropertyInfo[])e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item).GetType().GetProperties()).Select(k => k.Name).ToList();
                        OriginalValue = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item);

                        if (ListPropertysOfEntity.Contains("Код"))
                        {
                            PropertyName = "Код";
                        }
                        else if (ListPropertysOfEntity.Contains("Найменування"))
                        {
                            PropertyName = "Найменування";
                        }
                        else if (ListPropertysOfEntity.Contains("Повністю"))
                        {
                            PropertyName = "Повністю";
                        }
                        else if (ListPropertysOfEntity.Contains("Логін"))
                        {
                            PropertyName = "Логін";
                        }

                        if (OriginalValue.GetType().GetProperty(PropertyName).GetValue(OriginalValue) is null)
                        {
                            return false;
                        }

                        RealValue = OriginalValue.GetType().GetProperty(PropertyName).GetValue(OriginalValue);
                        typeValue = RealValue.GetType().Name;
                    }
                    else
                    {
                        typeValue = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item).GetType().Name;
                        RealValue = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item);
                    }


                    //Ветвление по операциям сравнения
                    if (micro_item["type"] == "-")
                    {
                        bool switcher = true;
                        string start = "";
                        string end = "";

                        for (int i = 0; i < micro_item["value"].Length; i++)
                        {
                            if (micro_item["value"][i] != '-')
                            {
                                if (switcher)
                                {
                                    start += micro_item["value"][i];
                                }
                                else
                                {
                                    end += micro_item["value"][i];
                                }
                            }
                            else
                            {
                                switcher = false;
                            }
                        }

                        if (typeValue == "String")
                        {
                            resultOfEquels.Add(RemoveBadSymbols(RealValue.ToString()).Length >= RemoveBadSymbols(start).Length &&
                                               RemoveBadSymbols(RealValue.ToString()).Length <= RemoveBadSymbols(end).Length);
                        }
                        else
                        {
                            resultOfEquels.Add(Tech.CodeGeneration.CodeGenerator.ExecuteCode<bool>(
                                $"return {typeValue}.Parse(RealValue) >= {typeValue}.Parse(FirstFilterValue) && {typeValue}.Parse(RealValue) <= {typeValue}.Parse(SecondFilterValue);",
                                                Tech.CodeGeneration.CodeParameter.Create("FirstFilterValue", RemoveBadSymbols(start)),
                                                Tech.CodeGeneration.CodeParameter.Create("SecondFilterValue", RemoveBadSymbols(end)),
                                                Tech.CodeGeneration.CodeParameter.Create("RealValue", RemoveBadSymbols(RealValue.ToString()))));
                        }
                    }
                    else if (micro_item["type"] == "[,]")
                    {
                        List<string> list = new List<string>();
                        string temp = "";

                        #region "FillList"
                        for (int i = 0; i < micro_item["value"].Length; i++)
                        {
                            if (micro_item["value"][i] != ',')
                            {
                                temp += micro_item["value"][i];
                                if (i == (micro_item["value"].Length - 1))
                                {
                                    list.Add(RemoveBadSymbols(temp));
                                }
                            }
                            else
                            {
                                list.Add(RemoveBadSymbols(temp));
                                temp = "";
                            }
                        }
                        #endregion
                        resultOfEquels.Add(list.Contains(RemoveBadSymbols(RealValue.ToString()).ToLower()));
                    }
                    else if (micro_item["type"] == ">|<")
                    {
                        resultOfEquels.Add(RemoveBadSymbols(RealValue.ToString()).ToLower()
                                          .Contains(RemoveBadSymbols(micro_item["value"].ToString()).ToLower()));
                    }
                    else
                    {
                        if (typeValue == "String")
                        {
                            resultOfEquels.Add(Tech.CodeGeneration.CodeGenerator.ExecuteCode<bool>(
                                $"return RealValue {micro_item["type"]} FilterValue ;",
                                Tech.CodeGeneration.CodeParameter.Create("FilterValue", RemoveBadSymbols(micro_item["value"].ToString())),
                                Tech.CodeGeneration.CodeParameter.Create("RealValue", RemoveBadSymbols(RealValue.ToString()))));
                        }
                        else if (typeValue is null && micro_item["value"] == "null")
                        {
                            resultOfEquels.Add(true);
                        }
                        else
                        {
                            resultOfEquels.Add(Tech.CodeGeneration.CodeGenerator.ExecuteCode<bool>(
                                    $"return {typeValue}.Parse(RealValue) {micro_item["type"]} {typeValue}.Parse(FilterValue);",
                                    Tech.CodeGeneration.CodeParameter.Create("FilterValue", RemoveBadSymbols(micro_item["value"].ToString())),
                                    Tech.CodeGeneration.CodeParameter.Create("RealValue", RemoveBadSymbols(RealValue.ToString()))));
                        }
                    }
                }

                return resultOfEquels.Where(w => w == false).Count() > 0 ? false : true;
            }
        }

        /// <summary>
        /// Удаляет проблемный символы
        /// </summary>
        /// <param name="s">Строка с проблемными символами</param>
        /// <returns>Чистая строка</returns>
        private static string RemoveBadSymbols(string s)
        {
            return s.Replace("\"", "");
        }

        /// <summary>
        /// Метод экспорта данных в эксель в зависимости от выделенных ячеек в датагрид
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void BTN_ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            #region "Variables"
            Window active_window = (Window)((Grid)((Expander)((Grid)((Button)sender).Parent).Parent).Parent).Parent;
            DataGrid DGM = (DataGrid)active_window.GetType().GetRuntimeFields().First(f => f.Name == "DGM").GetValue(active_window);
            ProgressBar PB = (ProgressBar)active_window.GetType().GetRuntimeFields().First(f => f.Name == "PB").GetValue(active_window);
            #endregion

            if (DGM.SelectedCells.Count > 0)
            {
                List<Dictionary<string, dynamic>> Entities = CopyEntities();

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm;";
                if (openFileDialog.ShowDialog() == true)
                {
                    PB.IsIndeterminate = true;

                    var Task = new Task(() =>
                    {
                        try
                        {
                            using (XLWorkbook wb = new XLWorkbook(openFileDialog.FileName, XLEventTracking.Disabled))
                            {
                                IXLWorksheet ws = wb.Worksheets.FirstOrDefault(f => f.Name == "Maestro_Data");
                                wb.CalculateMode = XLCalculateMode.Manual;
                                int countHeadersOfEntities = Entities.First().Keys.Count;
                                IXLTable table;

                                if (ws is null)
                                {
                                    ws = wb.AddWorksheet("Maestro_Data");
                                    DataTable newTable = new DataTable("Maestro_Table");

                            //Headers
                            foreach (var item in Entities.First().Keys)
                                    {
                                        DataColumn column = new DataColumn(item);
                                        newTable.Columns.Add(column);
                                    }

                            //Rows
                            foreach (var columnValue in Entities)
                                    {
                                        DataRow newRow = newTable.NewRow();
                                        newTable.Rows.Add(newRow);
                                    }

                                    ws.Cell(1, 1).InsertTable(newTable, "Maestro_Table", true); //Attach table
                                    table = ws.Tables.FirstOrDefault(f => f.Name == "Maestro_Table");

                                    SetStylesForValues();
                                    FillData();
                                }
                                else //Worksheet is exist
                                {


                                    table = ws.Tables.FirstOrDefault(f => f.Name == "Maestro_Table");
                                    if (table is null)
                                    {
                                        if (ws.Tables.Count() > 0)
                                        {
                                            PB.Dispatcher.Invoke(() => PB.IsIndeterminate = false);
                                            MessageBox.Show("На листі Maestro_Data є сторонні таблиці, якщо потрібна таблиця все таки є і вона коректна, будь ласка змініть її ім\'я на \"Maestro_Table\" та спробуйте знову", "Maestro:[Експорт]", MessageBoxButton.OK, MessageBoxImage.Hand);
                                            return;
                                        }
                                        DataTable newTable = new DataTable("Maestro_Table");

                                //Headers
                                foreach (var item in Entities.First().Keys)
                                        {
                                            DataColumn column = new DataColumn(item);
                                            newTable.Columns.Add(column);
                                        }

                                //Rows
                                foreach (var columnValue in Entities)
                                        {
                                            DataRow newRow = newTable.NewRow();
                                            newTable.Rows.Add(newRow);
                                        }

                                        ws.Cell(1, 1).InsertTable(newTable, "Maestro_Table", true); //Attach table
                                        table = ws.Tables.FirstOrDefault(f => f.Name == "Maestro_Table");

                                        SetStylesForValues();
                                        FillData();
                                        ws.Range(ws.Cell(1, table.Fields.Count() + 1), ws.Cell(table.RowCount() + 90001, table.Fields.Count() + 251)).Clear(XLClearOptions.All);
                                        ws.Range(ws.Cell(table.RowCount() + 1, 1), ws.Cell(90001 - table.RowCount(), table.Fields.Count())).Clear(XLClearOptions.All);
                                    }
                                    else //Table is exist
                                    {
                                        if (ws.Tables.Count() > 1)
                                        {
                                            PB.Dispatcher.Invoke(() => PB.IsIndeterminate = false);
                                            MessageBox.Show("На листі Maestro_Data є сторонні таблиці, будь ласка видаліть всі зайві таблиці за виключенням \"Maestro_Table\" та спробуйте знову", "Maestro:[Експорт]", MessageBoxButton.OK, MessageBoxImage.Hand);
                                            return;
                                        }
                                //Clear other values
                                ws.Range(ws.Cell(1, table.Fields.Count() + 1), ws.Cell(table.RowCount() + 90001, table.Fields.Count() + 251)).Clear(XLClearOptions.All);
                                        ws.Range(ws.Cell(table.RowCount() + 1, 1), ws.Cell(90001 - table.RowCount(), table.Fields.Count())).Clear(XLClearOptions.All);

                                        table.Resize(ws.Cell(1, 1), ws.Cell(Entities.Count + 1, countHeadersOfEntities));

                                //Headers
                                PrepareHeaders();
                                        for (int indexColumn = 0; indexColumn < countHeadersOfEntities; indexColumn++)
                                        {
                                            ws.Cell(1, indexColumn + 1).SetValue(Entities.First().Keys.ElementAt(indexColumn));
                                        }

                                        SetStylesForValues();
                                        FillData();
                                        ws.Range(ws.Cell(1, table.Fields.Count() + 1), ws.Cell(table.RowCount() + 90001, table.Fields.Count() + 251)).Clear(XLClearOptions.All);
                                        ws.Range(ws.Cell(table.RowCount() + 1, 1), ws.Cell(90001 - table.RowCount(), table.Fields.Count())).Clear(XLClearOptions.All);
                                    }
                                }

                                wb.CalculateMode = XLCalculateMode.Auto;
                                wb.Save();
                                MessageBox.Show("Виконано!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Information);

                                void PrepareHeaders()
                                {
                                    List<string> l = table.Fields.Select(s => s.Name).ToList();
                                    int i = 1;

                                    while (i <= l.Count)
                                    {
                                        string s = "MSTR" + new Random().Next(1, 1000);
                                        if (l.Contains(s) == false)
                                        {
                                            table.Column(i).SetValue(s);
                                            l[i - 1] = s;
                                            i++;
                                        }
                                    }
                                }
                                void SetStylesForValues()
                                {
                                    foreach (var item in Entities.First().Keys)
                                    {
                                        switch (Entities.First()[item])
                                        {
                                            case long val:
                                                ws.Table("Maestro_Table").DataRange.Column(item).Style.NumberFormat.Format = "#";
                                                break;
                                            case int val:
                                                ws.Table("Maestro_Table").DataRange.Column(item).Style.NumberFormat.Format = "#";
                                                break;
                                            case double val:
                                                ws.Table("Maestro_Table").DataRange.Column(item).Style.NumberFormat.Format = "# ### ### ##0.00;[Red]-# ### ### ###.00";
                                                break;
                                            case DateTime val:
                                                ws.Table("Maestro_Table").DataRange.Column(item).Style.NumberFormat.Format = "hh:mm:ss dd/mm/yyyy";
                                                break;
                                            case string val:
                                                ws.Table("Maestro_Table").DataRange.Column(item).Style.NumberFormat.Format = "@";
                                                break;
                                            default:
                                                break;
                                        }

                                    }
                                }
                                void FillData()
                                {
                                    for (int indexColumn = 0; indexColumn < countHeadersOfEntities; indexColumn++)
                                    {
                                        for (int indexRow = 0; indexRow < Entities.Count; indexRow++)
                                        {
                                            var v = Entities.ElementAt(indexRow).Values.ElementAt(indexColumn);
                                            table.DataRange.Cell(indexRow + 1, indexColumn + 1).SetValue(v ?? "");
                                        }
                                        ws.Column(indexColumn + 1).AdjustToContents();
                                    }
                                }
                            }
                            openFileDialog = null;
                        }
                        catch (System.IO.IOException ex)
                        {
                            MessageBox.Show($"Файл відкритий в іншій програмі або зайнятий іншим процесом.\n\nСистемне повідомлення: [{ex}]", "Maestro:[Експорт]", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        finally
                        {
                            PB.Dispatcher.Invoke(() => PB.IsIndeterminate = false);
                        }
                    });
                    Task.Start();
                }
            }
            else
            {
                MessageBox.Show("Виділіть всі строки, які будуть експортовані!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Hand);
            }

            List<Dictionary<string, dynamic>> CopyEntities()
            {
                List<Dictionary<string, dynamic>> Entities = new List<Dictionary<string, dynamic>>();

                //Copy uniques values
                IEnumerable<DataGridCellInfo> entitiesIntoCell = DGM.SelectedCells.Where(w => w.Item.ToString() != "{NewItemPlaceholder}").GroupBy((DataGridCellInfo g) => g.Item.GetType().GetProperty("Id").GetValue(g.Item)).Select(s => s.First());

                //Define headers of column
                List<string> headers = entitiesIntoCell.First().Item.GetType().GetProperties().Select(s => s.Name).ToList();

                foreach (DataGridCellInfo entityIntoCell in entitiesIntoCell)
                {
                    Dictionary<string, dynamic> dict = new Dictionary<string, dynamic>();

                    foreach (var header in headers)
                    {
                        dynamic correctValue;
                        dynamic insideProp = entityIntoCell.Item.GetType().GetProperty(header).GetValue(entityIntoCell.Item);
                        switch (header)
                        {
                            case "Створив":
                            case "Змінив":
                            case "Правовласник":
                                correctValue = insideProp?.GetType().GetProperty("Логін").GetValue(insideProp);
                                dict.Add(header, correctValue);
                                break;
                            case "Статус":
                                correctValue = insideProp?.GetType().GetProperty("Повністю").GetValue(insideProp);
                                dict.Add(header, correctValue);
                                break;
                            case "Головний_розпорядник":
                                correctValue = insideProp.GetType().GetProperty("Найменування").GetValue(insideProp);
                                dict.Add(header, correctValue);
                                break;
                            case "Мікрофонд":
                                correctValue = insideProp.GetType().GetProperty("Повністю").GetValue(insideProp);
                                dict.Add(header, correctValue);
                                if (((object)insideProp).GetType().GetProperties().Select((PropertyInfo s) => s.Name).ToList().Contains("Фонд") == false)
                                {
                                    correctValue = insideProp.GetType().GetProperty("Фонд").GetValue(insideProp);
                                    dynamic q3 = correctValue.GetType().GetProperty("Код").GetValue(correctValue);
                                    dict.Add("Фонд", q3);
                                }
                                break;
                            case "КФК":
                            case "Фонд":
                            case "КДБ":
                            case "КФБ":
                            case "КЕКВ":
                                correctValue = insideProp?.GetType().GetProperty("Код").GetValue(insideProp);
                                dict.Add(header, correctValue);
                                break;
                            default:
                                dict.Add(header, insideProp);
                                break;
                        }
                    }
                    Entities.Add(dict);
                }

                return Entities;
            }
        }

        /// <summary>
        /// Метод возвращает накопительный остаток по переданным аргументам, учитывая данные, как из БД, так и из локального хранилища
        /// </summary>
        /// <param name="db">Контекст БД</param>
        /// <param name="year">Год</param>
        /// <param name="kFK">КПБ(КФК)</param>
        /// <param name="main_Manager">Главный распорядитель</param>
        /// <param name="kEKB">КЕКВ</param>
        /// <param name="foundation">Фонд</param>
        /// <returns>Накопительный остаток</returns>
        static List<double> GetRemainderFromDBPerMonth(DBSolom.Db db, int year, KFK kFK, Main_manager main_Manager, KEKB kEKB, Foundation foundation)
        {
            DBSolom.Db mdb = new Db(GetConnectionString);

            #region "Financings"
            List<Financing> localFinancings = db.Financings.Local.Where(w =>
                                                                        w.Видалено == false &&
                                                                        w.Проведено.Year == year &&
                                                                        w.Головний_розпорядник.Найменування == main_Manager.Найменування &&
                                                                        w.КЕКВ.Код == kEKB.Код &&
                                                                        w.КФК.Код == kFK.Код &&
                                                                        w.Мікрофонд.Фонд.Код == foundation.Код)
                                                                        .ToList();

            List<Financing> DBfinancings = mdb.Financings
                              .Where(w =>
                              w.Видалено == false &&
                              w.Проведено.Year == year &&
                              w.Головний_розпорядник.Найменування == main_Manager.Найменування &&
                              w.КЕКВ.Код == kEKB.Код &&
                              w.КФК.Код == kFK.Код &&
                              w.Мікрофонд.Фонд.Код == foundation.Код)
                              .ToList();



            if (localFinancings.Count != 0)
            {
                DBfinancings = DBfinancings.Where(w => localFinancings.Select(s => s.Id).Contains(w.Id) == false).ToList();
                DBfinancings.AddRange(localFinancings);
            }
            #endregion

            var CorrectPlan = GetCurrentPlanFromDBPerMonth(db, year, kFK, main_Manager, kEKB, foundation);

            List<double> vs = new List<double>();

            //Вычисление месячных остатков
            foreach (var item in names_months)
            {
                int numberOfMonth = names_months.IndexOf(item) + 1;

                double monthPlan = CorrectPlan[numberOfMonth - 1];
                double monthFinancing = DBfinancings.Where(w => w.Проведено.Month == numberOfMonth).Sum(ss => ss.Сума);

                vs.Add(monthPlan - monthFinancing);
            }

            //Накопительно
            for (int i = 1; i < 12; i++)
            {
                vs[i] += vs[i - 1];
            }

            return vs;
        }

        /// <summary>
        /// Метод возвращает уточненный план по переданным аргументам, учитывая данные, как из БД, так и из локального хранилища
        /// </summary>
        /// <param name="db">Контекст БД</param>
        /// <param name="year">Год</param>
        /// <param name="kFK">КПБ(КФК)</param>
        /// <param name="main_Manager">Главный распорядитель</param>
        /// <param name="kEKB">КЕКВ</param>
        /// <param name="foundation">Фонд</param>
        /// <returns>Уточненный план</returns>
        static List<double> GetCurrentPlanFromDBPerMonth(DBSolom.Db db, int year, KFK kFK, Main_manager main_Manager, KEKB kEKB, Foundation foundation)
        {
            DBSolom.Db mdb = new Db(GetConnectionString);

            #region "Fillings"
            List<Filling> localFillings = db.Fillings.Local
                                                            .Where(w =>
                                                            w.Видалено == false &&
                                                            w.Проведено.Year == year &&
                                                            w.Головний_розпорядник.Найменування == main_Manager.Найменування &&
                                                            w.КЕКВ.Код == kEKB.Код &&
                                                            w.КФК.Код == kFK.Код &&
                                                            w.Фонд.Код == foundation.Код)
                                                            .ToList();

            List<Filling> DBfillings = mdb.Fillings
                                                    .Include(i => i.Фонд)
                                                    .Include(i => i.КФК)
                                                    .Include(i => i.Головний_розпорядник)
                                                    .Include(i => i.КЕКВ)
                                                    .Where(w =>
                                                    w.Проведено.Year == year &&
                                                    w.Видалено == false &&
                                                    w.Головний_розпорядник.Найменування == main_Manager.Найменування &&
                                                    w.КЕКВ.Код == kEKB.Код &&
                                                    w.КФК.Код == kFK.Код &&
                                                    w.Фонд.Код == foundation.Код)
                                                    .ToList();



            if (localFillings.Count != 0)
            {
                DBfillings = DBfillings.Where(w => localFillings.Select(s => s.Id).Contains(w.Id) == false).ToList();
                DBfillings.AddRange(localFillings);
            }

            var EndFillings = DBfillings
                .Select(s => new
                {
                    s.Фонд,
                    s.КФК,
                    s.Головний_розпорядник,
                    s.КЕКВ,
                    s.Січень,
                    s.Лютий,
                    s.Березень,
                    s.Квітень,
                    s.Травень,
                    s.Червень,
                    s.Липень,
                    s.Серпень,
                    s.Вересень,
                    s.Жовтень,
                    s.Листопад,
                    s.Грудень
                })
                .ToList();
            #endregion

            #region "Corrections"
            List<Correction> localCorrections = db.Corrections.Local.Where(w =>
                                                                        w.Видалено == false &&
                                                                        w.Проведено.Year == year &&
                                                                        w.Головний_розпорядник.Найменування == main_Manager.Найменування &&
                                                                        w.КЕКВ.Код == kEKB.Код &&
                                                                        w.КФК.Код == kFK.Код &&
                                                                        w.Мікрофонд.Фонд.Код == foundation.Код)
                                                                        .ToList();

            List<Correction> DBcorrections = mdb.Corrections
                                                                    .Include(i => i.Мікрофонд.Фонд)
                                                                    .Include(i => i.КФК)
                                                                    .Include(i => i.Головний_розпорядник)
                                                                    .Include(i => i.КЕКВ)
                                                                    .Where(w =>
                                                                    w.Видалено == false &&
                                                                    w.Головний_розпорядник.Найменування == main_Manager.Найменування &&
                                                                    w.КЕКВ.Код == kEKB.Код &&
                                                                    w.КФК.Код == kFK.Код &&
                                                                    w.Мікрофонд.Фонд.Код == foundation.Код &&
                                                                    w.Проведено.Year == year)
                                                                    .ToList();

            if (localCorrections.Count != 0)
            {
                DBcorrections = DBcorrections.Where(w => localCorrections.Select(s => s.Id).Contains(w.Id) == false).ToList();
                DBcorrections.AddRange(localCorrections);
            }

            var EndCorrections = DBcorrections
                .Select(s => new
                {
                    s.Мікрофонд.Фонд,
                    s.КФК,
                    s.Головний_розпорядник,
                    s.КЕКВ,
                    s.Січень,
                    s.Лютий,
                    s.Березень,
                    s.Квітень,
                    s.Травень,
                    s.Червень,
                    s.Липень,
                    s.Серпень,
                    s.Вересень,
                    s.Жовтень,
                    s.Листопад,
                    s.Грудень
                })
                .ToList();
            #endregion

            var CorrectPlan = EndFillings.Union(EndCorrections).ToList();

            List<double> vs = new List<double>();

            //Вычисление месячных планов
            foreach (var item in names_months)
            {
                int numberOfMonth = names_months.IndexOf(item) + 1;

                double monthPlan = CorrectPlan.Sum(s => (double)s.GetType().GetProperty(item).GetValue(s));

                vs.Add(monthPlan);
            }

            return vs;
        }

        /// <summary>
        /// Метод возвращает уточненный план и остатки (накопительно) по переданным аргументам, учитывая данные, как из БД, так и из локального хранилища
        /// </summary>
        /// <param name="db">Контекст БД</param>
        /// <param name="year">Год</param>
        /// <param name="kFK">КПБ(КФК)</param>
        /// <param name="main_Manager">Главный распорядитель</param>
        /// <param name="kEKB">КЕКВ</param>
        /// <param name="foundation">Фонд</param>
        /// <returns>Уточненный план и накопительный остаток</returns>
        public static Dictionary<TypeOfFinanceData, List<double>> GetCurrentPlanAndRemainderFromDBPerMonth(DBSolom.Db db, int year, KFK kFK, Main_manager main_Manager, KEKB kEKB, Foundation foundation)
        {
            Dictionary<TypeOfFinanceData, List<double>> vs = new Dictionary<TypeOfFinanceData, List<double>>();
            vs.Add(TypeOfFinanceData.CurrentPlan, GetCurrentPlanFromDBPerMonth(db, year, kFK, main_Manager, kEKB, foundation));
            vs.Add(TypeOfFinanceData.Remainders, GetRemainderFromDBPerMonth(db, year, kFK, main_Manager, kEKB, foundation));
            return vs;
        }

        /// <summary>
        /// Метод возвращает детализированное описание списка ошибок касаемых проведения фин. документов
        /// </summary>
        /// <param name="db">Контекст БД</param>
        /// <param name="date">Дата</param>
        /// <param name="kFK">КПБ(КФК)</param>
        /// <param name="main_Manager">Главный распорядитель</param>
        /// <param name="kEKB">КЕКВ</param>
        /// <param name="foundation">Фонд</param>
        /// <returns>Список ошибок</returns>
        public static List<string> ChangeFinDocIsAllow(DBSolom.Db db, DateTime date, KFK kFK, Main_manager main_Manager, KEKB kEKB, Foundation foundation)
        {
            var x = GetCurrentPlanAndRemainderFromDBPerMonth(db, date.Year, kFK, main_Manager, kEKB, foundation);
            List<string> errors = new List<string>();

            for (int i = 0; i < 12; i++)
            {
                if (x[TypeOfFinanceData.Remainders][i] < 0)
                {
                    errors.Add($"[Дата: {date.ToShortDateString()}] [Фонд: {foundation.Код}] [КПБ: {kFK.Код}]" +
                        $" [Головний розпорядник: {main_Manager.Найменування}]" +
                        $" [КЕКВ: {kEKB.Код}] [Місяць: {names_months[i]}] [Остаток:{x[TypeOfFinanceData.Remainders][i]}]");
                }
                if (x[TypeOfFinanceData.CurrentPlan][i] < 0)
                {
                    errors.Add($"[Дата: {date.ToShortDateString()}] [Фонд: {foundation.Код}] [КПБ: {kFK.Код}]" +
                        $" [Головний розпорядник: {main_Manager.Найменування}]" +
                        $" [КЕКВ: {kEKB.Код}] [Місяць: {names_months[i]}] [План:{x[TypeOfFinanceData.CurrentPlan][i]}]");
                }
            }

            return errors;
        }

        /// <summary>
        /// Finalaser
        /// </summary>
        /// <param name="e">Standart argument of event window.closing</param>
        /// <param name="db">Datebase context</param>
        public static void dbHaveNotSavedChanges<T>(CancelEventArgs e, ObservableCollection<T> C, DBSolom.Db db)
        {
            foreach (var item in C)
            {
                if (db.Entry(item).State != EntityState.Unchanged)
                {
                    var resultMsg = MessageBox.Show("У вас є незафіксовані зміни. Зафіксувати?", "Maestro:[Довідки]", MessageBoxButton.YesNoCancel, MessageBoxImage.Warning);
                    switch (resultMsg)
                    {
                        case MessageBoxResult.Cancel:
                            e.Cancel = true;
                            break;
                        case MessageBoxResult.Yes:
                            try
                            {
                                db.SaveChanges();
                                MessageBox.Show("Зміни збережено!");
                                e.Cancel = false;
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }
                            break;
                        case MessageBoxResult.No:
                            e.Cancel = false;
                            break;
                    }
                    break;
                }
            }
        }
    }

    public class CorrectionCntxMenu
    {
        public List<Label> cntx_dict_lbl { get; set; } = new List<Label>();
        public List<ComboBox> cntx_dict_cmb { get; set; } = new List<ComboBox>();
        public List<TextBox> cntx_dict_txt { get; set; } = new List<TextBox>();
    }

    public class Filters
    {
        public Filters()
        {
            GetFilters = new List<Dictionary<string, dynamic>>();
        }

        public List<Dictionary<string, dynamic>> GetFilters { get; set; }
    }
    public class FillingDigitConverter : IValueConverter
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
    public class AdditionalEntities
    {
        public string Property { get; set; }
        public Dictionary<string, dynamic> Value = new Dictionary<string, dynamic>();
    }

    public enum TypeOfFinanceData
    {
        CurrentPlan,
        Remainders
    }

    public class ListMonths
    {
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
}
