using Microsoft.CSharp;
using Microsoft.Win32;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
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
using System.Windows.Input;
using System.Windows.Media;

namespace Main
{
    public static class Func
    {
        public static string Login { get; set; }

        private static Dictionary<Window, Expander> expanders = new Dictionary<Window, Expander>();
        public static void AddExp(Window window, Expander exp)
        {
            expanders.Add(window, exp);
            System.Windows.Threading.DispatcherTimer dispatcherTimer = new System.Windows.Threading.DispatcherTimer() { Interval = TimeSpan.FromMilliseconds(150) };
            dispatcherTimer.Tick += new EventHandler((object c, EventArgs eventArgs) =>
            {
                var b = expanders.FirstOrDefault(f => f.Key == window && f.Value == exp);
                if (b.Value != null)
                {
                    window.Dispatcher.Invoke(() => b.Value.IsExpanded = true);
                }
                ((System.Windows.Threading.DispatcherTimer)c).Stop();
            });
            dispatcherTimer.Start();
        }

        static DBSolom.Db db;
        public static DBSolom.Db GetDB
        {
            get
            {
                if (db is null)
                {
#if DEBUG
                    var f = System.IO.File.OpenText(Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.LastIndexOf("Main")) + "Main\\Connection.imperiod");
#else
                    var f = System.IO.File.OpenText(Environment.CurrentDirectory + "\\Connection.imperiod");
#endif
                    db = new DBSolom.Db(f.ReadLine());
                    f.Close();
                }
                return db;
            }
        }

        static public void GenerateColumnForDataGrid(ref int counterForDGMColumns, DataGridAutoGeneratingColumnEventArgs e)
        {
            CultureInfo cultureInfo = new CultureInfo("ru-RU", true);

            switch (e.Column.Header.ToString())
            {
                case "Id":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Id") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Видалено":
                    e.Column = new DataGridCheckBoxColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Видалено") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Створив":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,

                        ItemsSource = GetDB.Users
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Логін).ToList(),

                        DisplayMemberPath = "Логін",
                        SelectedValueBinding = new Binding("Створив") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Створив.Логін"
                    };
                    break;
                case "Створино":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Створино") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged, StringFormat = "dd.MM.yyyy HH:mm" },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Змінив":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,

                        ItemsSource = GetDB.Users
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Логін).ToList(),

                        DisplayMemberPath = "Логін",
                        SelectedValueBinding = new Binding("Змінив") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Змінив.Логін"
                    };
                    break;
                case "Правовласник":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,

                        ItemsSource = GetDB.Users
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Логін).ToList(),

                        DisplayMemberPath = "Логін",
                        SelectedValueBinding = new Binding("Правовласник") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsReadOnly = false,
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Правовласник.Логін"
                    };
                    break;
                case "Контакти":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Контакти") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Логін":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Логін") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Змінено":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Змінено") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged, StringFormat = "dd.MM.yyyy HH:mm" },
                        IsReadOnly = true,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Проведено":
                    #region "DatePicker"

                    Binding dateBind = new Binding("Проведено") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged, StringFormat = "dd.MM.yyyy" };

                    FrameworkElementFactory datePickerFactoryElem = new FrameworkElementFactory(typeof(DatePicker));
                    datePickerFactoryElem.SetValue(DatePicker.SelectedDateProperty, dateBind);
                    datePickerFactoryElem.SetValue(DatePicker.DisplayDateProperty, dateBind);

                    FrameworkElementFactory frameworkElementFactory = new FrameworkElementFactory(typeof(TextBlock));
                    frameworkElementFactory.SetValue(TextBlock.TextProperty, dateBind);

                    DataTemplate cellTemplate = new DataTemplate() { VisualTree = datePickerFactoryElem };
                    DataTemplate dataTemplate = new DataTemplate() { VisualTree = frameworkElementFactory };

                    DataGridTemplateColumn templateColumn = new DataGridTemplateColumn()
                    {
                        Header = e.Column.Header,
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
                        Header = e.Column.Header,
                        Binding = new Binding("Підписано") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        IsThreeState = false,
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Внутрішній_номер":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Внутрішній_номер") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Підстава":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Підстава") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Статус":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        ItemsSource = GetDB.DocStatuses
                        .Include(i => i.Змінив)
                        .Include(i => i.Створив)
                        .Where(w => w.Видалено == false)
                        .OrderBy(o => o.Повністю).ToList(),

                        DisplayMemberPath = "Повністю",
                        SelectedValueBinding = new Binding("Статус") { UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns,
                        SortMemberPath = "Статус.Повністю"
                    };
                    break;
                case "Повністю":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Повністю") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Найменування":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Найменування") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Код":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Код") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "КПОЛ":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("КПОЛ") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Код_ГУДКСУ":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Код_ГУДКСУ") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Код_УДКСУ":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Код_УДКСУ") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "ЕГРПОУ":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("ЕГРПОУ") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Рівень_розпорядника":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Binding = new Binding("Рівень_розпорядника") { Mode = BindingMode.TwoWay, UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged },
                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Розпорядник":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,

                        ItemsSource = GetDB.Managers
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
                        Width = 140,
                        ItemsSource = GetDB.Main_Managers
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
                        Width = 80,
                        ItemsSource = GetDB.KFKs
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

                        ItemsSource = GetDB.MacroFoundations
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
                        Width = 60,
                        ItemsSource = GetDB.Foundations
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
                        Width = 220,
                        ItemsSource = GetDB.MicroFoundations
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
                case "КДБ":
                    e.Column = new DataGridComboBoxColumn()
                    {
                        Header = e.Column.Header,
                        Width = 90,
                        ItemsSource = GetDB.KDBs
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
                        Width = 60,
                        ItemsSource = GetDB.KEKBs
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
                        Width = 50,
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
                        ItemsSource = GetDB.names_months.Concat(new List<string> { "Рік", "Період" }),
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
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Рік")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Період":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Період")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Січень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Січень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Лютий":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Лютий")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Березень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Березень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Квітень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Квітень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Травень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Травень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Червень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Червень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Липень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Липень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Серпень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Серпень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Вересень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Вересень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Жовтень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Жовтень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Листопад":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Листопад")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Грудень":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Грудень")
                        {
                            Mode = BindingMode.TwoWay,
                            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
                            Converter = new FillingDigitConverter()
                        },

                        DisplayIndex = counterForDGMColumns
                    };
                    break;
                case "Сума":
                    e.Column = new DataGridTextColumn()
                    {
                        Header = e.Column.Header,
                        Width = 150,
                        Binding = new Binding("Сума")
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
                        Header = e.Column.Header,
                        IsReadOnly = true,
                        Binding = new Binding("Уточнений_план")
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
                ItemsSource = GetDB.list,
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

        public static void GetVisibilityOfColumns(int t, ItemPropertyInfo item, ref Grid EXPHDN)
        {
            var w = ((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)((FrameworkElement)EXPHDN.Parent).Parent).Parent).Parent).Parent).Parent;
            Style st = (Style)((FrameworkElement)w).Resources["Style"];

            ToggleButton toggleButton = new ToggleButton()
            {
                Content = item.Name,
                IsThreeState = false,
                IsChecked = true,
                Style = st,
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                HorizontalAlignment = HorizontalAlignment.Stretch
            };

            Grid.SetColumn(toggleButton, t);

            toggleButton.Checked += HiddenUnhiddenColumn;
            toggleButton.Unchecked += HiddenUnhiddenColumn;
            EXPHDN.Children.Add(toggleButton);
        }

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

        public static void CollectionView_Filter(object sender, FilterEventArgs e)
        {
            Window active_window = (Window)((TypeInfo)sender.GetType()).DeclaredProperties.FirstOrDefault(f => f.Name == "InheritanceContext").GetValue(sender);
            List<Filters> filters = (List<Filters>)((TypeInfo)active_window.GetType()).DeclaredFields.First(f => f.Name == "GetFilters").GetValue(active_window);
            bool result;


            if (filters.Count == 0)
            {
                result = true;
            }
            else if ((e.Item.GetType().GetProperties().FirstOrDefault(f => f.Name == "Id") is null) || ((e.Item.GetType().GetProperties().FirstOrDefault(f => f.Name == "Id") != null) && (long)e.Item.GetType().GetProperty("Id").GetValue(e.Item) != 0))
            {
                List<bool> fdv = new List<bool>();
                try
                {
                    foreach (Filters item in filters) //Перебор всех фильтров по типу ИЛИ
                    {
                        fdv.Add(GetFill(item));
                    }
                }
                catch
                {
                    result = false;
                }
                result = fdv.Where(w => w == true).Count() > 0 ? true : false;
            }
            else
            {
                result = true;
            }

            e.Accepted = result;

            bool GetFill(Filters filter)
            {
                object q = null;
                dynamic ValueQuery = null;
                string name = "";
                string stringBuilder = "";
                string endstring = "return ";
                foreach (var micro_item in filter.GetFilters) //Перебор всех фильтров по типу И (если хоть один не проходит тогда False
                {
                    string typeValue = "";

                    if (e.Item.GetType().GetProperty(micro_item["prop"]).PropertyType.FullName.Contains("DBSolom"))
                    {
                        var x = ((PropertyInfo[])e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item).GetType().GetProperties()).Select(k => k.Name).ToList();
                        q = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item);

                        if (x.Contains("Код"))
                        {
                            name = "Код";
                        }
                        else if (x.Contains("Найменування"))
                        {
                            name = "Найменування";
                        }
                        else if (x.Contains("Повністю"))
                        {
                            name = "Повністю";
                        }
                        else if (x.Contains("Логін"))
                        {
                            name = "Логін";
                        }

                        if (q.GetType().GetProperty(name).GetValue(q) is null)
                        {
                            return false;
                        }

                        ValueQuery = q.GetType().GetProperty(name).GetValue(q);
                        typeValue = ValueQuery.GetType().Name;
                    }
                    else
                    {
                        typeValue = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item).GetType().Name;
                        ValueQuery = e.Item.GetType().GetProperty(micro_item["prop"]).GetValue(e.Item);
                    }

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
                            stringBuilder = $"{RemoveBadSymbols(ValueQuery.ToString()).Length} >= {RemoveBadSymbols(start).Length} && {RemoveBadSymbols(ValueQuery.ToString()).Length} <= {RemoveBadSymbols(end).Length}";
                        }
                        else
                        {
                            stringBuilder = $"{typeValue}.Parse(\"{RemoveBadSymbols(ValueQuery.ToString())}\") >= {typeValue}.Parse(\"{RemoveBadSymbols(start)}\") && {typeValue}.Parse(\"{RemoveBadSymbols(ValueQuery.ToString())}\") <= {typeValue}.Parse(\"{RemoveBadSymbols(end)}\")";
                        }
                    }
                    else if (micro_item["type"] == "[,]")
                    {
                        List<dynamic> list = new List<dynamic>();
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

                        stringBuilder = $"new System.Collections.Generic.List<{typeValue}>()" + "{";

                        foreach (var item in list)
                        {
                            if (list.Last() == item)
                            {
                                if (typeValue == "String")
                                {
                                    stringBuilder += $"\"{item.ToString().ToLower()}\"" + "}" + $".Contains(\"{RemoveBadSymbols(ValueQuery.ToString().ToLower())}\")";
                                }
                                else
                                {
                                    stringBuilder += $"{typeValue}.Parse(\"{item.ToString()}\")" + "}" + $".Contains({typeValue}.Parse(\"{ValueQuery.ToString()}\"))";
                                }
                            }
                            else
                            {
                                if (typeValue == "String")
                                {
                                    stringBuilder += $"\"{item.ToString().ToLower()}\", ";
                                }
                                else
                                {
                                    stringBuilder += $"{typeValue}.Parse(\"{item.ToString()}\"), ";
                                }
                            }
                        }
                    }
                    else if (micro_item["type"] == ">|<")
                    {
                        stringBuilder = $"\"{RemoveBadSymbols(ValueQuery.ToString()).ToLower()}\".Contains(\"{RemoveBadSymbols(micro_item["value"].ToString()).ToLower()}\")";
                    }
                    else
                    {
                        if (typeValue == "String")
                        {
                            stringBuilder = $"\"{RemoveBadSymbols(ValueQuery.ToString())}\" {micro_item["type"].ToString()} \"{RemoveBadSymbols(micro_item["value"].ToString())}\"";
                        }
                        else
                        {
                            stringBuilder = $"{typeValue}.Parse(\"{RemoveBadSymbols(ValueQuery.ToString())}\") {micro_item["type"].ToString()} {typeValue}.Parse(\"{RemoveBadSymbols(micro_item["value"].ToString())}\")";
                        }
                    }
                    endstring += "(" + stringBuilder + ") && ";
                }
                endstring = endstring.Substring(0, endstring.Length - 4) + ";";

                return Tech.CodeGeneration.CodeGenerator.ExecuteCode<bool>(endstring);
            }
        }

        private static string RemoveBadSymbols(string s)
        {
            return s.Replace("\"", "");
        }

        public static void Expander_MouseEnter(object sender, MouseEventArgs e)
        {
            AddExp((Window)((Grid)((Expander)sender).Parent).Parent, (Expander)sender);
        }

        public static void Expander_MouseLeave(object sender, MouseEventArgs e)
        {
            var x = expanders.FirstOrDefault(f => f.Key == (Window)((Grid)((Expander)sender).Parent).Parent && f.Value == (Expander)sender);
            if (x.Value != null)
            {
                expanders.Remove(x.Key);
            }
            ((Expander)sender).IsExpanded = false;
        }

        public static void BTN_ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            Window active_window = (Window)((Grid)((Expander)((Grid)((Button)sender).Parent).Parent).Parent).Parent;
            DataGrid DGM = (DataGrid)active_window.GetType().GetRuntimeFields().First(f => f.Name == "DGM").GetValue(active_window);
            ProgressBar PB = (ProgressBar)active_window.GetType().GetRuntimeFields().First(f => f.Name == "PB").GetValue(active_window);

            bool WorksheetExist = false;
            bool TableExist = false;
            if (DGM.SelectedCells.Count > 0)
            {
                List<Dictionary<string, dynamic>> Entities = new List<Dictionary<string, dynamic>>();
                int countColumns = 0;
                int countRows = 0;

                //Перенос сущностей
                foreach (var item in DGM.SelectedCells)
                {
                    if (item.Item.ToString() != "{NewItemPlaceholder}" && Entities.FirstOrDefault(dict => dict.FirstOrDefault(d => d.Key == "Id").Value.ToString() == item.Item.GetType().GetProperty("Id").GetValue(item.Item).ToString()) is null)
                    {
                        foreach (var itemm in item.Item.GetType().GetProperties().Select(s => s.Name).ToList())
                        {
                            var q = Entities.FirstOrDefault(dict => dict.FirstOrDefault(d => d.Key == "Id").Value.ToString() == item.Item.GetType().GetProperty("Id").GetValue(item.Item).ToString());
                            if (q is null)
                            {
                                Entities.Add(new Dictionary<string, dynamic>() { { itemm, item.Item.GetType().GetProperty(itemm).GetValue(item.Item) } });
                            }
                            else
                            {
                                dynamic q1, q2, q3;
                                switch (itemm)
                                {
                                    case "Створив":
                                    case "Змінив":
                                    case "Правовласник":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Логін").GetValue(q1));
                                        break;
                                    case "Статус":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Повністю").GetValue(q1));
                                        break;
                                    case "Головний_розпорядник":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Найменування").GetValue(q1));
                                        break;
                                    case "КФК":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Код").GetValue(q1));
                                        break;
                                    case "Фонд":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Код").GetValue(q1));
                                        break;
                                    case "Мікрофонд":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Повністю").GetValue(q1));
                                        q2 = q1.GetType().GetProperty("Фонд").GetValue(q1);
                                        q3 = q2.GetType().GetProperty("Код").GetValue(q2);
                                        q.Add("Фонд", q3);
                                        break;
                                    case "КДБ":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Код").GetValue(q1));
                                        break;
                                    case "КЕКВ":
                                        q1 = item.Item.GetType().GetProperty(itemm).GetValue(item.Item);
                                        q.Add(itemm, q1.GetType().GetProperty("Код").GetValue(q1));
                                        break;
                                    default:
                                        q.Add(itemm, item.Item.GetType().GetProperty(itemm).GetValue(item.Item));
                                        break;
                                }
                            }
                        }
                        countRows++;
                    }
                }

                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                if (openFileDialog.ShowDialog() == true)
                {
                    PB.IsIndeterminate = true;

                    var Task = new Task(() =>
                    {
                        #region "Variables"
                        Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application
                        {
                            AskToUpdateLinks = false,
                            DisplayAlerts = false,
                            Visible = false
                        };
                        Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Open(openFileDialog.FileName);
                        application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual;
                        int currentRow = 2;
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
                        #endregion

                        //Check exist worksheet
                        foreach (Microsoft.Office.Interop.Excel.Worksheet item in workbook.Worksheets)
                        {
                            if (item.Name == "Maestro_Data")
                            {
                                WorksheetExist = true;
                                worksheet = item;
                                break;
                            }
                        }

                        if (WorksheetExist)
                        {
                            if (worksheet.ListObjects.Count != 0)
                            {
                                for (int i = 1; i <= worksheet.ListObjects.Count; i++)
                                {
                                    if (worksheet.ListObjects[i].Name == "Maestro_DataTable")
                                    {
                                        TableExist = true;
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

                        //Headers
                        foreach (var column in Entities.First().Keys)
                        {
                            countColumns++;
                            worksheet.Cells[1, countColumns] = column;
                        }

                        if (TableExist == false)
                        {
                            worksheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[countRows, countColumns]], Type.Missing, Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "Maestro_DataTable";
                        }
                        else
                        {
                            worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[1000, 50]].Clear();
                        }

                        //Filling
                        foreach (var item in Entities)
                        {
                            int currentColumn = 1;
                            foreach (string ent in item.Keys)
                            {
                                worksheet.Cells[currentRow, currentColumn] = item.FirstOrDefault(dict => dict.Key == ent).Value;
                                currentColumn++;
                            }
                            currentRow++;
                        }

                        application.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationAutomatic;
                        MessageBox.Show("Виконано!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Information);
                        application.Visible = true;
                        openFileDialog = null;
                        application = null;
                        workbook = null;
                        worksheet = null;
                    PB.Dispatcher.Invoke(() => PB.IsIndeterminate = false);
                    });

                    Task.Start();
                }
            }
            else
            {
                MessageBox.Show("Виділіть всі строки, які будуть експортовані!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Hand);
            }
        }
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
}
