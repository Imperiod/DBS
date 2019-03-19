using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DBSolom;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace Main
{
    public partial class MainWindow : Window
    {
        
        public MainWindow()
        {
            InitializeComponent();
        }

        private void MI_Macrofoundations_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.Macrofoundation macrofoundation = new Dictionary.Macrofoundation();
            macrofoundation.Show();
        }

        private void MI_Foundations_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.Foundation foundation = new Dictionary.Foundation();
            foundation.Show();
        }

        private void MI_KFB_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.KFB kFB = new Dictionary.KFB();
            kFB.Show();
        }

        private void MI_KDB_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.KDB kDB = new Dictionary.KDB();
            kDB.Show();
        }

        private void MI_KEKB_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.KEKB kEKB = new Dictionary.KEKB();
            kEKB.Show();
        }

        private void MI_KFK_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.KFK kFK = new Dictionary.KFK();
            kFK.Show();
        }

        private void MI_Main_managers_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.Main_manager main_Manager = new Dictionary.Main_manager();
            main_Manager.Show();
        }

        private void MI_Managers_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.Manager manager = new Dictionary.Manager();
            manager.Show();
        }

        private void MI_Filling_Click(object sender, RoutedEventArgs e)
        {
            Docs.Filling filling = new Docs.Filling();
            filling.Show();
        }

        private void MI_DocStatus_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.DocStatus docStatus = new Dictionary.DocStatus();
            docStatus.Show();
        }

        private void MI_Correct_Click(object sender, RoutedEventArgs e)
        {
            Docs.Correction correction = new Docs.Correction();
            correction.Show();
        }

        private void MI_Users_Click(object sender, RoutedEventArgs e)
        {
            Sys.User user = new Sys.User();
            user.Show();
        }

        private void MI_Lows_Click(object sender, RoutedEventArgs e)
        {
            Sys.Low low = new Sys.Low();
            low.Show();
        }

        private void MI_MicroFilling_Click(object sender, RoutedEventArgs e)
        {
            Docs.Microfilling microfilling = new Docs.Microfilling();
            microfilling.Show();
        }

        private void MI_Financing_Click(object sender, RoutedEventArgs e)
        {
            Docs.Financing financing = new Docs.Financing();
            financing.Show();
        }

        private void MI_Microfoundations_Click(object sender, RoutedEventArgs e)
        {
            Dictionary.MicroFoundation microFoundation = new Main.Dictionary.MicroFoundation();
            microFoundation.Show();
        }

        private void MI_CurrPlan_Click(object sender, RoutedEventArgs e)
        {
            Functional.CurrPlan currPlan = new Functional.CurrPlan();
            currPlan.Show();
        }

        private void MI_Remainders_Click(object sender, RoutedEventArgs e)
        {
            Maestro.Functional.Remainders currPlan = new Maestro.Functional.Remainders();
            currPlan.Show();
        }

        private void MI_Filling_FromExcel_Click(object sender, RoutedEventArgs e)
        {
            DBSolom.Db db = new Db(Func.GetConnectionString);

            if (db.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Filling == true) is null)
            {
                MessageBox.Show("У вас відсутні права на виконання цієї операції!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
            else
            {
                if (MessageBox.Show("Увага!\nЦя операція є небезпечною, перевірте чи всі головні розпорядники є в базі, кекв та інші необхідні властивості.\nВи підтверджуєте виконання?", "Maestro", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                    if (openFileDialog.ShowDialog() == true)
                    {
                        var Task = new Task(() =>
                        {
                            //Потому что Finally
                            Excel.Application application = null;
                            Excel.Workbook workbook = null;
                            Excel.Worksheet worksheet = null;
                            Excel.Range range = null;
                            User user = db.Users.First(f => f.Видалено == false && f.Логін == Func.Login);
                            try
                            {
                                #region "Variables"
                                application = new Excel.Application();
                                application.AskToUpdateLinks = false;
                                application.DisplayAlerts = false;
                                workbook = application.Workbooks.Open(openFileDialog.FileName);
                                worksheet = workbook.Worksheets["Maestro_Data"];
                                double[] months = new double[12];
                                DateTime Проведено_Е = DateTime.Now;
                                int fond_code = 0;
                                string main_manager = null;
                                long kfk_code = 0;
                                long kfb_code = 0;
                                long kdb_code = 0;
                                long kekv_code = 0;
                                Foundation foundation;
                                Main_manager Main_Manager;
                                KFK kFK;
                                KFB kFB;
                                KDB kDB;
                                KEKB kEKB;
                                List<Filling> fillings = new List<Filling>();
                                List<string> errors = new List<string>();
                                #endregion

                                for (int i = 2; i <= application.WorksheetFunction.CountA(worksheet.Columns[1]); i++)
                                {
                                    #region "Variables"
                                    range = worksheet.Cells[i, 1];
                                    Проведено_Е = Convert.ToDateTime(Convert.ToString(range.Value2));

                                    range = worksheet.Cells[i, 2];
                                    fond_code = Convert.ToInt32(Convert.ToString(range.Value2));
                                    foundation = db.Foundations.First(f => f.Видалено == false && f.Код == fond_code);

                                    range = worksheet.Cells[i, 3];
                                    main_manager = (string)range.Value2;
                                    Main_Manager = db.Main_Managers.First(f => f.Видалено == false && f.Найменування == main_manager);

                                    range = worksheet.Cells[i, 4];
                                    kfk_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kFK = db.KFKs.First(f => f.Видалено == false && f.Код == kfk_code);

                                    range = worksheet.Cells[i, 5];
                                    kfb_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kFB = db.KFBs.First(f => f.Видалено == false && f.Код == kfb_code);

                                    range = worksheet.Cells[i, 6];
                                    kdb_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kDB = db.KDBs.First(f => f.Видалено == false && f.Код == kdb_code);

                                    range = worksheet.Cells[i, 7];
                                    kekv_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kEKB = db.KEKBs.First(f => f.Видалено == false && f.Код == kekv_code);

                                    #endregion

                                    for (int k = 1; k <= 12; k++)
                                    {
                                        range = worksheet.Cells[i, 7 + k];
                                        months[k - 1] = Convert.ToDouble(Convert.ToString(range.Value2 ?? 0));
                                    }

                                    Filling filling = new Filling()
                                    {
                                        Створив = user,
                                        Змінив = user,
                                        Підписано = true,
                                        Проведено = Проведено_Е,

                                        Фонд = foundation,
                                        Головний_розпорядник = Main_Manager,
                                        КФК = kFK,
                                        КФБ = kFB,
                                        КДБ = kDB,
                                        КЕКВ = kEKB,

                                        Січень = months[0],
                                        Лютий = months[1],
                                        Березень = months[2],
                                        Квітень = months[3],
                                        Травень = months[4],
                                        Червень = months[5],
                                        Липень = months[6],
                                        Серпень = months[7],
                                        Вересень = months[8],
                                        Жовтень = months[9],
                                        Листопад = months[10],
                                        Грудень = months[11]
                                    };

                                    fillings.Add(filling);
                                }

                                foreach (var item in fillings)
                                {
                                    List<double> x = Func.GetRamainedFromDBPerMonth(db, item.Проведено.Year, item.КФК, item.Головний_розпорядник, item.КЕКВ, item.Фонд);
                                    for (int i = 0; i < 12; i++)
                                    {
                                        double t = (double)item.GetType().GetProperty(Func.names_months[i]).GetValue(item);
                                        if (x[i] + t < 0)
                                        {
                                            errors.Add($"[Дата: {item.Проведено.ToShortDateString()}] [Фонд: {item.Фонд.Код}] [КПБ: {item.КФК.Код}]" +
                                                $" [Головний розпорядник: {item.Головний_розпорядник.Найменування}]" +
                                                $" [КЕКВ: {item.КЕКВ.Код}] [Місяць: {Func.names_months[i]}] [Залишок:{x[i]}] [Корегування: {t}] [Різниця: {x[i] - t}]");
                                        }
                                    }
                                }

                                if (errors.Count == 0)
                                {
                                    db.Fillings.AddRange(fillings);
                                    db.SaveChanges();
                                    MessageBox.Show("Готово!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Information);
                                }
                                else
                                {
                                    Dispatcher.Invoke(() =>
                                    {
                                        Maestro.Sys.Errors er = new Maestro.Sys.Errors(errors);
                                        er.Show();
                                    });
                                }
                            }
                            catch (Exception Exp)
                            {
                                MessageBox.Show(Exp.Message);
                            }
                            finally
                            {
                                if (workbook != null)
                                {
                                    workbook.Close(false);
                                }
                                application = null;
                                openFileDialog = null;
                                workbook = null;
                                worksheet = null;
                            }
                        });

                        Task.Start();
                    }
                    else
                    {
                        MessageBox.Show("Оберіть файл з листом Maestro_Data!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Hand);
                    }
                }
            }
        }

        private void MI_MicroFilling_FromExcel_Click(object sender, RoutedEventArgs e)
        {

            DBSolom.Db db = new Db(Func.GetConnectionString);

            if (db.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Microfilling == true) is null)
            {
                MessageBox.Show("У вас відсутні права на виконання цієї операції!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
            else
            {
                if (MessageBox.Show("Увага!\nЦя операція є небезпечною, перевірте чи всі головні розпорядники є в базі, кекв та інші необхідні властивості.\nВи підтверджуєте виконання?", "Maestro", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                    if (openFileDialog.ShowDialog() == true)
                    {
                        var Task = new Task(() =>
                        {
                            Excel.Application application = null;
                            Excel.Workbook workbook = null;
                            Excel.Worksheet worksheet = null;
                            Excel.Range range = null;
                            User user = db.Users.First(f => f.Видалено == false && f.Логін == Func.Login);
                            try
                            {
                                application = new Excel.Application();
                                application.AskToUpdateLinks = false;
                                application.DisplayAlerts = false;
                                workbook = application.Workbooks.Open(openFileDialog.FileName);
                                worksheet = workbook.Worksheets["Maestro_Data"];
                                double[] months = new double[12];
                                DateTime Проведено_Е = DateTime.Now;
                                string main_manager = null;
                                string microfond = null;
                                long kfk_code = 0;
                                long kfb_code = 0;
                                long kdb_code = 0;
                                long kekv_code = 0;
                                Main_manager Main_Manager;
                                KFK kFK;
                                KFB kFB;
                                KDB kDB;
                                KEKB kEKB;
                                MicroFoundation microFoundation;
                                List<MicroFilling> fillings = new List<MicroFilling>();

                                for (int i = 2; i <= application.WorksheetFunction.CountA(worksheet.Columns[1]); i++)
                                {
                                    #region "Variables"
                                    range = worksheet.Cells[i, 1];
                                    Проведено_Е = Convert.ToDateTime(Convert.ToString(range.Value2));

                                    range = worksheet.Cells[i, 2];
                                    microfond = (string)range.Value2;
                                    microFoundation = db.MicroFoundations.First(f => f.Видалено == false && f.Повністю == microfond);

                                    range = worksheet.Cells[i, 3];
                                    main_manager = (string)range.Value2;
                                    Main_Manager = db.Main_Managers.First(f => f.Видалено == false && f.Найменування == main_manager);

                                    range = worksheet.Cells[i, 4];
                                    kfk_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kFK = db.KFKs.First(f => f.Видалено == false && f.Код == kfk_code);

                                    range = worksheet.Cells[i, 5];
                                    kfb_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kFB = db.KFBs.First(f => f.Видалено == false && f.Код == kfb_code);

                                    range = worksheet.Cells[i, 6];
                                    kdb_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kDB = db.KDBs.First(f => f.Видалено == false && f.Код == kdb_code);

                                    range = worksheet.Cells[i, 7];
                                    kekv_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kEKB = db.KEKBs.First(f => f.Видалено == false && f.Код == kekv_code);

                                    #endregion

                                    for (int k = 0; k < 12; k++)
                                    {
                                        range = worksheet.Cells[i, 8 + k];
                                        months[k] = Convert.ToDouble(Convert.ToString(range.Value2 ?? 0));
                                    }

                                    MicroFilling filling = new MicroFilling()
                                    {
                                        Створив = user,
                                        Змінив = user,
                                        Підписано = true,
                                        Проведено = Проведено_Е,
                                        Мікрофонд = microFoundation,
                                        Головний_розпорядник = Main_Manager,
                                        КФК = kFK,
                                        КФБ = kFB,
                                        КДБ = kDB,
                                        КЕКВ = kEKB,

                                        Січень = months[0],
                                        Лютий = months[1],
                                        Березень = months[2],
                                        Квітень = months[3],
                                        Травень = months[4],
                                        Червень = months[5],
                                        Липень = months[6],
                                        Серпень = months[7],
                                        Вересень = months[8],
                                        Жовтень = months[9],
                                        Листопад = months[10],
                                        Грудень = months[11]
                                    };

                                    fillings.Add(filling);
                                }

                                workbook.Close(false);
                                db.Microfillings.AddRange(fillings);
                                db.SaveChanges();
                                MessageBox.Show("Готово!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            catch (Exception Exp)
                            {
                                MessageBox.Show(Exp.Message);
                            }
                            finally
                            {
                                application = null;
                                openFileDialog = null;
                                workbook = null;
                                worksheet = null;
                            }
                        });

                        Task.Start();
                    }
                    else
                    {
                        MessageBox.Show("Оберіть файл з листом Maestro_Data!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Hand);
                    }
                }
            }
        }

        private void MI_Financing_FromExcel_Click(object sender, RoutedEventArgs e)
        {

            DBSolom.Db db = new Db(Func.GetConnectionString);

            if (db.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.Financing == true) is null)
            {
                MessageBox.Show("У вас відсутні права на виконання цієї операції!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Stop);
            }
            else
            {
                if (MessageBox.Show("Увага!\nЦя операція є небезпечною та дуже затратною в часі, перевірте чи всі головні розпорядники є в базі, кекв та інші необхідні властивості.\nВи підтверджуєте виконання?", "Maestro", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Excel files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls";
                    if (openFileDialog.ShowDialog() == true)
                    {
                        var Task = new Task(() =>
                        {
                            //Потому что Finally
                            Excel.Application application = null;
                            Excel.Workbook workbook = null;
                            Excel.Worksheet worksheet = null;
                            Excel.Range range = null;
                            User user = db.Users.First(f => f.Видалено == false && f.Логін == Func.Login);
                            try
                            {
                                #region "Variables"
                                application = new Excel.Application();
                                application.AskToUpdateLinks = false;
                                application.DisplayAlerts = false;
                                workbook = application.Workbooks.Open(openFileDialog.FileName);
                                worksheet = workbook.Worksheets["Maestro_Data"];
                                double sum = 0;
                                DateTime Проведено_Е = DateTime.Now;
                                string main_manager = null;
                                string microfond = null;
                                long kfk_code = 0;
                                long kekv_code = 0;
                                Main_manager Main_Manager;
                                KFK kFK;
                                KEKB kEKB;
                                MicroFoundation microFoundation;
                                List<Financing> localFinancings = new List<Financing>();
                                List<string> errors = new List<string>();
                                #endregion

                                for (int i = 2; i <= application.WorksheetFunction.CountA(worksheet.Columns[1]); i++)
                                {
                                    #region "Variables"
                                    range = worksheet.Cells[i, 1];
                                    Проведено_Е = Convert.ToDateTime(Convert.ToString(range.Value2));

                                    range = worksheet.Cells[i, 2];
                                    microfond = (string)range.Value2;
                                    microFoundation = db.MicroFoundations.Include(c => c.Фонд).First(f => f.Видалено == false && f.Повністю == microfond);

                                    range = worksheet.Cells[i, 3];
                                    main_manager = (string)range.Value2;
                                    Main_Manager = db.Main_Managers.First(f => f.Видалено == false && f.Найменування == main_manager);

                                    range = worksheet.Cells[i, 4];
                                    kfk_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kFK = db.KFKs.First(f => f.Видалено == false && f.Код == kfk_code);

                                    range = worksheet.Cells[i, 5];
                                    kekv_code = Convert.ToInt64(Convert.ToString(range.Value2));
                                    kEKB = db.KEKBs.First(f => f.Видалено == false && f.Код == kekv_code);

                                    range = worksheet.Cells[i, 6];
                                    sum = Convert.ToDouble(Convert.ToString(range.Value2 ?? 0));
                                    #endregion

                                    Financing financing = new Financing()
                                    {
                                        Створив = user,
                                        Змінив = user,
                                        Підписано = true,
                                        Проведено = Проведено_Е,
                                        Мікрофонд = microFoundation,
                                        Головний_розпорядник = Main_Manager,
                                        КФК = kFK,
                                        КЕКВ = kEKB,
                                        Сума = sum
                                    };

                                    localFinancings.Add(financing);
                                }
                                
                                var EndFinancings = localFinancings.GroupBy(g => new
                                {
                                    g.Проведено.Year,
                                    g.Головний_розпорядник,
                                    g.КЕКВ,
                                    g.КФК,
                                    g.Мікрофонд.Фонд
                                }).ToList();

                                foreach (var item in EndFinancings)
                                {
                                    List<double> x = Func.GetRamainedFromDBPerMonth(db, item.Key.Year, item.Key.КФК, item.Key.Головний_розпорядник, item.Key.КЕКВ, item.Key.Фонд);
                                    double f = 0;

                                    for (int i = 0; i < 12; i++)
                                    {
                                        f = item.Where(w => w.Проведено.Month == (i + 1)).Sum(s => s.Сума);

                                        if (x[i] - f < 0)
                                        {
                                            errors.Add($"[Дата: {item.Min(m=>m.Проведено).ToShortDateString()}-{item.Max(m=>m.Проведено).ToShortDateString()}] [Фонд: {item.Key.Фонд.Код}] [КПБ: {item.Key.КФК.Код}]" +
                                                $" [Головний розпорядник: {item.Key.Головний_розпорядник.Найменування}]" +
                                                $" [КЕКВ: {item.Key.КЕКВ.Код}] [Місяць: {Func.names_months[i]}] [Залишок:{x[i]}] [Фінансування: {f}] [Різниця: {x[i] - f}]");
                                        }
                                    }
                                }

                                if (errors.Count == 0)
                                {
                                    db.Financings.AddRange(localFinancings);
                                    db.SaveChanges();
                                    MessageBox.Show("Готово!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Information);
                                }
                                else
                                {
                                    Dispatcher.Invoke(() => 
                                    {
                                        Maestro.Sys.Errors er = new Maestro.Sys.Errors(errors);
                                        er.Show();
                                    });
                                }
                            }
                            catch (Exception Exp)
                            {
                                MessageBox.Show(Exp.Message);
                            }
                            finally
                            {
                                if (workbook != null)
                                {
                                    workbook.Close(false);
                                }
                                application = null;
                                openFileDialog = null;
                                workbook = null;
                                worksheet = null;
                            }
                        });

                        Task.Start();
                    }
                    else
                    {
                        MessageBox.Show("Оберіть файл з листом Maestro_Data!", "Maestro", MessageBoxButton.OK, MessageBoxImage.Hand);
                    }
                }
            }
        }
    }
}
