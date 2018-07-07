using System;
using System.Collections.Generic;
using System.Data.Entity;
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
    }
}
