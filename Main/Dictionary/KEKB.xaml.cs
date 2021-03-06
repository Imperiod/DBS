﻿using DBSolom;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data.Entity;
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

namespace Main.Dictionary
{
    public partial class KEKB : Window
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

        DBSolom.Db db = new Db(Func.GetConnectionString);
        #endregion

        public KEKB()
        {
            InitializeComponent();

            #region "Load entity"
            
            db.KEKBs
                    .Include(i => i.Змінив)
                    .Include(i => i.Створив)
                    .OrderBy(o => o.Код)
                    .Load();

            #endregion

            ((CollectionViewSource)FindResource("cvs")).Source = db.KEKBs.Local;

            ((CollectionViewSource)FindResource("cvs")).Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Accept.Click += BTN_Accept_Click;
            BTN_Reset.Click += BTN_Reset_Click;
            BTN_ResetGroup.Click += BTN_ResetGroup_Click;
            BTN_Save.Click += BTN_Save_Click;
            BTN_ExportToExcel.Click += Func.BTN_ExportToExcel_Click;

            GetViewOfToolBox();
        }

        private void GetViewOfToolBox()
        {
            int t = 0;
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
            if (db.Lows.Include(i => i.Правовласник).FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.KEKB == true) is null)
            {
                DGM.IsReadOnly = true;
            }
        }

        private void DGM_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (((DBSolom.KEKB)e.Row.Item).Id == 0)
            {
                ((DBSolom.KEKB)e.Row.Item).Створив = db.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            }
            ((DBSolom.KEKB)e.Row.Item).Змінив = db.Users.FirstOrDefault(f => f.Видалено == false && f.Логін == Func.Login);
            ((DBSolom.KEKB)e.Row.Item).Змінено = DateTime.Now;
        }

        private void DGM_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            Func.GenerateColumnForDataGrid(db, ref counterForDGMColumns, e);
        }
    }
}
