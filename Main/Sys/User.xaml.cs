using DBSolom;
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

namespace Main.Sys
{
    public partial class User : Window
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

        public User()
        {
            InitializeComponent();

            #region "Load entity"
            
            db.Users.Load();

            #endregion

            ((CollectionViewSource)FindResource("cvs")).Source = db.Users.Local;

            ((CollectionViewSource)FindResource("cvs")).Filter += Func.CollectionView_Filter;

            DGM.GroupStyle.Add(((GroupStyle)FindResource("one")));

            BTN_Save.Click += BTN_Save_Click;
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

        private void DGM_Loaded(object sender, RoutedEventArgs e)
        {
            if (db.Lows.FirstOrDefault(f => f.Видалено == false && f.Правовласник.Логін == Func.Login && f.User == true) is null)
            {
                DGM.IsReadOnly = true;
            }
        }

        private void DGM_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            Func.GenerateColumnForDataGrid(db, ref counterForDGMColumns, e);
        }

        private void DGM_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Column.Header.ToString() == "Видалено" && ((CheckBox)e.EditingElement).IsChecked == true)
            {
                ((DBSolom.User)e.Row.DataContext).New = true;
                ((DBSolom.User)e.Row.DataContext).Пароль = "";
                ((DBSolom.User)e.Row.DataContext).Видалено = false;
            }
        }
    }
}
