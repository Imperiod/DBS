using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.VisualBasic;

namespace Main
{
    public partial class Login : Window
    {
        DBSolom.Db db = new DBSolom.Db(Func.GetConnectionString);

        public Login()
        {
            InitializeComponent();
            txtLogin.Focus();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string t = txtLogin.Text;
            string p = pswMain.Password;
            Task.Factory.StartNew( () => GetAction(t, p) );
        }

        private async void GetAction (string l, string p = "")
        {
            await Dispatcher.BeginInvoke(new ThreadStart(() => { BTN.Visibility = Visibility.Collapsed; PB.Visibility = Visibility.Visible; }));
            if (l != "")
            {
                var q = db.Users.FirstOrDefault(w => w.Видалено == false && w.Логін == l);
                if (q is null)
                {
                    MessageBox.Show("Такого логіна не існує.", "Maestro", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else if (q.New)
                {
                    string pass;

                    do
                    {
                        pass = Interaction.InputBox("Введіть ваш новий пароль: \n Не менш ніж 4 символи.", "Maestro");
                    } while (pass != "" && pass.Length < 4);

                    if (pass == "")
                    {
                        await Dispatcher.BeginInvoke(new ThreadStart(() => { PB.Visibility = Visibility.Collapsed; BTN.Visibility = Visibility.Visible; }));
                        return;
                    }

                    q.Пароль = pass;
                    q.New = false;
                    await Task.Factory.StartNew(() => db.SaveChanges());
                    MessageBox.Show("Пароль встановлено.\n\nПри повторному вході використовуйте ваш новий пароль.", "Maestro", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    Func.Login = q.Логін;
                    await Dispatcher.BeginInvoke( new ThreadStart( () => { MainWindow mainWindow = new MainWindow(); mainWindow.Show(); Close(); } ) );
                }
                else
                {
                    if (pswMain.Password != "")
                    {
                        if (q.Пароль == pswMain.Password)
                        {
                            Func.Login = q.Логін;
                            await Dispatcher.BeginInvoke(new ThreadStart(() => { MainWindow mainWindow = new MainWindow(); mainWindow.Show(); Close(); }));
                        }
                        else
                        {
                            MessageBox.Show("Хибний пароль.", "Maestro", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Заповніть Пароль.", "Maestro", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
            }
            else
            {
                MessageBox.Show("Заповніть Логін.", "Maestro", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            await Dispatcher.BeginInvoke(new ThreadStart(() => { PB.Visibility = Visibility.Collapsed; BTN.Visibility = Visibility.Visible; }));
        }
    }
}
