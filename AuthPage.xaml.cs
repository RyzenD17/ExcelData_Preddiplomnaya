using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
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

using ExcelData.Properties;

using FluentEmail.Core;
using FluentEmail.MailKitSmtp;
using FluentEmail.Smtp;

using Microsoft.Extensions.DependencyInjection;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для AuthPage.xaml
    /// </summary>
    public partial class AuthPage : Page
    {
        bool isFirstTime;
        string forgotCode = "";
        public AuthPage()
        {
            InitializeComponent();
            //проверка на 1 запуск
            isFirstTime = Settings.Default.Password == "";
            if (isFirstTime)
            {
                groupTitle.Header = "Первый вход в систему!\nЗадайте новый пароль.";
               
            }
            //вырубаем все что связано со сбросом пароля
            groupCheckCode.Visibility = Visibility.Hidden;
            groupNewPass.Visibility = Visibility.Hidden;
            groupSendCode.Visibility = Visibility.Hidden;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //вводим код если 1 раз заходим иначе проверяем ввод
            if (isFirstTime)
            {
                Settings.Default.Password = txtPass.Password;
                Settings.Default.Save();
               
            }
            else
            {
                if (Settings.Default.Password != txtPass.Password)
                {
                    MessageBox.Show("Неправильный пароль!");
                    return;
                }
            }
            NavigationService.Navigate(new MainPage());
        }
        /// <summary>
        /// Отправка почты по smtp серверу
        /// </summary>
        /// <param name="txtTo">Адрем получателя</param>
        /// <param name="body">Сообщение</param>
        /// <returns></returns>
        private static async Task SendEmailAsync(string txtTo, string body = "")
        {
            var toS = txtTo;
            //smtp почта / логин
            var fromS = Properties.Settings.Default.EmailSmtp;
            //пароль от smtp 
            var pass = Properties.Settings.Default.PasswordSmtp;

            //формируем письмо
            MailAddress from = new MailAddress(fromS, "Классный руководитель");
            MailAddress to = new MailAddress(toS);
            MailMessage m = new MailMessage(from, to);

            
            //тема
            m.Subject = "Восстановление доступа от приложения Классный руководитель";
            //тело
            m.Body = body;

            //создаем подключение smtp и отправляем письмо
            SmtpClient smtp = new SmtpClient("smtp.mail.ru", 587);
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = new NetworkCredential(fromS, pass);
            smtp.EnableSsl = true;
            //ожидаем отправки
            await smtp.SendMailAsync(m);

            Console.WriteLine("Письмо отправлено");
        }
        private void btnForgot_Click(object sender, RoutedEventArgs e)
        {
            groupSendCode.Visibility = Visibility.Visible;
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            //формируем код востановления
            forgotCode = Guid.NewGuid().ToString();
            var body = $"Ваш проверочный код для восстановления:\n\n{forgotCode}";
            //и отправляем с ним письмо
            await SendEmailAsync(txtEmail.Text, body);
            groupCheckCode.Visibility = Visibility.Visible;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            //проверяем введенный код если совпадают идем дальше
            var enteredCode = txtCode.Text;
            if (enteredCode != forgotCode)
            {
                MessageBox.Show("Код введен неверно!", "Код");
                return;
            }
            groupNewPass.Visibility = Visibility.Visible;
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            //задаем новый код и меняем его в настройках
            var newPass = txtNewPass.Password;
            Properties.Settings.Default.Password = newPass;
            Properties.Settings.Default.Save();
            MessageBox.Show("Пароль изменен!", "Успех");

            NavigationService.Navigate(new MainPage());
        }
    }
}
