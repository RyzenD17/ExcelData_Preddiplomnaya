using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для SelectMonthsReport.xaml
    /// </summary>
    public partial class SelectMonthsReport : Window
    {
        MainPage main;

        //спсок соответствия месяц = нужно вносить в отчет или нет
        List<(string m, bool check)> listM;
        public SelectMonthsReport()
        {
            InitializeComponent();

            //получаем ссылку на главную страницу программы
            main = (MainPage)((MainWindow)App.Current.Windows[0]).frame.Content;

            //получаем список месяцев которые доступны для отчета
            var months = main.CBChooseList.Items.OfType<string>().ToList();

            //задаем контекст для списка
            listBox.DataContext = months;

            //стандартное заполнение
            listM = new List<(string m, bool check)>();
            foreach (var month in months)
            {
                listM.Add((month, false));
            }

        }

        /// <summary>
        /// При нажатии на чек бокс заносим данные в список
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var c = ((CheckBox)sender).Content;
            var i = listM.FindIndex(a => a.m == c.ToString());
            listM[i] = (listM[i].m, true);
        }
        /// <summary>
        /// При снятии чек бокс заносим данные в список
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            var c = ((CheckBox)sender).Content;
            var i = listM.FindIndex(a => a.m == c.ToString());
            listM[i] = (listM[i].m, false);
        }
        /// <summary>
        /// Кнопка создания отчета
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMakeReport_Click(object sender, RoutedEventArgs e)
        {
            //запускаем функцию из главной страницы (криво но эффективно)
            main.YearReport(listM.Where(a=>a.check).Select(a=>a.m).ToList());
            MessageBox.Show("Отчет по месяцам был создан!");
            this.Close();
        }
    }
}
