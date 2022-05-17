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
    /// Логика взаимодействия для AboutWindow.xaml
    /// </summary>
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
            string starttxt = "\nДля начала работы, необходимо открыть Excel файл с списком группы.\nДля этого необходимо нажать кнопку 'Выбрать файл', после необходимо выбрать нужный файл.\nДля перехода по листам нужно использовать выпадающий список.\nДля перехода на страницу просмотра листа пропусков необходимо нажать кнопку 'Предпросмотр'. \nДля распределения подгрупп необходимо нажать кнопку 'Распределить подгруппы'.";
            StartText.Text = starttxt;
            string skipstxt = "\nДля подсчёта пропусков студентов необходимо на странице предпросомтр нажать кнопку 'Пересчитать пропуски'.";
            SkipText.Text = skipstxt;
            string reporttxt = "\nДля создания отчёта в Word необходимо на странице предпросомтр нажать кнопку 'Сохранить в Word'\n(при создании отчёта Word есть возможность отправить его на почту)\nДля создания отчёта в Word необходимо на странице предпросомтр нажать кнопку 'Сохранить в Excel'\nДля создания отчёта по месяцам на странице предпросомтр нажать кнопку 'Отчёт по месяцам'\n(чтобы очистить информацию о пропусках за год необходимо нажать кнопку 'Очистить годовые данные').";
            ReportText.Text = reporttxt;
        }
    }
}
