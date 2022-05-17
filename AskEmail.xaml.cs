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

using ExcelData.Properties;

namespace ExcelData
{
    /// <summary>
    /// Логика взаимодействия для AskEmail.xaml
    /// </summary>
    public partial class AskEmail : Window
    {
        public AskEmail()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (txtEmail.Text == "")
            {
                return;
            }
            Settings.Default.EmailUser = txtEmail.Text;
            Settings.Default.Save();
            Close();
        }
    }
}
