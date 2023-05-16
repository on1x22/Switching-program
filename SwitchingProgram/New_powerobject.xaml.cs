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

namespace WpfApplication1
{
    /// <summary>
    /// Логика взаимодействия для New_powerobject.xaml
    /// </summary>
    public partial class New_powerobject : Window
    {
        bool isUsedPO;
        public New_powerobject()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            if (textBox1.Text != "")
            {
                if ((bool)checkBox1.IsChecked)
                {
                    isUsedPO = true;
                }
                else isUsedPO = false;
                PowerObject PO = new PowerObject();
                PO.NamePO = textBox1.Text;
                PO.isUsed = isUsedPO;
                PO.organisationPO = textBox2.Text;
                main.listOfPowerObjects.Add(PO);
                this.Close();
            }
            else MessageBox.Show("Не задано название энергообъекта");
        }
    }
}
