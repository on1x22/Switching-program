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
    /// Логика взаимодействия для Options.xaml
    /// </summary>
    public partial class Options : Window
    {
        public Options()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            /*if (main.can_change_list_of_objects == true)
            {
                //int num = Convert.ToInt32(comboBox1.Text);
                main.num_obj = Convert.ToInt32(comboBox1.Text);
            }
            else
            {
                MessageBox.Show("Запрещено изменение количества подстанций");
            }*/

            if (textBox1.Text != main.progOption.nameSP)
            {
                main.progOption.nameSP = textBox1.Text;
            }

            if (textBox2.Text != main.progOption.roditPadezh)
            {
                main.progOption.roditPadezh = textBox2.Text;
            }

            this.Close();
        }
    }    
}
