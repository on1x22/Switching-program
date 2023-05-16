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
    /// Логика взаимодействия для Changing.xaml
    /// </summary>
    public partial class Changing : Window
    {
        /*public string nameE;
        bool stateE;*/
        public Changing()
        {
            InitializeComponent();
            
            /*if ((bool)checkBox1.IsChecked)
            {
                 stateE = true;
            }
            else stateE = false;
           nameE = textBox1.Text;*/
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            if (main.nameStart != textBox1.Text || main.stateStart != checkBox1.IsChecked)
            {
                main.nameStart = textBox1.Text;
                if ((bool)checkBox1.IsChecked)
                {
                    main.stateStart = true;
                }
                else main.stateStart = false;
                main.checker = true;
            }
            this.Close();
        }
    }
}
