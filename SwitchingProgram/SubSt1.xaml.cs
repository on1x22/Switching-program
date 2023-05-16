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
    /// Логика взаимодействия для SubSt1.xaml
    /// </summary>
    public partial class SubSt1 : Window
    {
        

        public SubSt1()
        {
            
            InitializeComponent();
            MainWindow main = this.Owner as MainWindow;
            //if (main.par == 1)
            //{ }
            //main.textBlock1.Text = "123";
            //textBox1.Text = main.powObj1.Name;
        }

        public void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            string ps = textBox1.Text;
            if (main != null)
            {
                if (main.par == 1)
                {
                    main.powObj1.NamePO = textBox1.Text;                   
                }
                else if (main.par == 2)
                {
                    main.powObj2.NamePO = textBox1.Text;
                }
                this.Close();
            }

        }
    }
}
