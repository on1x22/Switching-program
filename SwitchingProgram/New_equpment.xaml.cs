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
    /// Логика взаимодействия для New_equpment.xaml
    /// </summary>
    public partial class New_equpment : Window
    {

        public New_equpment()
        {
            InitializeComponent();
            //MainWindow main = this.Owner as MainWindow;
        }
        bool state_eq;
        string name_eq;

private void button1_Click(object sender, RoutedEventArgs e)
        {
            

            if (textBox1.Text != String.Empty)
            {
                name_eq = textBox1.Text;
                if ((bool)checkBox1.IsChecked)
                {
                    state_eq = true;
                }
                else state_eq = false;

                MainWindow main = this.Owner as MainWindow;
                PowerObject.Equipment powEquip = new PowerObject.Equipment();
                powEquip.nameEquip = name_eq;
                powEquip.stateEquip = state_eq;
                if (/*main*/MainWindow.equipKeywords.ContainsKey(name_eq))
                {
                    MessageBox.Show("Оборудование с данным названием уже существует. Измените название оборудования");
                }
                else
                {
                    /*main.equipKeywords.Add(name_eq, FontWeights.Bold);*/
                    MainWindow.equipKeywords.Add(name_eq, FontWeights.Bold);
                    //main.fordict = name_eq;
                    if (main.wind == 0)
                    {
                        powEquip.NamePO = main.powObj1.NamePO;
                        main.listEquipment1.Add(powEquip);
                    }
                    else if (main.wind == 1)
                    {
                        powEquip.NamePO = main.powObj2.NamePO;
                        main.listEquipment2.Add(powEquip);
                    }
                    //main.perechen.Add(powEquip); 
                    this.Close();
                }

               
            }
            else MessageBox.Show("Не введено название оборудования", "Ошибка", MessageBoxButton.OK);
        }
    }
}
