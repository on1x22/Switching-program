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
    /// Логика взаимодействия для New_equipment2.xaml
    /// </summary>
    public partial class New_equipment2 : Window
    {
        bool state_eq;
        string name_eq;
        string type_eq;

        public New_equipment2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            if (textBox1.Text != "")
            {
                name_eq = textBox1.Text;
                if ((bool)checkBox1.IsChecked)
                {
                    state_eq = true;
                }
                else state_eq = false;
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        type_eq = "Switch";
                        break;
                    case 1:
                        type_eq = "Disconnector";
                        break;
                    case 2:
                        type_eq = "GroundDisconnector";
                        break;
                }              
                
                PowerObject.Equipment equip = new PowerObject.Equipment();
                equip.nameEquip = name_eq;
                equip.stateEquip = state_eq;
                equip.typeEquip = type_eq;
                if (/*main*/MainWindow.equipKeywords.ContainsKey(name_eq))
                {
                    MessageBox.Show("Оборудование с данным названием уже существует. Измените название оборудования");
                }
                else
                {
                    /*main*/MainWindow.equipKeywords.Add(name_eq, FontWeights.Bold);
                    equip.NamePO = main.listOfPowerObjects[main.numb].NamePO;
                    equip.isUsed = main.listOfPowerObjects[main.numb].isUsed;
                    main.One_listEquipment.Add(equip);
                    this.Close();
                }
            }
        }
    }
}
