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
    /// Логика взаимодействия для Change_equipment_parameters.xaml
    /// </summary>
    public partial class Change_equipment_parameters : Window
    {
        bool state_eq;
        string name_eq;
        string type_eq;
        string oldEqName;

        public Change_equipment_parameters()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            oldEqName = textBox1.Text;
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            for (int i = 0; i < main.One_listEquipment.Count; i++)
            {
                if (main.One_listEquipment[i].nameEquip == oldEqName)
                {
                    main.One_listEquipment[i].nameEquip = textBox1.Text;

                    if ((bool)checkBox1.IsChecked)
                    {
                        main.One_listEquipment[i].stateEquip = true;
                    }
                    else main.One_listEquipment[i].stateEquip = false;

                    switch (comboBox1.SelectedIndex)
                    {
                        case 0:
                            main.One_listEquipment[i].typeEquip = "Switch";
                            break;
                        case 1:
                            main.One_listEquipment[i].typeEquip = "Disconnector";
                            break;
                        case 2:
                            main.One_listEquipment[i].typeEquip = "GroundDisconnector";
                            break;
                    }
                }
            }
            this.Close();
        }        
    }
}
