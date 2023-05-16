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
    /// Логика взаимодействия для Change_name_of_power_object.xaml
    /// </summary>
    public partial class Change_name_of_power_object : Window
    {
        string old_namePO;
        string old_organisationPO;

        public Change_name_of_power_object()
        {
            InitializeComponent();
        }
                      

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                old_namePO = main.listOfPowerObjects[main.numb].NamePO;
                old_organisationPO = main.listOfPowerObjects[main.numb].organisationPO;
                main.listOfPowerObjects[main.numb].NamePO = textBox1.Text;
                main.listOfPowerObjects[main.numb].organisationPO = textBox2.Text;
                if ((bool)checkBox1.IsChecked)
                {
                    main.listOfPowerObjects[main.numb].isUsed = true;
                }
                else main.listOfPowerObjects[main.numb].isUsed = false;
                for (int i = 0; i < main.One_listEquipment.Count; i++)
                {
                    if (main.One_listEquipment[i].NamePO == old_namePO )
                    {
                        main.One_listEquipment[i].NamePO = textBox1.Text;                        
                    }
                }
                for (int i = 0; i < main.One_listEquipment.Count; i++)
                {
                    if (main.One_listEquipment[i].organisationPO == old_organisationPO)
                    {
                        if (main.One_listEquipment[i].NamePO == old_namePO)
                        {
                            main.One_listEquipment[i].organisationPO = textBox2.Text;
                        }
                    }
                }
                this.Close();
            }
            else MessageBox.Show("Не задано название энергообъекта или его организация");
        }
    }
}
