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
    /// Логика взаимодействия для AddAction.xaml
    /// </summary>
    public partial class AddAction : Window
    {
        public AddAction()
        {
            InitializeComponent();
        }
        List<string> powObjectsList = new List<string>();
        public List<PowerObject.Equipment> listOfSelectPO = new List<PowerObject.Equipment>();

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            comboBox3.IsEnabled = false;
            MainWindow main = this.Owner as MainWindow;
            switch (main.positionOfSelectedItem.Count)
            {
                case 3:
                    if (comboBox1.SelectedValue.ToString() == main.comboBox3.Text)
                    {
                        checkBox1.IsEnabled = true;
                        checkBox1.IsChecked = true;
                        comboBox3.IsEnabled = false;
                        //comboBox2.SelectedIndex = 0;
                        string command = "Проверить фиксацию ОТС " + main.textBox2.Text + " в положении «Отключено» в ОИК СК-2007, при несоответствии зафиксировать вручную";
                        comboBox2.Items.Add(command);
                    }
                    else
                    {
                        //comboBox3.IsEnabled = true;
                        powObjectsList.Clear();
                        comboBox2.Items.Add("На " + comboBox1.Text);
                        powObjectsList.Add(comboBox1.Text);
                        for (int i = 0; i < main.listOfPowerObjects.Count; i++)
                        {
                            if (comboBox1.SelectedValue.ToString() != main.listOfPowerObjects[i].NamePO)
                            {
                                string command = "На " + main.listOfPowerObjects[i].NamePO + " отключен линейный разъединитель ";
                                comboBox2.Items.Add(command);
                                powObjectsList.Add(main.listOfPowerObjects[i].NamePO);
                            }
                        }
                    }
                    break;
                case 4:
                    checkBox1.IsEnabled = true;
                    checkBox1.IsChecked = true;
                    comboBox3.IsEnabled = true;
                    string nameofPO = comboBox1.Text;
                    nameofPO = nameofPO.Substring(3);
                    comboBox2.Items.Add("Отключить выключатель ");
                    comboBox2.Items.Add("Снять оперативный ток с цепей управления выключателем ");
                    comboBox2.Items.Add("Отключить линейный разъединитель ");
                    comboBox2.Items.Add("На привод линейного разъединителя ^ вывесить плакат «Не включать! Работа на линии»");                    
                    comboBox2.Items.Add("Включить заземляющие ножи ");
                    comboBox2.Items.Add("Проверить отсутствие напряжения на " + main.textBox2.Text);
                    comboBox2.Items.Add("Подтвердить принятие мер, препятствующих подаче напряжения на " + main.textBox2.Text + " вследствие ошибочного или самопроизвольного включения коммутационных аппаратов");                    
                    for (int i = 0; i < main.One_listEquipment.Count; i++)
                    {
                        if (nameofPO == main.One_listEquipment[i].NamePO)
                        {
                            comboBox3.Items.Add(main.One_listEquipment[i].nameEquip);
                        }
                    }
                    break;
            }
        }

        private void comboBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            switch (main.positionOfSelectedItem.Count)
            {
                case (1):
                    richTextBox1.Document.Blocks.Clear();
                    richTextBox1.AppendText(comboBox1.SelectedValue.ToString());
                    break;
            }

        }

        private void comboBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            switch (main.positionOfSelectedItem.Count)
            {
                case (2):
                    richTextBox1.Document.Blocks.Clear();
                    richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
                    break;
                case 3:
                    if (main.positionOfSelectedItem.Count == 3 && comboBox1.Text == main.comboBox3.Text)     // Если в дереве выбран узел РДУ/ОДУ
                    {
                        richTextBox1.Document.Blocks.Clear();
                        richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
                    }
                    else if (main.positionOfSelectedItem.Count == 3 && comboBox1.Text != main.comboBox3.Text)
                    {
                        richTextBox1.Document.Blocks.Clear();
                        comboBox3.Items.Clear();
                        if (comboBox2.SelectedIndex == 0)
                        {
                            checkBox1.IsEnabled = false;
                            checkBox1.IsChecked = false;
                            comboBox3.IsEnabled = false;
                            richTextBox1.Document.Blocks.Clear();
                            richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
                        }
                        else
                        {
                            checkBox1.IsEnabled = true;
                            comboBox3.IsEnabled = true;
                            for (int i = 0; i < main.One_listEquipment.Count; i++)
                            {
                                if (comboBox2.SelectedIndex != 0 && powObjectsList[comboBox2.SelectedIndex] == main.One_listEquipment[i].NamePO)
                                {
                                    comboBox3.Items.Add(main.One_listEquipment[i].nameEquip);
                                }
                            }
                        }
                    }
                    break;
                case 4:
                    richTextBox1.Document.Blocks.Clear();
                    int ft = comboBox2.Items.Count;
                    if (comboBox2.SelectedIndex == ft - 1 || comboBox2.SelectedIndex == ft - 2)
                    {
                        comboBox3.IsEnabled = false;
                        richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
                    }
                    else
                    {
                        comboBox3.IsEnabled = true;
                        if (comboBox3.SelectedIndex >= 0)
                        {
                            string richtext = comboBox2.SelectedValue.ToString();
                            string compar = comboBox2.SelectedValue.ToString();
                            richtext = richtext.Replace("^", comboBox3.SelectedValue.ToString());
                            if (richtext != compar)
                            {
                                richTextBox1.AppendText(richtext);
                            }
                            else richTextBox1.AppendText(comboBox2.SelectedValue.ToString() + comboBox3.SelectedValue.ToString());
                        }
                    }
                    break;
            }

        }

        private void comboBox3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            richTextBox1.Document.Blocks.Clear();
            switch (main.positionOfSelectedItem.Count)
            {
                case 3:
                    richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
                    if (comboBox3.SelectedItem != null)
                    {
                        richTextBox1.AppendText(comboBox3.SelectedValue.ToString());
                    }
                    break;
                case 4:
                    string richtext = comboBox2.SelectedValue.ToString();
                    string compar = comboBox2.SelectedValue.ToString();
                    richtext = richtext.Replace("^", comboBox3.SelectedValue.ToString());
                    if (richtext != compar)
                    {
                        richTextBox1.AppendText(richtext);
                    }
                    else richTextBox1.AppendText(comboBox2.SelectedValue.ToString() + comboBox3.SelectedValue.ToString());
                    break;
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            /*if (main.actionsList.Count != 0 && main.treeView2.SelectedItem is FirstLevelClass || main.treeView2.SelectedItem == null)
            {
                main.actionsList.Add(new FirstLevelClass(comboBox1.SelectedValue.ToString()));
                this.Close();
            }
            else if (main.treeView2.SelectedItem is SecondLevelClass)
            {
                var cat12 = new SecondLevelClass(comboBox2.SelectedValue.ToString());
                switch (main.positionOfSelectedItem.Count)
                {
                    case 1:
                        break;
                    case 2:
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList.Insert(main.positionOfSelectedItem[1]+1,cat12);
                        break;
                    case 3:
                        break;
                    case 4:
                        break;
                }
                this.Close();
            }*/
            switch (main.positionOfSelectedItem.Count)
            {
                case 1:             // Если в дереве добавляется узел первого уровня
                    string richText0 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText0 = richText0.Remove(richText0.Length - 2);
                    var sat = new FirstLevelClass(richText0);
                    sat.flCommand = comboBox1.SelectedValue.ToString();
                    main.actionsList.Insert(main.positionOfSelectedItem[0] + 1, sat);                    
                    break;
                case 2:             // Если в дереве добавляется узел второго уровня
                    string richText1 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText1 = richText1.Remove(richText1.Length - 2);
                    var cat = new SecondLevelClass(richText1);
                    cat.slCommand = comboBox2.SelectedValue.ToString();
                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList.Insert(main.positionOfSelectedItem[1] + 1,cat);                    
                    break;
                case 3:             // Если в дереве добавляется узел третьего уровня
                    if (comboBox1.SelectedValue.ToString() == main.comboBox3.Text)
                    {
                        string richText = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                        richText = richText.Remove(richText.Length - 2);
                        var cats = new ThirdLevelClass(richText);
                        if ((bool)checkBox1.IsChecked)
                        {
                            cats.isNumerated = true;
                            cats.itemNumber = "5.~";
                            cats.TlName = cats.itemNumber + cats.TlName;                            
                        }
                        else
                        {
                            cats.isNumerated = false;                            
                        }
                        cats.tlCommand = comboBox2.SelectedValue.ToString();
                        cats.isConsistEquip = false;
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList.Insert(main.positionOfSelectedItem[2] + 1,cats);
                    }
                    else
                    {
                        richTextBox1.SelectAll();
                        //string sdgs = richTextBox1.Selection.Text;
                        string richText = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                        richText = richText.Remove(richText.Length - 2);

                        var cats = new ThirdLevelClass(richText);
                        if ((bool)checkBox1.IsChecked)
                        {
                            cats.isNumerated = true;
                            cats.itemNumber = "5.~";
                            cats.TlName = cats.itemNumber + cats.TlName;
                            cats.isConsistEquip = true;
                        }
                        else
                        {
                            cats.isNumerated = false;
                            cats.isConsistEquip = false;
                        }
                        if (comboBox2.SelectedIndex != 0)
                        {
                            cats.equipmentName = comboBox3.SelectedValue.ToString();
                        }
                        cats.equipmentName = comboBox3.SelectedValue.ToString();
                        
                        //main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList.Add(cats);
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList.Insert(main.positionOfSelectedItem[2] + 1, cats);
                    }
                    break;
                case 4:             // Если в дереве добавляется узел четвертого уровня
                    richTextBox1.SelectAll();
                    string richText3 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText3 = richText3.Remove(richText3.Length - 2);
                    var catt = new FourthLevelClass(richText3);
                    if ((bool)checkBox1.IsChecked)
                    {
                        catt.isNumerated = true;
                        catt.itemNumber = "5.~";
                        catt.FlName = catt.itemNumber + catt.FlName;
                    }
                    else catt.isNumerated = false;
                    catt.fourthlCommand = comboBox2.SelectedValue.ToString();
                    if (comboBox2.SelectedIndex != comboBox2.Items.Count - 1 && comboBox2.SelectedIndex != comboBox2.Items.Count - 2)
                    {
                        catt.equipmentName = comboBox3.SelectedValue.ToString();
                        catt.isConsistEquip = true;
                    }                    
                    else catt.isConsistEquip = false;
                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList.Insert(
                        main.positionOfSelectedItem[3] + 1,catt);
                    break;
            }
            this.Close();
        }

        
    }
}
