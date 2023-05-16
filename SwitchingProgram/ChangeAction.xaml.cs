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
    /// Логика взаимодействия для ChangeAction.xaml
    /// </summary>
    public partial class ChangeAction : Window
    {
        public ChangeAction()
        {
            InitializeComponent();
        }
        List<string> powObjectsList = new List<string>();
        public List<PowerObject.Equipment> listOfSelectPO = new List<PowerObject.Equipment>();

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            checkBox1.IsEnabled = false;
            comboBox3.IsEnabled = false;
            switch (main.positionOfSelectedItem.Count)
            {
                case 1:
                    break;
                case 2:
                    break;
                case 3:
                    if (comboBox1.SelectedValue.ToString() == main.comboBox3.Text)
                    {
                        checkBox1.IsEnabled = true;
                        checkBox1.IsChecked = true;
                        //comboBox3.IsEnabled = false;
                        //comboBox2.SelectedIndex = 0;
                        string command = "Проверить фиксацию ОТС " + main.textBox2.Text + " в положении «Отключено» в ОИК СК-2007, при несоответствии зафиксировать вручную";
                        comboBox2.Items.Add(command);
                        comboBox2.SelectedIndex = 0;

                        richTextBox1.Document.Blocks.Clear();
                        richTextBox1.AppendText(main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList[main.positionOfSelectedItem[2]].TlName);
                    }
                    else
                    {
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
                        for (int i = 0; i < comboBox2.Items.Count; i++)
                        {
                            if (comboBox2.Items[i].ToString() == main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                            ThirdLevelList[main.positionOfSelectedItem[2]].tlCommand)
                            {
                                comboBox2.SelectedIndex = i;
                            }
                        }
                        if (comboBox2.SelectedIndex == 0)
                        {
                            checkBox1.IsEnabled = false;
                            checkBox1.IsChecked = false;
                            comboBox3.IsEnabled = false;
                        }
                        else
                        {
                            checkBox1.IsEnabled = true;
                            checkBox1.IsChecked = main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                                ThirdLevelList[main.positionOfSelectedItem[2]].isNumerated;
                            comboBox3.IsEnabled = true;
                            for (int i = 0; i < comboBox3.Items.Count; i++)
                            {
                                if (comboBox3.Items[i].ToString() == main.actionsList[main.positionOfSelectedItem[0]].
                                        SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName)
                                {
                                    comboBox3.SelectedIndex = i;
                                }
                            }
                        }
                    }
                    richTextBox1.Document.Blocks.Clear();
                    richTextBox1.AppendText(main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList[main.positionOfSelectedItem[2]].TlName);
                    break;
                case 4:
                    checkBox1.IsEnabled = true;
                    checkBox1.IsChecked = true;
                    comboBox3.IsEnabled = true;

                    comboBox1.Items.Add(main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList[main.positionOfSelectedItem[2]].TlName);
                    comboBox1.SelectedIndex = 0;

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
                    for (int i = 0; i < comboBox2.Items.Count; i++)
                    {
                        if (comboBox2.Items[i].ToString() == main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                                ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].fourthlCommand)
                        {
                            comboBox2.SelectedIndex = i;
                        }
                    }
                    checkBox1.IsChecked = main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                            ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].isNumerated;
                    for (int i = 0; i < comboBox3.Items.Count; i++)
                    {
                        if (comboBox3.Items[i].ToString() == main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                                ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].equipmentName)
                        {
                            comboBox3.SelectedIndex = i;
                        }
                    }
                    richTextBox1.Document.Blocks.Clear();
                    richTextBox1.AppendText(main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                        ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].FlName);
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
                case 2:
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
            /*richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
            if (comboBox3.SelectedItem != null)
            {
                richTextBox1.AppendText(comboBox3.SelectedValue.ToString());
            }*/
            MainWindow main = this.Owner as MainWindow;
            richTextBox1.Document.Blocks.Clear();
            if (main.positionOfSelectedItem.Count == 4)
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
            else
            {
                //richTextBox1.Document.Blocks.Clear();
                richTextBox1.AppendText(comboBox2.SelectedValue.ToString());
                if (comboBox3.SelectedItem != null)
                {
                    richTextBox1.AppendText(comboBox3.SelectedValue.ToString());
                }
            }
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            switch (main.positionOfSelectedItem.Count)
            {
                case 1:                 // Если в дереве изменяются параметры узла первого уровня
                    string richText0 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText0 = richText0.Remove(richText0.Length - 2);
                    //var sat = new FirstLevelClass(richText0);
                    //sat.flCommand = comboBox1.SelectedValue.ToString();
                    //main.actionsList.Insert(main.positionOfSelectedItem[0] + 1, sat);
                    main.actionsList[main.positionOfSelectedItem[0]].FlName = richText0;
                    main.actionsList[main.positionOfSelectedItem[0]].flCommand = comboBox1.SelectedValue.ToString();                    
                    break;
                case 2:
                    string richText1 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText1 = richText1.Remove(richText1.Length - 2);
                    if (comboBox1.SelectedValue.ToString() != main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].slCommand)
                    {                        
                        if (MessageBox.Show("При изменении энергообъекта будут удалены все дочерние элементы. Продолжить?", "Изменение параметров узла",
                                MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK)
                        {
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList.Clear();
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].SlName = richText1;
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].slCommand = comboBox1.SelectedValue.ToString();
                        }                      
                    }                    
                    break;
                case 3:
                    string richText2 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText2 = richText2.Remove(richText2.Length - 2);
                    if (comboBox2.SelectedValue.ToString() != main.actionsList[main.positionOfSelectedItem[0]].
                                SecondLevelList[main.positionOfSelectedItem[1]].ThirdLevelList[main.positionOfSelectedItem[2]].tlCommand)
                    {
                        if (main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].tlCommand == comboBox2.Items[0].ToString())      // Если ранее был выбран пункт №0, то появляется окно с вопросом
                        { 
                            if (MessageBox.Show("При изменении параметров будут удалены все дочерние элементы. Продолжить?", "Изменение параметров узла",
                                    MessageBoxButton.OKCancel, MessageBoxImage.Information) == MessageBoxResult.OK)
                            {
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList.Clear();

                                bool chek = false;
                                if ((bool)checkBox1.IsChecked)
                                {
                                    chek = true;
                                }
                                else chek = false;
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isNumerated = chek;
                                if (chek == true)
                                {
                                    richText2 = "5.~" + richText2;
                                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                        ThirdLevelList[main.positionOfSelectedItem[2]].itemNumber = "5.~";
                                }
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].TlName = richText2;
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].tlCommand = comboBox2.SelectedValue.ToString();

                                if (comboBox2.SelectedIndex == 0)
                                {
                                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = null;
                                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isConsistEquip = false;
                                }
                                else
                                {
                                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                      ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = comboBox3.SelectedValue.ToString();
                                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isConsistEquip = true;
                                }
                            }
                        }
                        else
                        {
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList.Clear();

                            bool chek = false;
                            if ((bool)checkBox1.IsChecked)
                            {
                                chek = true;
                            }
                            else chek = false;
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].isNumerated = chek;
                            if (chek == true)
                            {
                                richText2 = "5.~" + richText2;
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].itemNumber = "5.~";
                            }
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].TlName = richText2;
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].tlCommand = comboBox2.SelectedValue.ToString();

                            if (comboBox2.SelectedIndex == 0)
                            {
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = null;
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isConsistEquip = false;
                            }
                            else
                            {
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                  ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = comboBox3.SelectedValue.ToString();
                                main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isConsistEquip = true;
                            }
                        }
                    }
                    else
                    {
                        bool chek = false;
                        if ((bool)checkBox1.IsChecked)
                        {
                            chek = true;
                        }
                        else chek = false;
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                            ThirdLevelList[main.positionOfSelectedItem[2]].isNumerated = chek;
                        if (chek == true)
                        {
                            richText2 = "5.~" + richText2;
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].itemNumber = "5.~";
                            
                        }
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].TlName = richText2;
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].tlCommand = comboBox2.SelectedValue.ToString();
                        
                        if (comboBox2.SelectedIndex == 0)
                        {
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                            ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = null;
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isConsistEquip = false;
                        }
                        else
                        {
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                              ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = comboBox3.SelectedValue.ToString();
                            main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                    ThirdLevelList[main.positionOfSelectedItem[2]].isConsistEquip = true;
                        }
                    }
                        break;
                case 4:
                    string richText3 = new TextRange(richTextBox1.Document.ContentStart, richTextBox1.Document.ContentEnd).Text;
                    richText3 = richText3.Remove(richText3.Length - 2);
                    bool chek3 = false;
                    if ((bool)checkBox1.IsChecked)
                    {
                        chek3 = true;
                    }
                    else chek3 = false;
                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                        ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].isNumerated = chek3;
                    if (chek3 == true)
                    {
                        richText3 = "5.~" + richText3;
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                            ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].itemNumber = "5.~";

                    }
                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].FlName = richText3;
                    main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                                ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].fourthlCommand = comboBox2.SelectedValue.ToString();
                    if (comboBox2.SelectedIndex == comboBox2.Items.Count - 1 || comboBox2.SelectedIndex == comboBox2.Items.Count - 2)
                    {
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                              ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].equipmentName = null;
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                              ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].isConsistEquip = false;
                    }
                    else
                    {
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                              ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].equipmentName = comboBox3.SelectedValue.ToString();
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                              ThirdLevelList[main.positionOfSelectedItem[2]].FourthLevelList[main.positionOfSelectedItem[3]].isConsistEquip = true;
                    }
                    /*if (comboBox2.SelectedIndex == 0)
                    {
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                        ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = null;
                    }
                    else
                    {
                        main.actionsList[main.positionOfSelectedItem[0]].SecondLevelList[main.positionOfSelectedItem[1]].
                          ThirdLevelList[main.positionOfSelectedItem[2]].equipmentName = comboBox3.SelectedValue.ToString();
                    }*/
                    break;
            }
            this.Close();
        }
    }
}
