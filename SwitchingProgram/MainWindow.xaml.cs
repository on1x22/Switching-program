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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Collections.ObjectModel;
using System.Collections;
using System.IO;
using Microsoft.Win32;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using System.Threading;
using System.Reflection;

namespace WpfApplication1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    
        //The main window class
        public partial class MainWindow : System.Windows.Window
        {
        //public MainPowerObject mPowObj = new MainPowerObject();

        OpenFileDialog openDialog = new OpenFileDialog();
        SaveFileDialog saveDialog = new SaveFileDialog();

        string filename;


        public PowerObject powObj1 = new PowerObject(); // При запуске объявляем, что ПП создается для 2 подстанций
        public PowerObject powObj2 = new PowerObject();
        /*public PowerObject powObj3 = new PowerObject();
        public PowerObject powObj4 = new PowerObject();
        public PowerObject powObj5 = new PowerObject();
        public PowerObject powObj6 = new PowerObject();
        public PowerObject powObj7 = new PowerObject();*/

        ContextMenu contextMenu = new ContextMenu();

        //int fourtlpoz;
        public List<int> positionOfSelectedItem = new List<int>();

        public int par;
        public int wind;
        public int numb;        

        public bool can_change_list_of_objects; // Переменая, которая определяет можно ли изменять количество подстанций или нет
        public int num_obj;                     // Переменая, которая определяет количество энергообъектов в программе
        public int num_obj_start;

        public bool stateStart;
        public string nameStart;
        public bool checker;

        int iNumber;

        public ProgramOptions progOption = new ProgramOptions();

        public MainParametrsOfSwitchingProgram mainParamsOfSP = new MainParametrsOfSwitchingProgram();

        public List<PowerObject.Equipment> listEquipment1 = new List<PowerObject.Equipment>();
        public List<PowerObject.Equipment> listEquipment2 = new List<PowerObject.Equipment>();

        public List<Org_arrangs> listOfOrgArrs = new List<Org_arrangs>();
        public List<Org_arrangs> listOfOldOrgArrs = new List<Org_arrangs>();

        public List<Personal> listOfPersonal = new List<Personal>();
        public List<Personal> listOfPersonalTemp = new List<Personal>();

        public static Dictionary<string, FontWeight> equipKeywords = new Dictionary<string, FontWeight>(); // Словарь оборудования для выделения в TreeView
        //public string fordict;

        public List<PowerObject> listOfPowerObjects = new List<PowerObject>();
        //public List<PowerObject> listOfUsingPowerObjects = new List<PowerObject>(); // Список энергообъектов, которые включены в обработку
        public List<PowerObject.Equipment> One_listEquipment = new List<PowerObject.Equipment>();
        public List<PowerObject.Equipment> listOfSelectedPowerObject = new List<PowerObject.Equipment>();

        ObservableCollection<Node> nodes;
        public ObservableCollection<FirstLevelClass> actionsList = new ObservableCollection<FirstLevelClass>();

        public static Dictionary<string, FontWeight> tvKeywords = new Dictionary<string, FontWeight>();
        public MainWindow()
            {
            DataContext = this;
            /*powObj1.NamePO = "ПС1"; // Объявляются стандартные имена подстанций
            powObj2.NamePO = "ПС2";*/
            

            InitializeComponent();
            


            comboBox3.SelectedIndex = 2;

            num_obj = 2;
            num_obj_start = num_obj;
            /*button2.DataContext = powObj1;
            button3.DataContext = powObj2;*/

            //listBox1.Items.Clear();

            can_change_list_of_objects = true;

            checker = false;
            /*Binding binding = new Binding();

            binding.ElementName = "powObj1"; // элемент-источник
            binding.Path = new PropertyPath("NamePO"); // свойство элемента-источника
            button2.SetBinding(ContentPresenter.ContentProperty, binding); // установка привязки для элемента-приемника*/

        }
        
        /*// Надо убрать когда перейду к TreeView
        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            // \u00A0 - это НЕРАЗРЫВНЫЙ ПРОБЕЛ
            tvKeywords.Add("Иван\u00A0ан", FontWeights.Bold);
            tvKeywords.Add("Ольга", FontWeights.ExtraBold);
            tvKeywords.Add("В\u00A01", FontWeights.ExtraBold);
            tvKeywords.Add("В1", FontWeights.ExtraBold);

            string ing = "В\u00A01";


            nodes = new ObservableCollection<Node>
        {
            new Node
            {
                Name ="Иванов Иван\u00A0ан Иванович",
                Nodes = new ObservableCollection<Node>
                {                    
                    new Node
                    {
                        Name ="Иванова Юлия " + ing +" Ивановна",
                        Nodes = new ObservableCollection<Node>
                        {
                            new Node {Name="Иванов Иван Петрович" },                            
                            new Node
                            {
                                Name = "Иванова Ольга Петровна",
                                Nodes = new ObservableCollection<Node>
                                {
                                    new Node {Name="Иванов Иван Иванович" },
                                }
                            }                    
                        }
                    }
                }
            },
            new Node
            {
                Name ="Петров Петр Иванович",
                Nodes = new ObservableCollection<Node>
                {
                    new Node {Name="Петров Антон Петрович" },
                    new Node {Name="Петров Иван Петрович" }                    
                }
            }            
        };
            treeView1.ItemsSource = nodes;
        }*/



        /*// Открытие окна изменения названия ПС1
        public void button2_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Right)
            {                
                SubSt1 subSt1 = new SubSt1();
                subSt1.Owner = this;
                subSt1.textBox1.Text = powObj1.NamePO;
                par = 1;
                subSt1.ShowDialog();
                button2.Content = powObj1.NamePO;
                int kk = 0;
            }
        }*/

        /*// Открытие окна изменения названия ПС2
        private void button3_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Right)
            {
                SubSt1 subSt2 = new SubSt1();
                subSt2.Owner = this;
                subSt2.textBox1.Text = powObj2.NamePO;
                par = 2;
                subSt2.ShowDialog();
                button3.Content = powObj2.NamePO;
                int kk = 0;
            }
        }*/



        /*// Создание оборудования на ПС1
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            wind = 0;                                               // Показывает, что создавать надо в левую ПС
            New_equpment nEq = new New_equpment();
            nEq.Owner = this;
            nEq.ShowDialog();
            listView1.ItemsSource = listEquipment1;
            listView1.Items.Refresh();*/
            /*if (equipKeywords.ContainsKey(fordict))
            {
                MessageBox.Show("Nazvanie est'");
            }
            else
            {
                equipKeywords.Add(fordict, FontWeights.Bold);
            }*/
       /*}*/
        
        /*// Создание оборудования на ПС2
        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {            
            wind = 1;                                               // Показывает, что создавать надо в правую ПС
            New_equpment nEq = new New_equpment();
            nEq.Owner = this;
            nEq.ShowDialog();
            //DataContext = this;
            listView2.ItemsSource = listEquipment2;
            listView2.Items.Refresh();
        }*/



        // Открытие меню настроек
        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            Options options = new Options();
            options.Owner = this;
            options.textBox1.Text = progOption.nameSP;
            options.textBox2.Text = progOption.roditPadezh;
            options.ShowDialog();
            /*if (num_obj != num_obj_start)
            {
                switch (num_obj)
                {
                    case 2:
                        powObj1.NamePO = "ПС1";
                        powObj2.NamePO = "ПС2";   
                        break;
                    case 3:
                        powObj1.NamePO = "ПС1";
                        powObj2.NamePO = "ПС2";
                        powObj3.NamePO = "ПС3";
                        break;
                    case 4:
                        powObj1.NamePO = "ПС1";
                        powObj2.NamePO = "ПС2";
                        powObj3.NamePO = "ПС3";
                        powObj4.NamePO = "ПС4";
                        break;
                    case 5:
                        powObj1.NamePO = "ПС1";
                        powObj2.NamePO = "ПС2";
                        powObj3.NamePO = "ПС3";
                        powObj4.NamePO = "ПС4";
                        powObj5.NamePO = "ПС5";
                        break;
                    case 6:
                        powObj1.NamePO = "ПС1";
                        powObj2.NamePO = "ПС2";
                        powObj3.NamePO = "ПС3";
                        powObj4.NamePO = "ПС4";
                        powObj5.NamePO = "ПС5";
                        powObj6.NamePO = "ПС6";
                        break;
                    case 7:
                        powObj1.NamePO = "ПС1";
                        powObj2.NamePO = "ПС2";
                        powObj3.NamePO = "ПС3";
                        powObj4.NamePO = "ПС4";
                        powObj5.NamePO = "ПС5";
                        powObj6.NamePO = "ПС6";
                        powObj7.NamePO = "ПС7";
                        break;
                }
            }*/
        }



        /*// Удаление оборудования с ПС1
        private void MenuItem_Click_6(object sender, RoutedEventArgs e)
        {
            if (listView1.SelectedIndex >= 0)
            {
                int itm = Convert.ToInt32(listView1.SelectedIndex);                
                equipKeywords.Remove(listEquipment1[itm].nameEquip);
                listEquipment1.RemoveAt(itm);
                listView1.Items.Refresh();
            }
        }*/

        /*// Удаление оборудования с ПС2
        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            if (listView2.SelectedIndex >= 0)
            {
                int itm = Convert.ToInt32(listView2.SelectedIndex);
                equipKeywords.Remove(listEquipment2[itm].nameEquip);
                listEquipment2.RemoveAt(itm);
                listView2.Items.Refresh();
            }
        }*/



        /*// Изменение параметров у ПС1
        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            if (listView1.SelectedIndex >= 0)
            {
                wind = 0;
                Changing changePar = new Changing();
                changePar.Owner = this;
                changePar.checkBox1.IsChecked = listEquipment1[listView1.SelectedIndex].stateEquip;
                stateStart = listEquipment1[listView1.SelectedIndex].stateEquip;
                changePar.textBox1.Text = listEquipment1[listView1.SelectedIndex].nameEquip;
                nameStart = listEquipment1[listView1.SelectedIndex].nameEquip;
                changePar.ShowDialog();
                if (checker == true)
                {
                    equipKeywords.Remove(listEquipment1[listView1.SelectedIndex].nameEquip);
                    listEquipment1[listView1.SelectedIndex].nameEquip = nameStart;
                    listEquipment1[listView1.SelectedIndex].stateEquip = stateStart;
                    equipKeywords.Add(listEquipment1[listView1.SelectedIndex].nameEquip, FontWeights.Bold);
                    checker = false;
                }
                listView1.Items.Refresh();
            }
        }*/
                
        /*// Изменение параметров у ПС2 
        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            if (listView2.SelectedIndex >= 0)
            {
                wind = 1;
                Changing changePar = new Changing();
                changePar.Owner = this;
                changePar.checkBox1.IsChecked = listEquipment2[listView2.SelectedIndex].stateEquip;
                stateStart = listEquipment2[listView2.SelectedIndex].stateEquip;
                changePar.textBox1.Text = listEquipment2[listView2.SelectedIndex].nameEquip;
                nameStart = listEquipment2[listView2.SelectedIndex].nameEquip;
                changePar.ShowDialog();
                if (checker == true)
                {
                    equipKeywords.Remove(listEquipment2[listView2.SelectedIndex].nameEquip);
                    listEquipment2[listView2.SelectedIndex].nameEquip = nameStart;
                    listEquipment2[listView2.SelectedIndex].stateEquip = stateStart;
                    equipKeywords.Add(listEquipment2[listView2.SelectedIndex].nameEquip, FontWeights.Bold);
                    checker = false;
                }
                listView2.Items.Refresh();
            }
        }*/

        
        
        // Меню. Открыть файл Оборудования
        private void MenuItem_Click_7(object sender, RoutedEventArgs e)
        {
            openDialog.Filter = "XML files(*.xml)|*.xml|All files(*.*)|*.*";

            Nullable<bool> open_result = openDialog.ShowDialog();

            if (open_result == true)
            {
                // Save document
                filename = openDialog.FileName;

                XDocument xdoc = XDocument.Load(filename);

                foreach (XElement info in xdoc.Element("Root").Elements("Information"))
                {
                    XAttribute actionDate = info.Attribute("actionDate");
                    XAttribute nameLine = info.Attribute("nameLine");
                    XAttribute lineOrganisation = info.Attribute("lineOrganisation");
                    XAttribute typeDo = info.Attribute("typeDO");
                    XAttribute dispOffice = info.Attribute("dispOffice");
                    XAttribute aim = info.Attribute("aim");

                    XAttribute inducedVoltage = info.Attribute("inducedVoltage");
                    XAttribute isUsedARM = info.Attribute("isUsedARM");
                    XAttribute ferroresonance = info.Attribute("ferroresonance");

                    datePicker1.Text = actionDate.Value;
                    textBox1.Text = dispOffice.Value;
                    textBox2.Text = nameLine.Value;
                    textBox3.Text = lineOrganisation.Value;
                    string forAim = aim.Value;
                    switch (forAim)
                    {
                        case "Вывод в ремонт":
                            comboBox1.SelectedIndex = 0;
                            break;
                        case "Вывод в резерв":
                            comboBox1.SelectedIndex = 1;
                            break;
                        case "Ввод из резерва":
                            comboBox1.SelectedIndex = 2;
                            break;
                        case "Ввод в работу":
                            comboBox1.SelectedIndex = 3;
                            break;
                    }
                    string forTypeDO = typeDo.Value;
                    switch (forTypeDO)
                    {
                        case "ИА":
                            comboBox3.SelectedIndex = 0;
                            break;
                        case "ОДУ":
                            comboBox3.SelectedIndex = 1;
                            break;
                        case "РДУ":
                            comboBox3.SelectedIndex = 2;
                            break;                        
                    }

                }

                foreach (XElement equip in xdoc.Element("Root").Elements("Equipment"))
                {
                    XAttribute name_PO = equip.Attribute("namePO");
                    XAttribute organisation_PO = equip.Attribute("organisationPO");
                    XAttribute is_Used = equip.Attribute("isUsed");
                    XAttribute nameEquip = equip.Attribute("nameEquip");
                    XAttribute stateEquip = equip.Attribute("stateEquip");
                    XAttribute typeEquip = equip.Attribute("typeEquip");

                    PowerObject.Equipment powEquip = new PowerObject.Equipment();
                    powEquip.NamePO = name_PO.Value;
                    powEquip.organisationPO = organisation_PO.Value;
                    powEquip.isUsed = Convert.ToBoolean(is_Used.Value);
                    powEquip.nameEquip = nameEquip.Value;
                    powEquip.stateEquip = Convert.ToBoolean(stateEquip.Value);
                    powEquip.typeEquip = typeEquip.Value;
                    One_listEquipment.Add(powEquip);                    
                }

                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    int new_object = 1;
                    PowerObject new_PO = new PowerObject();
                    if (listOfPowerObjects.Count == 0)
                    {
                        new_PO.NamePO = One_listEquipment[i].NamePO;
                        new_PO.organisationPO = One_listEquipment[i].organisationPO;
                        new_PO.isUsed = One_listEquipment[i].isUsed;
                        listOfPowerObjects.Add(new_PO);
                    }
                    else
                    {
                        for (int j = 0; j < listOfPowerObjects.Count; j++)
                        {
                            if (listOfPowerObjects[j].NamePO == One_listEquipment[i].NamePO)
                            {
                                new_object = new_object - 1;                                
                            }
                        }
                        if(new_object != 0)
                        {
                            new_PO.NamePO = One_listEquipment[i].NamePO;
                            new_PO.organisationPO = One_listEquipment[i].organisationPO;
                            new_PO.isUsed = One_listEquipment[i].isUsed;
                            listOfPowerObjects.Add(new_PO);
                            new_object = 1;
                        }
                    }
                }
                /*listBox1.ItemsSource = listOfPowerObjects;
                listBox1.Items.Refresh();*/
                listView4.ItemsSource = listOfPowerObjects;
                listView4.Items.Refresh();
                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    string replSpace = "\u00A0";
                    One_listEquipment[i].nameEquip = One_listEquipment[i].nameEquip.Replace(" ", replSpace);
                    string equipm = One_listEquipment[i].nameEquip;
                    if (equipKeywords.ContainsKey(equipm))
                    {
                        MessageBox.Show("Оборудование с данным названием уже существует. Измените название оборудования");
                    }
                    else
                    {
                        equipKeywords.Add(equipm, FontWeights.Bold);                        
                    }
                }

                
                    // Заполнение организационных мероприятий
                    for (int i = 0; i < listOfPowerObjects.Count; i++)
                {
                    if (listOfPowerObjects[i].isUsed == true)
                    {
                        var oars = new Org_arrangs();
                        oars.NameObj = "На " + listOfPowerObjects[i].NamePO;
                        oars.PObject = listOfPowerObjects[i].NamePO;
                        oars.isWork = false;

                        listOfOrgArrs.Add(oars);
                    }
                }
                var oar = new Org_arrangs();
                oar.NameObj = "На " + textBox2.Text;
                oar.PObject = textBox2.Text;
                oar.isWork = false;
                listOfOrgArrs.Add(oar);

                oar = new Org_arrangs();
                oar.NameObj = "Не будут производиться";
                oar.PObject = "Не будут производиться";
                oar.isWork = false;
                listOfOrgArrs.Add(oar);

                for (int i = 0; i < listOfOrgArrs.Count; i++)
                {
                    var uuu = new Org_arrangs();
                    uuu.NameObj = listOfOrgArrs[i].NameObj;
                    uuu.PObject = listOfOrgArrs[i].PObject;
                    uuu.isWork = listOfOrgArrs[i].isWork;
                    listOfOldOrgArrs.Add(uuu);
                }
                //listOfOldOrgArrs = listOfOrgArrs;
                listView5.ItemsSource = listOfOrgArrs;

                // Заполнение Персонала, участвующего в переключениях
                foreach (XElement persSP in xdoc.Element("Root").Elements("PersonalOfSP"))
                {
                    foreach (XElement organis in persSP.Elements("Organisation"))
                    {
                        var perpers = new Personal("");
                        XAttribute organisation_Of_Personal = organis.Attribute("organisationOfPersonal");

                        perpers.organisationOfPersonal = organisation_Of_Personal.Value;
                        foreach (XElement pers in organis.Elements("Personal"))
                        {
                            var man = new PersonalClass("");

                            XAttribute name_Of_Person = pers.Attribute("nameOfPerson");
                            XAttribute role = pers.Attribute("role");
                            XAttribute action = pers.Attribute("action");

                            man.nameOfPerson = name_Of_Person.Value;
                            man.role = role.Value;
                            man.action = action.Value;

                            perpers.Person.Add(man);
                        }
                        listOfPersonal.Add(perpers);
                    }
                }

                MessageBox.Show("Файл открыт");
            }
        }

        // Меню. Открыть файл Действий
        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            openDialog.Filter = "XML files(*.xml)|*.xml|All files(*.*)|*.*";

            Nullable<bool> open_result = openDialog.ShowDialog();

            if (open_result == true)
            {
                // Save document
                filename = openDialog.FileName;

                //var lev1 = new FirstLevelClass("");




                XDocument xdoc = XDocument.Load(filename);
                foreach (XElement L1element in xdoc.Element("actions").Elements("actionL1"))
                {
                    var lev1 = new FirstLevelClass("");
                    XAttribute nameAttr1 = L1element.Attribute("flName");
                    lev1.FlName = nameAttr1.Value;
                    XAttribute commandAttr1 = L1element.Attribute("flCommand");
                    lev1.flCommand = commandAttr1.Value;
                    //XElement priceElement = phoneElement.Element("price");
                    foreach (XElement L2element in L1element.Elements("actionL2"))
                    {
                        var lev2 = new SecondLevelClass("");
                        XAttribute nameAttr2 = L2element.Attribute("slName");
                        lev2.SlName = nameAttr2.Value;
                        XAttribute commandAttr2 = L2element.Attribute("slCommand");
                        lev2.slCommand = commandAttr2.Value;

                        foreach (XElement L3element in L2element.Elements("actionL3"))
                        {
                            var lev3 = new ThirdLevelClass("");
                            XAttribute nameAttr3 = L3element.Attribute("tlName");
                            lev3.TlName = nameAttr3.Value;
                            XAttribute commandAttr3 = L3element.Attribute("tlCommand");
                            lev3.tlCommand = commandAttr3.Value;
                            XAttribute isNumAttr3 = L3element.Attribute("isNumerated");
                            lev3.isNumerated = Convert.ToBoolean(isNumAttr3.Value);
                            if (isNumAttr3.Value == "true")
                            {
                                XAttribute itemNumAttr3 = L3element.Attribute("itemNumber");
                                lev3.itemNumber = itemNumAttr3.Value;
                            }
                            XAttribute isConsEquipAttr3 = L3element.Attribute("isConsistEquip");
                            lev3.isConsistEquip = Convert.ToBoolean(isConsEquipAttr3.Value);
                            if (isConsEquipAttr3.Value == "true")
                            {
                                XAttribute EquipNameAttr3 = L3element.Attribute("equipmentName");
                                lev3.equipmentName = EquipNameAttr3.Value;
                            }
                            if (isNumAttr3.Value == "false" && isConsEquipAttr3.Value == "false")
                            {
                                foreach (XElement L4element in L3element.Elements("actionL4"))
                                {
                                    var lev4 = new FourthLevelClass("");
                                    XAttribute nameAttr4 = L4element.Attribute("flName");
                                    lev4.FlName = nameAttr4.Value;
                                    XAttribute commandAttr4 = L4element.Attribute("fourthlCommand");
                                    lev4.fourthlCommand = commandAttr4.Value;
                                    XAttribute isNumAttr4 = L4element.Attribute("isNumerated");
                                    lev4.isNumerated = Convert.ToBoolean(isNumAttr4.Value);
                                    if (isNumAttr4.Value == "true")
                                    {
                                        XAttribute itemNumAttr4 = L4element.Attribute("itemNumber");
                                        lev4.itemNumber = itemNumAttr4.Value;
                                    }
                                    XAttribute isConsEquipAttr4 = L4element.Attribute("isConsistEquip");
                                    lev4.isConsistEquip = Convert.ToBoolean(isConsEquipAttr4.Value);
                                    if (isConsEquipAttr4.Value == "true")
                                    {
                                        XAttribute EquipNameAttr4 = L4element.Attribute("equipmentName");
                                        lev4.equipmentName = EquipNameAttr4.Value;
                                    }
                                    lev3.FourthLevelList.Add(lev4);
                                }
                            }
                            lev2.ThirdLevelList.Add(lev3);
                        }
                        lev1.SecondLevelList.Add(lev2);
                    }



                    /*if (nameAttribute != null && companyElement != null && priceElement != null)
                    {
                        Console.WriteLine("Смартфон: {0}", nameAttribute.Value);
                        Console.WriteLine("Компания: {0}", companyElement.Value);
                        Console.WriteLine("Цена: {0}", priceElement.Value);
                    }
                    Console.WriteLine();*/

                    actionsList.Add(lev1);
                }
            }

            treeView2.ItemsSource = actionsList;
        }

        // Меню. Сохранить файл как Оборудование
        private void MenuItem_Click_8(object sender, RoutedEventArgs e)
        {
            saveDialog.Filter = "XML files(*.xml)|*.xml|All files(*.*)|*.*";
            // Show save file dialog box
            Nullable<bool> result = saveDialog.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                filename = saveDialog.FileName;

                XDocument xdoc = new XDocument();
                // создаем корневой элемент
                XElement root = new XElement("Root");

                XElement Information = new XElement("Information");
                XAttribute aim = new XAttribute("aim", comboBox1.Text);
                XAttribute typeDO = new XAttribute("typeDO", comboBox3.Text);
                XAttribute dispOffice = new XAttribute("dispOffice", textBox1.Text);
                XAttribute nameLine = new XAttribute("nameLine", textBox2.Text);
                XAttribute isACLineSegment = new XAttribute("isACLineSegment", mainParamsOfSP.isACLineSegment);
                XAttribute ACLineSegment = new XAttribute("ACLineSegment", "non");
                if (mainParamsOfSP.isACLineSegment == true)
                {
                    ACLineSegment = new XAttribute("ACLineSegment", mainParamsOfSP.ACLineSegment);
                }                
                XAttribute lineOrganisation = new XAttribute("lineOrganisation", textBox3.Text);
                XAttribute actionDate = new XAttribute("actionDate", datePicker1.Text);

                Information.Add(aim);
                Information.Add(typeDO);
                Information.Add(dispOffice);
                Information.Add(nameLine);
                Information.Add(isACLineSegment);
                Information.Add(ACLineSegment);
                Information.Add(lineOrganisation);
                Information.Add(actionDate);

                root.Add(Information);

                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    // создаем элемент
                    XElement equip = new XElement("Equipment");
                    XAttribute typeEquip = new XAttribute("typeEquip", One_listEquipment[i].typeEquip);
                    XAttribute stateEquip = new XAttribute("stateEquip", One_listEquipment[i].stateEquip);
                    XAttribute nameEquip = new XAttribute("nameEquip", One_listEquipment[i].nameEquip);
                    XAttribute namePO = new XAttribute("namePO", One_listEquipment[i].NamePO);
                    XAttribute isUsed = new XAttribute("isUsed", One_listEquipment[i].isUsed);
                    XAttribute organisationPO = new XAttribute("organisationPO", One_listEquipment[i].organisationPO);

                    equip.Add(organisationPO);
                    equip.Add(namePO);
                    equip.Add(isUsed);
                    equip.Add(nameEquip);
                    equip.Add(stateEquip);
                    equip.Add(typeEquip);

                    root.Add(equip);
                }
                

                // Добавление персонала, участвующего в переключениях
                XElement persOfSP = new XElement("PersonalOfSP");
                for (int i = 0; i < listOfPersonal.Count; i++)
                {
                    XElement org = new XElement("Organisation");
                    XAttribute orgOfPers = new XAttribute("organisationOfPersonal", listOfPersonal[i].organisationOfPersonal);
                    org.Add(orgOfPers);

                    for (int j = 0; j < listOfPersonal[i].Person.Count; j++)
                    {
                        XElement pers = new XElement("Personal");
                        XAttribute nameOfPers = new XAttribute("nameOfPerson", listOfPersonal[i].Person[j].nameOfPerson);
                        XAttribute roleOfPers = new XAttribute("role", listOfPersonal[i].Person[j].role);
                        XAttribute actionOfPers = new XAttribute("action", listOfPersonal[i].Person[j].action);

                        pers.Add(nameOfPers);
                        pers.Add(roleOfPers);
                        pers.Add(actionOfPers);

                        org.Add(pers);
                    }
                    persOfSP.Add(org);
                }
                root.Add(persOfSP);

                xdoc.Add(root);

                xdoc.Save(filename);
                 
                /*List<PowerObject.Equipment> commonListOfEquipment = new List<PowerObject.Equipment>();

                if (listEquipment1.Count > 0)
                {
                    commonListOfEquipment.AddRange(listEquipment1.ToArray());
                    if (listEquipment2.Count > 0)
                    {
                        commonListOfEquipment.AddRange(listEquipment2.ToArray());
                    }
                    else
                    {
                        MessageBox.Show("Не все энергообъекты имеют оборудование. Сохранение прервано.");
                    }
                }
                else
                {
                    MessageBox.Show("Не все энергообъекты имеют оборудование. Сохранение прервано.");
                }*/
            }                   
        }

        // Меню. Сохранить файл как Действия
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            saveDialog.Filter = "XML files(*.xml)|*.xml|All files(*.*)|*.*";
            // Show save file dialog box
            Nullable<bool> result = saveDialog.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                filename = saveDialog.FileName;

                XDocument xdoc = new XDocument();
                XElement actions = new XElement("actions");
                XElement actionL1;
                XElement actionL2;
                XElement actionL3;
                XElement actionL4;
                for (int i = 0; i < actionsList.Count; i++)
                {
                    /*XElement*/ actionL1 = new XElement("actionL1");
                    XAttribute actL1attrsName = new XAttribute("flName", actionsList[i].FlName);
                    actionL1.Add(actL1attrsName);
                    XAttribute actL1attrsCommand = new XAttribute("flCommand", actionsList[i].flCommand);
                    actionL1.Add(actL1attrsCommand);
                    if (actionsList[i].SecondLevelList.Count != 0)
                    {
                        for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                        {
                            /*XElement*/ actionL2 = new XElement("actionL2");
                            XAttribute actL2attrsName = new XAttribute("slName", actionsList[i].SecondLevelList[j].SlName);
                            actionL2.Add(actL2attrsName);
                            XAttribute actL2attrsCommand = new XAttribute("slCommand", actionsList[i].SecondLevelList[j].slCommand);
                            actionL2.Add(actL2attrsCommand);
                            if (actionsList[i].SecondLevelList[j].ThirdLevelList.Count != 0)
                            {
                                for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                                {
                                    /*XElement*/ actionL3 = new XElement("actionL3");
                                    XAttribute actL3attrsName = new XAttribute("tlName", actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName);
                                    actionL3.Add(actL3attrsName);
                                    XAttribute actL3attrsCommand = new XAttribute("tlCommand", actionsList[i].SecondLevelList[j].ThirdLevelList[k].tlCommand);
                                    actionL3.Add(actL3attrsCommand);
                                    XAttribute actL3attrsIsNum = new XAttribute("isNumerated", actionsList[i].SecondLevelList[j].ThirdLevelList[k].isNumerated);
                                    actionL3.Add(actL3attrsIsNum);
                                    XAttribute actL3attrsIsConsEquip = new XAttribute("isConsistEquip", actionsList[i].SecondLevelList[j].ThirdLevelList[k].isConsistEquip);
                                    actionL3.Add(actL3attrsIsConsEquip);
                                    if (actionsList[i].SecondLevelList[j].slCommand != comboBox3.Text &&
                                        actionsList[i].SecondLevelList[j].ThirdLevelList[k].isNumerated == true) // Если родитель не РДУ и узел нумеруемый
                                    {
                                        XAttribute actL3attrsEquip = new XAttribute("equipmentName", actionsList[i].SecondLevelList[j].ThirdLevelList[k].equipmentName);
                                        actionL3.Add(actL3attrsEquip);                                        
                                        XAttribute actL3attrsNumber = new XAttribute("itemNumber", actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber);
                                        actionL3.Add(actL3attrsNumber);
                                    }
                                    else if (actionsList[i].SecondLevelList[j].slCommand == comboBox3.Text)     // Если родитель РДУ
                                    {
                                        XAttribute actL3attrsNumber = new XAttribute("itemNumber", actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber);
                                        actionL3.Add(actL3attrsNumber);
                                    }
                                    else if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count != 0)
                                    {
                                        for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                        {
                                            /*XElement*/ actionL4 = new XElement("actionL4");
                                            XAttribute actL4attrsName = new XAttribute("flName", actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName);
                                            actionL4.Add(actL4attrsName);
                                            XAttribute actL4attrsCommand = new XAttribute("fourthlCommand", actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].fourthlCommand);
                                            actionL4.Add(actL4attrsCommand);
                                            XAttribute actL4attrsIsNum = new XAttribute("isNumerated", actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isNumerated);
                                            actionL4.Add(actL4attrsIsNum);
                                            if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isNumerated == true)
                                            {
                                                XAttribute actL4attrsNumber = new XAttribute("itemNumber", actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber);
                                                actionL4.Add(actL4attrsNumber);
                                            }
                                            XAttribute actL4attrsIsConsEquip = new XAttribute("isConsistEquip", actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isConsistEquip);
                                            actionL4.Add(actL4attrsIsConsEquip);
                                            if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isConsistEquip == true)
                                            {
                                                XAttribute actL4attrsNameEquip = new XAttribute("equipmentName", actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].equipmentName);
                                                actionL4.Add(actL4attrsNameEquip);
                                            }

                                            actionL3.Add(actionL4);
                                        }
                                    }
                                    actionL2.Add(actionL3);
                                }
                            }
                            actionL1.Add(actionL2);
                        }
                    }
                    actions.Add(actionL1);
                }
                
                xdoc.Add(actions);
                xdoc.Save(filename);
            }
        }



        // Открытие окна создания нового энергообъекта
        private void MenuItem_Click_9(object sender, RoutedEventArgs e)
        {
            New_powerobject npo = new New_powerobject();
            npo.Owner = this;
            npo.ShowDialog();
            /*listBox1.ItemsSource = listOfPowerObjects;
            listBox1.Items.Refresh();*/
            listView4.ItemsSource = listOfPowerObjects;
            listView4.Items.Refresh();

            refreshOrgArrs();
        }

        // Удаление энергообъекта из списка
        private void MenuItem_Click_10(object sender, RoutedEventArgs e)
        {
            /*if (listBox1.SelectedIndex >= 0)
            {
                int itm = Convert.ToInt32(listBox1.SelectedIndex);
                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    if (One_listEquipment[i].NamePO == listOfPowerObjects[itm].NamePO)
                    {
                        equipKeywords.Remove(One_listEquipment[i].nameEquip);
                        One_listEquipment.RemoveAt(i);
                        i = i - 1;
                    }
                }
                refreshListView();
                listOfPowerObjects.RemoveAt(itm);
                           
                listBox1.Items.Refresh();
            }*/
            if (listView4.SelectedIndex >= 0)
            {
                int itm = Convert.ToInt32(listView4.SelectedIndex);
                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    if (One_listEquipment[i].NamePO == listOfPowerObjects[itm].NamePO)
                    {
                        equipKeywords.Remove(One_listEquipment[i].nameEquip);
                        One_listEquipment.RemoveAt(i);
                        i = i - 1;
                    }
                }
                refreshListView();
                listOfPowerObjects.RemoveAt(itm);

                listView4.Items.Refresh();
            }
            refreshOrgArrs();
        }

        // Изменение названия у энергообъекта
        private void MenuItem_Click_11(object sender, RoutedEventArgs e)
        {
            
            /*if (listBox1.SelectedIndex >= 0)
            {
                numb = Convert.ToInt32(listBox1.SelectedIndex);
                Change_name_of_power_object new_name_po = new Change_name_of_power_object();
                new_name_po.Owner = this;
                new_name_po.textBox1.Text = listOfPowerObjects[numb].NamePO;
                new_name_po.ShowDialog();
                refreshListView();                
                listBox1.Items.Refresh();
            }*/

            if (listView4.SelectedIndex >= 0)
            {
                numb = Convert.ToInt32(listView4.SelectedIndex);
                Change_name_of_power_object new_name_po = new Change_name_of_power_object();
                new_name_po.Owner = this;
                new_name_po.textBox1.Text = listOfPowerObjects[numb].NamePO;
                new_name_po.textBox2.Text = listOfPowerObjects[numb].organisationPO;
                new_name_po.checkBox1.IsChecked = listOfPowerObjects[numb].isUsed;
                new_name_po.ShowDialog();
                refreshListView();
                listView4.Items.Refresh();
            }

            refreshOrgArrs();

        }



        // Добавление нового оборудования на энергообъект
        private void MenuItem_Click_12(object sender, RoutedEventArgs e)
        {
            /*if (listBox1.SelectedIndex >= 0)
            {
                numb = Convert.ToInt32(listBox1.SelectedIndex);
                New_equipment2 nequip = new New_equipment2();
                nequip.Owner = this;
                nequip.ShowDialog();
                refreshListView();
                listView3.Items.Refresh();                
                

                /*Change_name_of_power_object new_name_po = new Change_name_of_power_object();
                new_name_po.Owner = this;
                new_name_po.textBox1.Text = listOfPowerObjects[numb].NamePO;
                new_name_po.ShowDialog();
                listBox1.Items.Refresh();*/
            /*}
            else*/

            if (listView4.SelectedIndex >= 0)
            {
                numb = Convert.ToInt32(listView4.SelectedIndex);
                New_equipment2 nequip = new New_equipment2();
                nequip.Owner = this;
                nequip.ShowDialog();
                refreshListView();
                listView3.Items.Refresh();


                /*Change_name_of_power_object new_name_po = new Change_name_of_power_object();
                new_name_po.Owner = this;
                new_name_po.textBox1.Text = listOfPowerObjects[numb].NamePO;
                new_name_po.ShowDialog();
                listBox1.Items.Refresh();*/
            }
            else
            {
                MessageBox.Show("Не выбран энергообъект");
            }
        }



        // Событие при выборе энергообъекта
        private void listBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            refreshListView();            
        }

        // Событие при выборе энергообъекта
        private void listView4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            refreshListView();
        }



        // Метод обновляющий список оборудования у выбранного энергообъекта
        public void refreshListView ()
        {
            /*if (listBox1.SelectedIndex >= 0)
            {
                listOfSelectedPowerObject.Clear();
                int itm = Convert.ToInt32(listBox1.SelectedIndex);
                if (One_listEquipment.Count != 0)
                {
                    for (int i = 0; i < One_listEquipment.Count; i++)
                    {
                        if (One_listEquipment[i].NamePO == listOfPowerObjects[itm].NamePO)
                        {
                            PowerObject.Equipment newEquip = new PowerObject.Equipment();
                            newEquip.NamePO = One_listEquipment[i].NamePO;
                            newEquip.nameEquip = One_listEquipment[i].nameEquip;
                            newEquip.stateEquip = One_listEquipment[i].stateEquip;
                            listOfSelectedPowerObject.Add(newEquip);
                        }
                    }
                    listView3.ItemsSource = listOfSelectedPowerObject;
                    listView3.Items.Refresh();
                }
            }*/
            if (listView4.SelectedIndex >= 0)
            {
                listOfSelectedPowerObject.Clear();
                int itm = Convert.ToInt32(listView4.SelectedIndex);
                if (One_listEquipment.Count != 0)
                {
                    for (int i = 0; i < One_listEquipment.Count; i++)
                    {
                        if (One_listEquipment[i].NamePO == listOfPowerObjects[itm].NamePO)
                        {
                            PowerObject.Equipment newEquip = new PowerObject.Equipment();
                            newEquip.NamePO = One_listEquipment[i].NamePO;
                            newEquip.nameEquip = One_listEquipment[i].nameEquip;
                            newEquip.stateEquip = One_listEquipment[i].stateEquip;
                            newEquip.typeEquip = One_listEquipment[i].typeEquip;
                            listOfSelectedPowerObject.Add(newEquip);
                        }
                    }
                    listView3.ItemsSource = listOfSelectedPowerObject;
                    listView3.Items.Refresh();
                }
            }
        }

        // Событие при изменении состояния КА
        private void checkBoxSlct_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < listOfSelectedPowerObject.Count; i++)
            {
                for (int j = 0; j < One_listEquipment.Count; j++)
                {
                    if (listOfSelectedPowerObject[i].nameEquip == One_listEquipment[j].nameEquip &&
                        listOfSelectedPowerObject[i].stateEquip != One_listEquipment[j].stateEquip)
                    {
                        One_listEquipment[j].stateEquip = listOfSelectedPowerObject[i].stateEquip;
                        break;
                    }
                }
            }
        }

        // Изменение параметров оборудования на энрегообъекте
        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            if (listView3.SelectedIndex >= 0)
            {
                numb = Convert.ToInt32(listView3.SelectedIndex);
                Change_equipment_parameters new_equip = new Change_equipment_parameters();                
                new_equip.Owner = this;
                new_equip.textBox1.Text = listOfSelectedPowerObject[numb].nameEquip;
                switch (listOfSelectedPowerObject[numb].typeEquip)
                {
                    case "Switch":
                        new_equip.comboBox1.SelectedIndex = 0;
                        break;
                    case "Disconnector":
                        new_equip.comboBox1.SelectedIndex = 1;
                        break;
                    case "GroundDisconnector":
                        new_equip.comboBox1.SelectedIndex = 2;
                        break;
                }                    
                //new_equip.comboBox1.Text = listOfSelectedPowerObject[numb].typeEquip;
                new_equip.checkBox1.IsChecked = listOfSelectedPowerObject[numb].stateEquip;
                new_equip.ShowDialog();
                int oldIndx = listView4.SelectedIndex;
                listView4.SelectedIndex = -1;
                listView4.SelectedIndex = oldIndx;
                listView3.Items.Refresh();
            }
        }

        // Удаление оборудования
        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            if (listView3.SelectedIndex >= 0)
            {
                int itm = Convert.ToInt32(listView3.SelectedIndex);
                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    if (One_listEquipment[i].nameEquip == listOfSelectedPowerObject[itm].nameEquip)
                    {
                        equipKeywords.Remove(One_listEquipment[i].nameEquip);
                        One_listEquipment.RemoveAt(i);
                        i = i - 1;
                    }
                }
                //refreshListView();
                //listOfPowerObjects.RemoveAt(itm);

                int selItm = listView4.SelectedIndex;
                listView4.SelectedIndex = -1;
                listView4.SelectedIndex = selItm;
            }
        }

        // Тест для создания дерева с командами
        private void button_Click(object sender, RoutedEventArgs e)
        {

            /*ProductClass pcl = new ProductClass();
            CategoryClass cCls = new CategoryClass();
            cCls.levName = "root";
            pcl.prodName = "first prod";           
            cCls.ProductsList.Add(pcl);
            pcl.prodName = "another one";
            cCls.ProductsList.Add(pcl);*/

            //actionsList = new ObservableCollection<CategoryClass>();
            var cat1 = new FirstLevelClass("Начало");                 // создается узел первого уровня    
            cat1.flCommand = "Начало";        
            var cat12 = new SecondLevelClass("cat12");              // создается узел второго уровня
            var cat13 = new SecondLevelClass("РДУ");              // создается узел второго уровня
            var cat123 = new ThirdLevelClass("cat123");             // создается узел третьего уровня

            cat123.FourthLevelList = new ObservableCollection<FourthLevelClass>();  // в узле третьего уровня создается список четвертого уровня
            cat123.FourthLevelList.Add(new FourthLevelClass("cat12_2_3"));          // в узел третьего уровня вносится узел четвертого уровня

            cat12.ThirdLevelList = new ObservableCollection<ThirdLevelClass>();     // в узле второго уровня создается список третьего уровня
            cat12.ThirdLevelList.Add(new ThirdLevelClass("cat12_1"));               // в узел второго уровня вносится узел третьего уровня
            //cat12.ThirdLevelList.Add(new ThirdLevelClass("cat12_2"));               // в узел второго уровня вносится узел третьего уровня
            cat12.ThirdLevelList.Add(cat123);                                       // в узел второго уровня вносится узел третьего уровня

            cat13.ThirdLevelList = new ObservableCollection<ThirdLevelClass>();
            cat13.ThirdLevelList.Add(new ThirdLevelClass("Проверить фиксацию ОТС ВЛ 220 кВ А - Б в положении «Отключено» в ОИК СК-2007, при несоответствии зафиксировать вручную"));
            cat13.ThirdLevelList.Add(new ThirdLevelClass("cat13_2"));

            cat1.SecondLevelList = new ObservableCollection<SecondLevelClass>();    // в узле первого уровня создается список второго уровня
            //cat1.SecondLevelList.Add(new SecondLevelClass("p1"));                   // в узел первого уровня вносится узел второго уровня
            cat1.SecondLevelList.Add(new SecondLevelClass("РДУ"));                   // в узел первого уровня вносится узел второго уровня  
            cat1.SecondLevelList.Add(new SecondLevelClass("ПС 220 кВ Б"));                   // в узел первого уровня вносится узел второго уровня          

            cat1.SecondLevelList.Add(cat12);                        // в узел первого уровня вносится узел второго уровня
            cat1.SecondLevelList.Add(cat13);                        // в узел первого уровня вносится узел второго уровня                     

            actionsList.Add(cat1);                                  // в корневой список вносится узел первого уровня        
            var cat2 = new FirstLevelClass("Операции по пп. ^ выполнять одновременно");  // создается узел первого уровня
            cat2.flCommand = "Операции по пп. ^ выполнять одновременно";
            actionsList.Add(cat2); // в корневой список вносится узел первого уровня    
            actionsList.Add(new FirstLevelClass("cat2"));           // в корневой список вносится узел первого уровня
            actionsList.Add(new FirstLevelClass("cat3"));           // в корневой список вносится узел первого уровня


            treeView2.ItemsSource = actionsList;



        }

        // Тест удаления узлов из дерева команд
        private void button4_Click(object sender, RoutedEventArgs e)
        {
            actionsList[0].SecondLevelList[2].ThirdLevelList[2].FourthLevelList.RemoveAt(0);
        }

        // При загрузке формы сразу же открывается файл с оборудованием
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Переменные для автозагруки нужных файлов Оборудования и Действий
            int openEquip = 0;
            int openAction = 0;

            if (openEquip == 1)
            {
                // Save document
                filename = "D:\\Для работы\\Программа переключений (C#)\\Switching Program (WPF)\\WpfApplication6\\Оборудование.xml";

                XDocument xdoc = XDocument.Load(filename);

                foreach (XElement info in xdoc.Element("Root").Elements("Information"))
                {
                    XAttribute actionDate = info.Attribute("actionDate");
                    XAttribute nameLine = info.Attribute("nameLine");
                    XAttribute isACLineSegment = info.Attribute("isACLineSegment");
                    XAttribute ACLineSegment = info.Attribute("ACLineSegment");
                    XAttribute lineOrganisation = info.Attribute("lineOrganisation");
                    XAttribute typeDo = info.Attribute("typeDO");
                    XAttribute dispOffice = info.Attribute("dispOffice");
                    XAttribute aim = info.Attribute("aim");

                    XAttribute inducedVoltage = info.Attribute("inducedVoltage");
                    XAttribute isUsedARM = info.Attribute("isUsedARM");
                    XAttribute ferroresonance = info.Attribute("ferroresonance");

                    datePicker1.Text = actionDate.Value;
                    textBox1.Text = dispOffice.Value;
                    textBox2.Text = nameLine.Value;
                    textBox3.Text = lineOrganisation.Value;
                    string forAim = aim.Value;
                    switch (forAim)
                    {
                        case "Вывод в ремонт":
                            comboBox1.SelectedIndex = 0;
                            break;
                        case "Вывод в резерв":
                            comboBox1.SelectedIndex = 1;
                            break;
                        case "Ввод из резерва":
                            comboBox1.SelectedIndex = 2;
                            break;
                        case "Ввод в работу":
                            comboBox1.SelectedIndex = 3;
                            break;
                    }
                    string forTypeDO = typeDo.Value;
                    switch (forTypeDO)
                    {
                        case "ИА":
                            comboBox3.SelectedIndex = 0;
                            break;
                        case "ОДУ":
                            comboBox3.SelectedIndex = 1;
                            break;
                        case "РДУ":
                            comboBox3.SelectedIndex = 2;
                            break;
                    }


                    // Запись основной информации в класс с основными параметрами Программы переключений
                    mainParamsOfSP.Aim = aim.Value;
                    mainParamsOfSP.typeDO = typeDo.Value;
                    mainParamsOfSP.dispOffice = dispOffice.Value;
                    mainParamsOfSP.nameLine = nameLine.Value;
                    mainParamsOfSP.isACLineSegment = Convert.ToBoolean(isACLineSegment.Value);
                    mainParamsOfSP.ACLineSegment = ACLineSegment.Value;
                    mainParamsOfSP.lineOrganisation = lineOrganisation.Value;
                    mainParamsOfSP.actionDate = actionDate.Value;

                    mainParamsOfSP.inducedVoltage = Convert.ToBoolean(inducedVoltage.Value);
                    mainParamsOfSP.isUsedARM = Convert.ToBoolean(isUsedARM.Value);
                    mainParamsOfSP.ferroresonance = Convert.ToBoolean(ferroresonance.Value);

                    checkBox4.IsChecked = mainParamsOfSP.isACLineSegment;
                    if (mainParamsOfSP.isACLineSegment == true)
                    {
                        textBox4.Text = mainParamsOfSP.ACLineSegment;
                    }
                    else textBox4.Text = "";

                    checkBox1.IsChecked = mainParamsOfSP.inducedVoltage;
                    checkBox2.IsChecked = mainParamsOfSP.isUsedARM;
                    checkBox3.IsChecked = mainParamsOfSP.ferroresonance;
                }

                foreach (XElement equip in xdoc.Element("Root").Elements("Equipment"))
                {
                    XAttribute name_PO = equip.Attribute("namePO");
                    XAttribute organisation_PO = equip.Attribute("organisationPO");
                    XAttribute is_Used = equip.Attribute("isUsed");
                    XAttribute nameEquip = equip.Attribute("nameEquip");
                    XAttribute stateEquip = equip.Attribute("stateEquip");
                    XAttribute typeEquip = equip.Attribute("typeEquip");

                    PowerObject.Equipment powEquip = new PowerObject.Equipment();
                    powEquip.NamePO = name_PO.Value;
                    powEquip.organisationPO = organisation_PO.Value;
                    powEquip.isUsed = Convert.ToBoolean(is_Used.Value);
                    powEquip.nameEquip = nameEquip.Value;
                    powEquip.stateEquip = Convert.ToBoolean(stateEquip.Value);
                    powEquip.typeEquip = typeEquip.Value;
                    One_listEquipment.Add(powEquip);
                }

                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    int new_object = 1;
                    PowerObject new_PO = new PowerObject();
                    if (listOfPowerObjects.Count == 0)
                    {
                        new_PO.NamePO = One_listEquipment[i].NamePO;
                        new_PO.organisationPO = One_listEquipment[i].organisationPO;
                        new_PO.isUsed = One_listEquipment[i].isUsed;
                        listOfPowerObjects.Add(new_PO);
                    }
                    else
                    {
                        for (int j = 0; j < listOfPowerObjects.Count; j++)
                        {
                            if (listOfPowerObjects[j].NamePO == One_listEquipment[i].NamePO)
                            {
                                new_object = new_object - 1;
                            }
                        }
                        if (new_object != 0)
                        {
                            new_PO.NamePO = One_listEquipment[i].NamePO;
                            new_PO.organisationPO = One_listEquipment[i].organisationPO;
                            new_PO.isUsed = One_listEquipment[i].isUsed;
                            listOfPowerObjects.Add(new_PO);
                            new_object = 1;
                        }
                    }
                }

                listView4.ItemsSource = listOfPowerObjects;
                listView4.Items.Refresh();
                for (int i = 0; i < One_listEquipment.Count; i++)
                {
                    string replSpace = "\u00A0";
                    One_listEquipment[i].nameEquip = One_listEquipment[i].nameEquip.Replace(" ", replSpace);
                    string equipm = One_listEquipment[i].nameEquip;
                    if (equipKeywords.ContainsKey(equipm))
                    {
                        MessageBox.Show("Оборудование с данным названием уже существует. Измените название оборудования");
                    }
                    else
                    {
                        equipKeywords.Add(equipm, FontWeights.Bold);
                    }
                }

                //textBox2.Text = "ВЛ 220 кВ А - Б";    //------------------------------------------
                //comboBox1.SelectedIndex = 1;
                progOption.nameSP = "Программа переключений № 1 по выводу в ремонт ВЛ 220 кВ А – Б";
                progOption.roditPadezh = "Нижегородского РДУ";
                //textBlock6.Text = progOption.nameSP;

                // Заполнение Персонала, участвующего в переключениях
                foreach (XElement persSP in xdoc.Element("Root").Elements("PersonalOfSP"))
                {
                    foreach (XElement organis in persSP.Elements("Organisation"))
                    {
                        var perpers = new Personal("");
                        XAttribute organisation_Of_Personal = organis.Attribute("organisationOfPersonal");

                        perpers.organisationOfPersonal = organisation_Of_Personal.Value;
                        foreach (XElement pers in organis.Elements("Personal"))
                        {
                            var man = new PersonalClass("");

                            XAttribute name_Of_Person = pers.Attribute("nameOfPerson");
                            XAttribute role = pers.Attribute("role");
                            XAttribute action = pers.Attribute("action");

                            man.nameOfPerson = name_Of_Person.Value;
                            man.role = role.Value;
                            man.action = action.Value;

                            perpers.Person.Add(man);
                        }
                        listOfPersonal.Add(perpers);
                    }
                }




                // Заполнение организационных мероприятий
                for (int i = 0; i < listOfPowerObjects.Count; i++)
                {
                    if (listOfPowerObjects[i].isUsed == true)
                    {
                        var oars = new Org_arrangs();
                        oars.NameObj = "На " + listOfPowerObjects[i].NamePO;
                        oars.PObject = listOfPowerObjects[i].NamePO;
                        oars.isWork = false;

                        listOfOrgArrs.Add(oars);
                    }
                }
                var oar = new Org_arrangs();
                oar.NameObj = "На " + textBox2.Text;
                oar.PObject = textBox2.Text;
                oar.isWork = false;
                listOfOrgArrs.Add(oar);

                oar = new Org_arrangs();
                oar.NameObj = "Не будут производиться";
                oar.PObject = "Не будут производиться";
                oar.isWork = false;
                listOfOrgArrs.Add(oar);

                for (int i = 0; i < listOfOrgArrs.Count; i++)
                {
                    var uuu = new Org_arrangs();
                    uuu.NameObj = listOfOrgArrs[i].NameObj;
                    uuu.PObject = listOfOrgArrs[i].PObject;
                    uuu.isWork = listOfOrgArrs[i].isWork;
                    listOfOldOrgArrs.Add(uuu);
                }
                //listOfOldOrgArrs = listOfOrgArrs;
                listView5.ItemsSource = listOfOrgArrs;
            }

            // Открытие действий
            if (openAction == 1)
            { 
                string filenameA = "D:\\Для работы\\Программа переключений (C#)\\Switching Program (WPF)\\WpfApplication5\\Действия.xml";
            //-------------------------
            //openDialog.Filter = "XML files(*.xml)|*.xml|All files(*.*)|*.*";

            //Nullable<bool> open_result = openDialog.ShowDialog();

            //if (open_result == true)
            //{
            // Save document
            //filename = openDialog.FileName;

            //var lev1 = new FirstLevelClass("");
            //-----------------------------------\\



            XDocument xdocA = XDocument.Load(filenameA);
            foreach (XElement L1element in xdocA.Element("actions").Elements("actionL1"))
            {
                var lev1 = new FirstLevelClass("");
                XAttribute nameAttr1 = L1element.Attribute("flName");
                lev1.FlName = nameAttr1.Value;
                XAttribute commandAttr1 = L1element.Attribute("flCommand");
                lev1.flCommand = commandAttr1.Value;
                //XElement priceElement = phoneElement.Element("price");
                foreach (XElement L2element in L1element.Elements("actionL2"))
                {
                    var lev2 = new SecondLevelClass("");
                    XAttribute nameAttr2 = L2element.Attribute("slName");
                    lev2.SlName = nameAttr2.Value;
                    XAttribute commandAttr2 = L2element.Attribute("slCommand");
                    lev2.slCommand = commandAttr2.Value;

                    foreach (XElement L3element in L2element.Elements("actionL3"))
                    {
                        var lev3 = new ThirdLevelClass("");
                        XAttribute nameAttr3 = L3element.Attribute("tlName");
                        lev3.TlName = nameAttr3.Value;
                        XAttribute commandAttr3 = L3element.Attribute("tlCommand");
                        lev3.tlCommand = commandAttr3.Value;
                        XAttribute isNumAttr3 = L3element.Attribute("isNumerated");
                        lev3.isNumerated = Convert.ToBoolean(isNumAttr3.Value);
                        if (isNumAttr3.Value == "true")
                        {
                            XAttribute itemNumAttr3 = L3element.Attribute("itemNumber");
                            lev3.itemNumber = itemNumAttr3.Value;
                        }
                        XAttribute isConsEquipAttr3 = L3element.Attribute("isConsistEquip");
                        lev3.isConsistEquip = Convert.ToBoolean(isConsEquipAttr3.Value);
                        if (isConsEquipAttr3.Value == "true")
                        {
                            XAttribute EquipNameAttr3 = L3element.Attribute("equipmentName");
                            lev3.equipmentName = EquipNameAttr3.Value;
                        }
                        if (isNumAttr3.Value == "false" && isConsEquipAttr3.Value == "false")
                        {
                            foreach (XElement L4element in L3element.Elements("actionL4"))
                            {
                                var lev4 = new FourthLevelClass("");
                                XAttribute nameAttr4 = L4element.Attribute("flName");
                                lev4.FlName = nameAttr4.Value;
                                XAttribute commandAttr4 = L4element.Attribute("fourthlCommand");
                                lev4.fourthlCommand = commandAttr4.Value;
                                XAttribute isNumAttr4 = L4element.Attribute("isNumerated");
                                lev4.isNumerated = Convert.ToBoolean(isNumAttr4.Value);
                                if (isNumAttr4.Value == "true")
                                {
                                    XAttribute itemNumAttr4 = L4element.Attribute("itemNumber");
                                    lev4.itemNumber = itemNumAttr4.Value;
                                }
                                XAttribute isConsEquipAttr4 = L4element.Attribute("isConsistEquip");
                                lev4.isConsistEquip = Convert.ToBoolean(isConsEquipAttr4.Value);
                                if (isConsEquipAttr4.Value == "true")
                                {
                                    XAttribute EquipNameAttr4 = L4element.Attribute("equipmentName");
                                    lev4.equipmentName = EquipNameAttr4.Value;
                                }
                                lev3.FourthLevelList.Add(lev4);
                            }
                        }
                        lev2.ThirdLevelList.Add(lev3);
                    }
                    lev1.SecondLevelList.Add(lev2);
                }

                //----------------------------------------

                //if (nameAttribute != null && companyElement != null && priceElement != null)
                //{
                //Console.WriteLine("Смартфон: {0}", nameAttribute.Value);
                //Console.WriteLine("Компания: {0}", companyElement.Value);
                //Console.WriteLine("Цена: {0}", priceElement.Value);
                //}
                //Console.WriteLine();
                //-------------------------------------------\\
                actionsList.Add(lev1);
            }
        }
            //} // это так и оставить закомментированным

            treeView2.ItemsSource = actionsList;
        }

        // Метод, позволяющий обновлять список Организационных мероприятий
        public void refreshOrgArrs ()
        {
            listOfOrgArrs.Clear();
            for (int i = 0; i < listOfPowerObjects.Count; i++)
            {
                if (listOfPowerObjects[i].isUsed == true)
                {
                    var oars = new Org_arrangs();
                    oars.NameObj = "На " + listOfPowerObjects[i].NamePO;
                    oars.PObject = listOfPowerObjects[i].NamePO;
                    oars.isWork = false;

                    listOfOrgArrs.Add(oars);
                }
            }
            var oar = new Org_arrangs();
            oar.NameObj = "На " + textBox2.Text;
            oar.PObject = textBox2.Text;
            oar.isWork = false;
            listOfOrgArrs.Add(oar);

            oar = new Org_arrangs();
            oar.NameObj = "Не будут производиться";
            oar.PObject = "Не будут производиться";
            oar.isWork = false;
            listOfOrgArrs.Add(oar);

            for (int i = 0; i < listOfOldOrgArrs.Count; i++)
            {
                for (int j = 0; j < listOfOrgArrs.Count; j++)
                {
                    if (listOfOrgArrs[j].PObject == listOfOldOrgArrs[i].PObject)
                    {
                        listOfOrgArrs[j].isWork = listOfOldOrgArrs[i].isWork;
                    }
                }
            }

            for (int i = 0; i < listOfOrgArrs.Count; i++)
            {
                var uuu = new Org_arrangs();
                uuu.NameObj = listOfOrgArrs[i].NameObj;
                uuu.PObject = listOfOrgArrs[i].PObject;
                uuu.isWork = listOfOrgArrs[i].isWork;
                listOfOldOrgArrs.Add(uuu);
            }
            /*listOfOldOrgArrs.Clear();
            listOfOldOrgArrs = listOfOrgArrs;*/
            listView5.ItemsSource = listOfOrgArrs;
            listView5.Items.Refresh();          
        }

        // Событие при нажатии на правую кнопку мыши в дереве
        private void treeView2_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Right)
            {
                var poc = treeView2.SelectedItem;            
                
                if (treeView2.SelectedItem != null && treeView2.SelectedItem is FourthLevelClass != true) // Если выбранный узел не четвертого уровня
                {
                    searchPositionOfSelectedItem();
                    contextMenu.Items.Clear();
                    //int innZ = 19;
                    // Create menu item.
                    MenuItem addCurrentLevelMenuItem = new MenuItem();
                    addCurrentLevelMenuItem.Header = "Добавить узел текущего уровня после выбранного";
                    addCurrentLevelMenuItem.Click += new RoutedEventHandler(currentLevelMenuItem_Click);
                    contextMenu.Items.Add(addCurrentLevelMenuItem);

                    if (treeView2.SelectedItem is ThirdLevelClass == true)
                    {
                        searchPositionOfSelectedItem();
                        if (actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].SlName != comboBox3.Text &&
                            actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].ThirdLevelList[positionOfSelectedItem[2]].isNumerated != true)
                        {
                            MenuItem addChildLevelMenuItem = new MenuItem();
                            addChildLevelMenuItem.Header = "Добавить дочерний узел";
                            addChildLevelMenuItem.Click += new RoutedEventHandler(childLevelMenuItem_Click);
                            //addChildLevelMenuItem.Click += new RoutedEventHandler(exitMenuItem_Click);
                            // Create contextual menu.
                            //contextMenu = new ContextMenu();
                            contextMenu.Items.Add(addChildLevelMenuItem);
                        }
                    }
                    else
                    {
                        MenuItem addChildLevelMenuItem = new MenuItem();
                        addChildLevelMenuItem.Header = "Добавить дочерний узел";
                        addChildLevelMenuItem.Click += new RoutedEventHandler(childLevelMenuItem_Click);
                        //addChildLevelMenuItem.Click += new RoutedEventHandler(exitMenuItem_Click);
                        // Create contextual menu.
                        //contextMenu = new ContextMenu();
                        contextMenu.Items.Add(addChildLevelMenuItem);
                    }
                    

                    MenuItem chahgeMenuItem = new MenuItem();
                    chahgeMenuItem.Header = "Изменить узел";
                    chahgeMenuItem.Click += new RoutedEventHandler(chahgeMenuItem_Click);
                    contextMenu.Items.Add(chahgeMenuItem);

                    MenuItem deleteMenuItem = new MenuItem();
                    deleteMenuItem.Header = "Удалить узел";
                    deleteMenuItem.Click += new RoutedEventHandler(deleteMenuItem_Click);
                    contextMenu.Items.Add(deleteMenuItem);
                }
                else if (treeView2.SelectedItem is FourthLevelClass) // Если выбранный узел четвертого уровня
                {
                    contextMenu.Items.Clear();
                    MenuItem addCurrentLevelMenuItem = new MenuItem();
                    addCurrentLevelMenuItem.Header = "Добавить узел текущего уровня после выбранного";
                    addCurrentLevelMenuItem.Click += new RoutedEventHandler(currentLevelMenuItem_Click);
                    contextMenu.Items.Add(addCurrentLevelMenuItem);                    

                    MenuItem chahgeMenuItem = new MenuItem();
                    chahgeMenuItem.Header = "Изменить узел";
                    chahgeMenuItem.Click += new RoutedEventHandler(chahgeMenuItem_Click);
                    contextMenu.Items.Add(chahgeMenuItem);

                    MenuItem deleteMenuItem = new MenuItem();
                    deleteMenuItem.Header = "Удалить узел";
                    deleteMenuItem.Click += new RoutedEventHandler(deleteMenuItem_Click);
                    contextMenu.Items.Add(deleteMenuItem);
                }
                else // Если узел не выбран
                {
                    contextMenu.Items.Clear();
                    MenuItem createFirstNodeMenuItem = new MenuItem();
                    createFirstNodeMenuItem.Header = "Добавить узел первого уровня";
                    createFirstNodeMenuItem.Click += new RoutedEventHandler(currentLevelMenuItem_Click);
                    contextMenu.Items.Add(createFirstNodeMenuItem);
                }
                
                // Attach context-menu to something.
                treeView2.ContextMenu = contextMenu; // Assuming there a button on window named "myButton".
            }
        }

        // Добавление узла текущего уровня
        public void currentLevelMenuItem_Click(object sender, RoutedEventArgs e)
        {
            searchPositionOfSelectedItem();
            // This gets executed when user right-clicks button, and presses x down on their keyboard.
            // TODO: Exit logic.
            if (actionsList.Count == 0)
            {                
                var sat = new FirstLevelClass("Начало");
                sat.flCommand = "Начало";
                actionsList.Add(sat);
                treeView2.ItemsSource = actionsList;
            }
            else
            {
                AddAction adac = new AddAction();
                adac.Owner = this;
                if (actionsList.Count != 0)
                {
                    if (treeView2.SelectedItem is FirstLevelClass || treeView2.SelectedItem == null)
                    {
                        adac.textBlock1.Text = "Первый уровень";
                        adac.comboBox1.IsEnabled = true;                        
                        adac.comboBox2.IsEnabled = false;                        
                        adac.checkBox1.IsEnabled = false;
                        adac.comboBox3.IsEnabled = false;                     
                        adac.richTextBox1.IsEnabled = true;
                        
                        adac.comboBox1.Items.Add("Начало");
                        adac.comboBox1.Items.Add("Операции по пп. ^ выполнять одновременно");
                    }
                    if (treeView2.SelectedItem is SecondLevelClass)
                    {
                        adac.textBlock1.Text = "Первый уровень";
                        adac.comboBox1.IsEnabled = false;                        
                        adac.comboBox2.IsEnabled = true;                        
                        adac.checkBox1.IsEnabled = false;
                        adac.comboBox3.IsEnabled = false;
                        adac.richTextBox1.IsEnabled = true;

                        adac.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].FlName);
                        adac.comboBox1.SelectedIndex = 0;
                        adac.comboBox2.Items.Add(comboBox3.Text);
                                             
                        for (int i = 0; i < listOfPowerObjects.Count; i++)
                        {
                            adac.comboBox2.Items.Add(listOfPowerObjects[i].NamePO);
                        }
                    }
                    if (treeView2.SelectedItem is ThirdLevelClass)
                    {
                        adac.textBlock1.Text = "Второй уровень";
                        adac.comboBox1.IsEnabled = false;
                        adac.comboBox2.IsEnabled = true;
                        adac.checkBox1.IsEnabled = false;
                        adac.comboBox3.IsEnabled = false;
                        adac.richTextBox1.IsEnabled = true;

                        adac.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].SlName);
                        adac.comboBox1.SelectedIndex = 0;
                        /*adac.comboBox2.Items.Add(comboBox3.Text);
                        for (int i = 0; i < listOfPowerObjects.Count; i++)
                        {
                            adac.comboBox2.Items.Add(listOfPowerObjects[i].NamePO);
                        }*/
                        
                    }
                    if (treeView2.SelectedItem is FourthLevelClass)
                    {
                        adac.textBlock1.Text = "Третий уровень";
                        adac.comboBox1.IsEnabled = false;
                        adac.comboBox2.IsEnabled = true;
                        adac.checkBox1.IsEnabled = true;
                        adac.checkBox1.IsChecked = true;
                        adac.comboBox3.IsEnabled = true;
                        adac.richTextBox1.IsEnabled = true;

                        adac.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].
                            ThirdLevelList[positionOfSelectedItem[2]].TlName);
                        adac.comboBox1.SelectedIndex = 0;
                        /*adac.comboBox2.Items.Add(comboBox3.Text);
                        for (int i = 0; i < listOfPowerObjects.Count; i++)
                        {
                            adac.comboBox2.Items.Add(listOfPowerObjects[i].NamePO);
                        }*/
                        
                    }
                }
                adac.ShowDialog();
            }
            /*var numNode = (treeView2.SelectedItem as SecondLevelClass);
            if (treeView2.SelectedItem is SecondLevelClass)
            {
                int innZ = 19;
            }*/


            /*if ((treeView2.SelectedItem as SecondLevelClass).Level == 2)
            {
                int iin = 42;
            }*/
            numerationOfCommands();
            treeView2.ItemsSource = nodes;      // Хрень, которая помогает при обновлении нумерации в дереве
            treeView2.ItemsSource = actionsList;
        }

        // Добавление дочернего узла
        public void childLevelMenuItem_Click(object sender, RoutedEventArgs e)
        {
            searchPositionOfSelectedItem();
            AddChildAction aca = new AddChildAction();
            aca.Owner = this;
            aca.comboBox1.IsEnabled = false;
            aca.comboBox2.IsEnabled = true;
            //aca.checkBox1.IsEnabled = false;
            aca.richTextBox1.IsEnabled = true;
            switch (positionOfSelectedItem.Count)
            {
                case 1:
                    aca.textBlock1.Text = "Первый уровень";
                    aca.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].FlName);
                    aca.comboBox1.SelectedIndex = 0;                    
                    //aca.comboBox1.IsEnabled = false;
                    aca.comboBox3.IsEnabled = false;
                    aca.checkBox1.IsEnabled = false;
                    //aca.richTextBox1.IsEnabled = true;

                    aca.comboBox2.Items.Add(comboBox3.Text);
                    for (int i = 0; i < listOfPowerObjects.Count; i++)
                    {
                        aca.comboBox2.Items.Add(listOfPowerObjects[i].NamePO);
                    }

                    aca.ShowDialog();
                    break;
                case 2:
                    aca.textBlock1.Text = "Второй уровень";
                    aca.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].SlName);
                    aca.comboBox1.SelectedIndex = 0;
                    aca.checkBox1.IsEnabled = true;
                    aca.comboBox3.IsEnabled = true;

                    aca.ShowDialog();

                    numerationOfCommands();
                    treeView2.ItemsSource = nodes;      // Хрень, которая помогает при обновлении нумерации в дереве
                    //treeView2.ItemsSource = actionsList;
                    break;
                case 3:
                    aca.textBlock1.Text = "Третий уровень";
                    aca.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].ThirdLevelList[positionOfSelectedItem[2]].TlName);
                    aca.comboBox1.SelectedIndex = 0;
                    aca.checkBox1.IsEnabled = true;
                    aca.comboBox3.IsEnabled = true;

                    aca.ShowDialog();

                    numerationOfCommands();
                    treeView2.ItemsSource = nodes;      // Хрень, которая помогает при обновлении нумерации в дереве
                    break;
            }
            treeView2.ItemsSource = actionsList;
        }

        // Изменение параметров выбранного узла
        public void chahgeMenuItem_Click(object sender, RoutedEventArgs e)
        {
            searchPositionOfSelectedItem();
            ChangeAction cac = new ChangeAction();
            cac.Owner = this;
            switch (positionOfSelectedItem.Count)
            {
                case 1:
                    cac.textBlock1.Text = "Первый уровень";
                    cac.comboBox1.IsEnabled = true;
                    cac.comboBox2.IsEnabled = false;
                    cac.checkBox1.IsEnabled = false;
                    cac.comboBox3.IsEnabled = false;

                    cac.comboBox1.Items.Add("Начало");
                    cac.comboBox1.Items.Add("Операции по пп. ^ выполнять одновременно");
                    for (int i = 0; i < cac.comboBox1.Items.Count; i++)
                    {
                        if (cac.comboBox1.Items[i].ToString() == actionsList[positionOfSelectedItem[0]].flCommand)
                        {
                            cac.comboBox1.SelectedIndex = i;
                        }
                    }

                    cac.richTextBox1.Document.Blocks.Clear();
                    cac.richTextBox1.AppendText(actionsList[positionOfSelectedItem[0]].FlName);

                    cac.ShowDialog();
                    numerationOfCommands();
                    treeView2.ItemsSource = nodes;
                    break;
                case 2:
                    cac.textBlock1.Text = "Первый уровень";
                    cac.comboBox1.IsEnabled = false;                    
                    cac.comboBox2.IsEnabled = true;
                    cac.checkBox1.IsEnabled = false;
                    cac.comboBox3.IsEnabled = false;

                    cac.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].FlName);
                    cac.comboBox1.SelectedIndex = 0;

                    cac.comboBox2.Items.Add(comboBox3.Text);
                    for (int i = 0; i < listOfPowerObjects.Count; i++)
                    {
                        cac.comboBox2.Items.Add(listOfPowerObjects[i].NamePO);
                    }

                    for (int i = 0; i < cac.comboBox2.Items.Count; i++)
                    {
                        if (cac.comboBox2.Items[i].ToString() == actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].slCommand)
                        {
                            cac.comboBox2.SelectedIndex = i;
                        }
                    }

                    cac.richTextBox1.Document.Blocks.Clear();
                    cac.richTextBox1.AppendText(actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].SlName);

                    cac.ShowDialog();
                    numerationOfCommands();
                    treeView2.ItemsSource = nodes;
                    break;
                case 3:
                    cac.textBlock1.Text = "Второй уровень";
                    cac.comboBox1.IsEnabled = false;
                    cac.comboBox2.IsEnabled = true;
                    cac.checkBox1.IsEnabled = false;
                    cac.comboBox3.IsEnabled = false;

                    cac.comboBox1.Items.Add(actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].SlName);
                    cac.comboBox1.SelectedIndex = 0;

                    cac.ShowDialog();
                    numerationOfCommands();
                    treeView2.ItemsSource = nodes;
                    break;
                case 4:
                    cac.textBlock1.Text = "Третий уровень";
                    cac.comboBox1.IsEnabled = false;
                    cac.comboBox2.IsEnabled = true;
                    cac.checkBox1.IsEnabled = false;
                    cac.comboBox3.IsEnabled = false;
                    cac.ShowDialog();
                    numerationOfCommands();
                    treeView2.ItemsSource = nodes;
                    break;
            }
            treeView2.ItemsSource = actionsList;
        }

        // Удаление выбранного узла
        public void deleteMenuItem_Click(object sender, RoutedEventArgs e)
        {
            searchPositionOfSelectedItem();
            switch (positionOfSelectedItem.Count)
            {
                case (1):
                    actionsList.RemoveAt(positionOfSelectedItem[0]);
                    break;
                case (2):
                    actionsList[positionOfSelectedItem[0]].SecondLevelList.RemoveAt(positionOfSelectedItem[1]);
                    break;
                case (3):
                    actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].
                                ThirdLevelList.RemoveAt(positionOfSelectedItem[2]);
                    break;
                case (4):
                    actionsList[positionOfSelectedItem[0]].SecondLevelList[positionOfSelectedItem[1]].
                                ThirdLevelList[positionOfSelectedItem[2]].FourthLevelList.RemoveAt(positionOfSelectedItem[3]);
                    break;
            }
            numerationOfCommands();

            treeView2.ItemsSource = nodes;      // Хрень, которая помогает при обновлении нумерации в дереве
            treeView2.ItemsSource = actionsList;
        }


        // Тестовая кнопка ПОИСК КООРДИНАТ ВЫБРАННОГО В ДЕРЕВЕ УЗЛА
        private void button2_Click(object sender, RoutedEventArgs e)
        {
            int flpoz = -1;
            int slpoz = -1;
            int tlpoz = -1;
            int fourtlpoz = -1;
            

            for (int i = 0; i < actionsList.Count; i++)
            {
                if (actionsList[i] as FirstLevelClass == treeView2.SelectedItem)
                {
                    flpoz = i;
                    textBox1.Text = "a: " + flpoz;
                }
                else if (actionsList[i].SecondLevelList != null)
                {
                    for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                    {
                        if (actionsList[i].SecondLevelList[j] as SecondLevelClass == treeView2.SelectedItem)
                        {
                            flpoz = i;
                            slpoz = j;
                            textBox1.Text = "a: " + flpoz + "  b: " + slpoz;
                        }
                        else if (actionsList[i].SecondLevelList[j].ThirdLevelList != null)
                        {
                            for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                            {
                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k] as ThirdLevelClass == treeView2.SelectedItem)
                                {
                                    flpoz = i;
                                    slpoz = j;
                                    tlpoz = k;
                                    textBox1.Text = "a: " + flpoz + "  b: " + slpoz + "  c: " + tlpoz;
                                }
                                else if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList != null)
                                {
                                    for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                    {
                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l] as FourthLevelClass == treeView2.SelectedItem)
                                        {
                                            flpoz = i;
                                            slpoz = j;
                                            tlpoz = k;
                                            fourtlpoz = l;
                                            textBox1.Text = "a: " + flpoz + "  b: " + slpoz + "  c: " + tlpoz + "  d: " + fourtlpoz;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // Собираем все координаты вместе
            positionOfSelectedItem.Clear();
            if (flpoz > -1)
            {
                positionOfSelectedItem.Add(flpoz);
                if (slpoz > -1)
                {
                    positionOfSelectedItem.Add(slpoz);
                    if (tlpoz > -1)
                    {
                        positionOfSelectedItem.Add(tlpoz);
                        if (fourtlpoz > -1)
                        {
                            positionOfSelectedItem.Add(fourtlpoz);
                        }
                    } 
                }
            }
        }

        // Метод, вычисляющий позицию выбранного в дереве узла
        public void searchPositionOfSelectedItem()
        {
            int flpoz = -1;
            int slpoz = -1;
            int tlpoz = -1;
            int fourtlpoz = -1;


            for (int i = 0; i < actionsList.Count; i++)
            {
                if (actionsList[i] as FirstLevelClass == treeView2.SelectedItem)
                {
                    flpoz = i;
                    /*textBox1.Text = "a: " + flpoz;*/
                }
                else if (actionsList[i].SecondLevelList != null)
                {
                    for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                    {
                        if (actionsList[i].SecondLevelList[j] as SecondLevelClass == treeView2.SelectedItem)
                        {
                            flpoz = i;
                            slpoz = j;
                            /*textBox1.Text = "a: " + flpoz + "  b: " + slpoz;*/
                        }
                        else if (actionsList[i].SecondLevelList[j].ThirdLevelList != null)
                        {
                            for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                            {
                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k] as ThirdLevelClass == treeView2.SelectedItem)
                                {
                                    flpoz = i;
                                    slpoz = j;
                                    tlpoz = k;
                                    /*textBox1.Text = "a: " + flpoz + "  b: " + slpoz + "  c: " + tlpoz;*/
                                }
                                else if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList != null)
                                {
                                    for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                    {
                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l] as FourthLevelClass == treeView2.SelectedItem)
                                        {
                                            flpoz = i;
                                            slpoz = j;
                                            tlpoz = k;
                                            fourtlpoz = l;
                                            /*textBox1.Text = "a: " + flpoz + "  b: " + slpoz + "  c: " + tlpoz + "  d: " + fourtlpoz;*/
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // Собираем все координаты вместе
            positionOfSelectedItem.Clear();
            if (flpoz > -1)
            {
                positionOfSelectedItem.Add(flpoz);
                if (slpoz > -1)
                {
                    positionOfSelectedItem.Add(slpoz);
                    if (tlpoz > -1)
                    {
                        positionOfSelectedItem.Add(tlpoz);
                        if (fourtlpoz > -1)
                        {
                            positionOfSelectedItem.Add(fourtlpoz);
                        }
                    }
                }
            }
        }

        // Метод, нумерующий команды в дереве
        public void numerationOfCommands()
        {
            iNumber = 1;
            for (int i = 0; i < actionsList.Count; i++)
            {
                if (actionsList[i].SecondLevelList != null)
                {
                    for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                    {
                        if (actionsList[i].SecondLevelList[j].ThirdLevelList != null)
                        {
                            for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                            {
                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isNumerated == true)
                                {
                                    string numnum = "5." + iNumber + " ";
                                    string oldNum = actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber;
                                    string newName = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName.Replace(oldNum, numnum);
                                    actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName = newName;
                                    //        actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName.Replace(oldNum, numnum);
                                    actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber = numnum;
                                    iNumber++;
                                }
                                else if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList != null)
                                {
                                    for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                    {
                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isNumerated == true)
                                        {
                                            string numnum = "5." + iNumber + " ";
                                            string oldNum = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber;
                                            string newName = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName.Replace(oldNum, numnum);
                                            actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName = newName;
                                            //   actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName.Replace(oldNum, numnum);
                                            actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber = numnum;
                                            iNumber++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            NOC2();
        }

        // Метод, вычисляющий номера одновременных действий
        public void NOC2()
        {   
            for (int i = 0; i < actionsList.Count; i++)
            {   
                List<string> listOfNumbers = new List<string>();             
                if (actionsList[i].SecondLevelList.Count == 0)
                {
                    if (i == 0)
                    {
                        actionsList[i].flCommand = "Начало";
                        actionsList[i].FlName = actionsList[i].flCommand;
                    }
                    else
                    {
                        actionsList[i].flCommand = "Операции по пп. ^ выполнять одновременно";
                        actionsList[i].FlName = actionsList[i].flCommand;
                    }
                }
                else
                {
                    for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                    {
                        if (actionsList[i].SecondLevelList[j].ThirdLevelList != null)
                        {
                            for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                            {
                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isNumerated == true)
                                {
                                    string numberCom = actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber;
                                    numberCom = numberCom.Remove(numberCom.Count() - 1);
                                    listOfNumbers.Add(numberCom);
                                    //iNumber++;
                                }
                                else if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList != null)
                                {
                                    for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                    {
                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isNumerated == true)
                                        {
                                            string nC = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber;
                                            nC = nC.Remove(nC.Count() - 1);
                                            listOfNumbers.Add(nC);
                                            //iNumber++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (actionsList[i].flCommand != "Начало" && listOfNumbers.Count != 0)
                    {
                        string command = listOfNumbers[0] + "-" + listOfNumbers[listOfNumbers.Count - 1] + " ";
                        actionsList[i].flCommand = "Операции по пп. ^ выполнять одновременно";
                        actionsList[i].FlName = actionsList[i].flCommand;
                        actionsList[i].FlName = actionsList[i].FlName.Replace("^", command);
                    }
                }
            }
        }

        // Формирование файла программы переключений
        private void button1_Click(object sender, RoutedEventArgs e)
        {
            //Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            /*var application = new Microsoft.Office.Interop.Word.Application();
            var document = new Microsoft.Office.Interop.Word.Document();
            document = application.Documents.Add();
            application.Visible = true;*/
            try
            {
                Microsoft.Office.Interop.Word.Application winword =
                    new Microsoft.Office.Interop.Word.Application();

                winword.Visible = false;

                //Заголовок документа
                winword.Documents.Application.Caption = "www.CSharpCoderR.com";

                object missing = System.Reflection.Missing.Value;

                //Создание нового документа
                Microsoft.Office.Interop.Word.Document document =
                    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                //добавление новой страницы
                //winword.Selection.InsertNewPage();

                //Добавление верхнего колонтитула
                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[
                    Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(
                   headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment =
                   Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex =
                   Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    headerRange.Text = "Верхний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
                }

                //Добавление нижнего колонтитула
                foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                {
                    //
                    Microsoft.Office.Interop.Word.Range footerRange =
            wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    //Установка цвета текста
                    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                    //Размер
                    footerRange.Font.Size = 10;
                    //Установка расположения по центру
                    footerRange.ParagraphFormat.Alignment =
                        Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Установка текста для вывода в нижнем колонтитуле
                    footerRange.Text = "Нижний колонтитул" + Environment.NewLine + "www.CSharpCoderR.com";
                }

                //Добавление текста в документ
                document.Content.SetRange(0, 0);
                document.Content.Text = "www.CSharpCoderR.com" + Environment.NewLine;

                //Добавление текста со стилем Заголовок 1
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                object styleHeading1 = "Заголовок 1";
                para1.Range.set_Style(styleHeading1);
                para1.Range.Text = "Исходники по языку программирования CSharp";
                para1.Range.InsertParagraphAfter();

                //Создание таблицы 5х5
               Microsoft.Office.Interop.Word.Table firstTable = document.Tables.Add(para1.Range, 5, 5, ref missing, ref missing);

                
                firstTable.Borders.Enable = 1;
                firstTable.Columns[1].PreferredWidth = 50;
                foreach (Microsoft.Office.Interop.Word.Row row in firstTable.Rows)
                {
                    foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
                    {
                        //Заголовок таблицы
                        if (cell.RowIndex == 1)
                        {
                            cell.Range.Text = "Колонка " + cell.ColumnIndex.ToString();
                            cell.Range.Font.Bold = 1;
                            //Задаем шрифт и размер текста
                            cell.Range.Font.Name = "verdana";
                            cell.Range.Font.Size = 10;
                            cell.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray25;
                            //Выравнивание текста в заголовках столбцов по центру
                            cell.VerticalAlignment =
                            Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment =
                            Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        //Значения ячеек
                        else
                        {
                            cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                        }
                    }
                }
                int rowNum = 3;
                Microsoft.Office.Interop.Word.Row row1 = firstTable.Rows[rowNum];
                Microsoft.Office.Interop.Word.Cell firstCell = row1.Cells[2];
                foreach (Microsoft.Office.Interop.Word.Cell currCell in row1.Cells)
                {
                    if (currCell.ColumnIndex != firstCell.ColumnIndex && currCell.ColumnIndex != 1) // объединение ячеек только по правую сторону от выбраной
                    {
                        firstCell.Merge(currCell);
                    }
                }

                //firstTable.Rows[3].Cells[2].Merge(firstTable.Rows[4].Cells[2]);
                

                Microsoft.Office.Interop.Word.Row row2 = firstTable.Rows.Add();
                document.Range(row2.Cells[1].Range.Start, row2.Cells[2].Range.End).Cells.Merge();//Объединение ячеек в третьей снизу строке
                row2.Cells[1].Range.Text = "test " /*+ i.ToString()*/;//Запись текста в первую, уже объединённую ячейку

                row2 = firstTable.Rows.Add();
                //document.Range(row2.Cells[1].Range.Start, row2.Cells[2].Range.End).Cells.Merge();//Объединение ячеек в третьей снизу строке
                row2.Cells[1].Range.Text = "tes2 " /*+ i.ToString()*/;//Запись текста в первую, уже объединённую ячейку

                //para1.Range.Paragraphs.LeftIndent = -30;
                //              document.Paragraphs[4].Range.ParagraphFormat.LeftIndent = 20;











                // Таблица
                para1.Format.LeftIndent = -30;    // Сдвиг всей таблицы влево
                
                
                para1.Range.InsertParagraphAfter();
                //para1.Range.ParagraphFormat.LeftIndent = -40;

                Microsoft.Office.Interop.Word.Table section5 = document.Tables.Add(para1.Range, 1, 5, 2, ref missing);  // Создание таблицы
                //section5.LeftPadding = 40;
                //section5.Range.ParagraphFormat.LeftIndent = -20;
                //Microsoft.Office.Interop.Word.Range rng = section5.Range;
                //rng = Microsoft.Office.Interop.Word.ParagraphFormat
                 // Выделяем следующие 3 абзаца
                /* Object unit = Microsoft.Office.Interop.Word.WdUnits.wdParagraph;
                Object count = 4;
                Object extend = Microsoft.Office.Interop.Word.WdMovementType.wdMove;
                winword.Selection.MoveDown(ref unit, ref count, ref extend);
                winword.Selection.ParagraphFormat.LeftIndent = -20;*/

                section5.Range.Font.Size = 13;                          // Задаётся формат шрифта в таблице
                section5.Range.Font.Name = "Times New Roman";

                /*para1.Range.Font.Size = 13;
                para1.Range.Font.Name = "Times New Roman";*/


                // Хорошие размеры таблицы
                section5.Columns[1].SetWidth(63, WdRulerStyle.wdAdjustNone);
                section5.Columns[2].SetWidth(42, WdRulerStyle.wdAdjustNone);
                section5.Columns[3].SetWidth(340, WdRulerStyle.wdAdjustNone);
                section5.Columns[4].SetWidth(43, WdRulerStyle.wdAdjustNone);
                section5.Columns[5].SetWidth(43, WdRulerStyle.wdAdjustNone);

                section5.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);

                //section5.Rows.Add();
                // Шапка таблицы для пунктов 4 - 7
                section5.Rows[1].Cells[1].Range.ParagraphFormat.LeftIndent = -5;        // Смещение текста в ячейке влево
                section5.Rows[1].Cells[1].Range.Text = "Персонал,\nвыполняющий операцию";
                section5.Rows[1].Cells[2].Range.Text = "п.п.";
                section5.Rows[1].Cells[3].Range.Text = "Объект переключений,\nоперация, сообщение";                
                section5.Rows[1].Cells[4].Range.Text = "Время\nотдачи команды";
                section5.Rows[1].Cells[5].Range.Text = "Время\nвыполнения\nкоманды";
                section5.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                section5.Rows[1].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                
                section5.Rows.Add();

                //section5.Cell(2, 1).Range.Text = "testure";

                int nORs = 0;          //numberOfRows
                nORs = section5.Rows.Count;

                // Заполнение 5 пункта Программы переключений
                section5.Rows[nORs].Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop;
                section5.Rows[nORs].Cells[2].Range.Text = "5.";

                section5.Rows.Add();

                section5.Rows[nORs].Cells[3].Range.Font.Bold = 1;
                section5.Rows[nORs].Cells[3].Range.Text = "ПОРЯДОК И ПОСЛЕДОВАТЕЛЬНОСТЬ ВЫПОЛНЕНИЯ ОПЕРАЦИЙ:";

                


                section5.Rows[nORs].Cells[4].Merge(section5.Rows[nORs].Cells[5]);   // Объединение 4 и 5 ячейки в строке

                // Заполнение из actionsList
                for (int i = 0; i < actionsList.Count; i++)
                {
                    if (actionsList[i].FlName != "Начало")
                    {
                        nORs = section5.Rows.Count;
                        section5.Rows[nORs].Cells[1].Range.Text = actionsList[i].FlName;
                        section5.Rows.Add();

                        Row row3 = section5.Rows[nORs];
                        Cell fCell = row3.Cells[1];
                        foreach (Cell currCell in row3.Cells)
                        {
                            if (currCell.ColumnIndex != fCell.ColumnIndex && currCell.ColumnIndex != 1) // объединение ячеек только по правую сторону от выбраной
                            {
                                fCell.Merge(currCell);
                            }
                        }
                    }

                    if (actionsList[i].SecondLevelList.Count > 0)
                    {
                        for (int j = 0; j < actionsList[i].SecondLevelList.Count; j++)
                        {
                            nORs = section5.Rows.Count;
                            section5.Rows[nORs].Cells[1].Range.Text = actionsList[i].SecondLevelList[j].SlName;
                            section5.Rows.Add();
                            if (actionsList[i].SecondLevelList[j].slCommand == comboBox3.Text)
                            {
                                section5.Rows[nORs].Cells[4].Merge(section5.Rows[nORs].Cells[5]);   // Объединение 4 и 5 ячейки в строке
                            }
                            

                            if (actionsList[i].SecondLevelList[j].ThirdLevelList.Count > 0)
                            {
                                for (int k = 0; k < actionsList[i].SecondLevelList[j].ThirdLevelList.Count; k++)
                                {
                                    if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isNumerated == true)
                                    {
                                        section5.Rows[nORs].Cells[2].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].itemNumber;

                                        section5.Rows[nORs].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                                        section5.Rows[nORs].Cells[3].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;
                                        string gng = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;
                                        int lastNors = nORs;
                                        
                                        nORs = section5.Rows.Count;
                                        if (k != actionsList[i].SecondLevelList[j].ThirdLevelList.Count - 1)
                                        { section5.Rows.Add(); }
                                        int intovich = section5.Rows.Count;

                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].isConsistEquip == true)
                                        {
                                            section5.Rows[lastNors].Cells[4].Merge(section5.Rows[lastNors].Cells[5]);   // Объединение 4 и 5 ячейки в строке
                                        }
                                    }
                                    else
                                    {
                                        
                                        section5.Rows[nORs].Cells[2].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].TlName;

                                        Row row3 = section5.Rows[nORs];
                                        Cell fCell = row3.Cells[2];
                                        foreach (Cell currCell in row3.Cells)
                                        {
                                            if (currCell.ColumnIndex != fCell.ColumnIndex && currCell.ColumnIndex > 2) // объединение ячеек только по правую сторону от выбраной
                                            {
                                                fCell.Merge(currCell);
                                            }
                                        }
                                        section5.Rows[nORs].Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                        if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count > 0)
                                        {
                                            nORs = section5.Rows.Count;

                                            int forCombine4_5 = nORs;   // Для объединение по вертикали 4и пятого столбцов
                                            for (int l = 0; l < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; l++)
                                            {
                                                if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].isNumerated == true)
                                                {
                                                    string prov = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                    section5.Rows[nORs].Cells[2].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].itemNumber;
                                                    section5.Rows[nORs].Cells[3].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                }
                                                else
                                                {
                                                    section5.Rows[nORs].Cells[3].Range.Text = actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList[l].FlName;
                                                }

                                                section5.Rows[nORs].Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                                                /*if (l != actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count -1)
                                                { */
                                                section5.Rows.Add(); /*}*/
                                                nORs = section5.Rows.Count;
                                            }
                                            if (actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count > 1)
                                            {
                                                for (int d = 1; d < actionsList[i].SecondLevelList[j].ThirdLevelList[k].FourthLevelList.Count; d++)
                                                {
                                                    section5.Rows[forCombine4_5].Cells[4].Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
                                                    section5.Rows[forCombine4_5].Cells[4].TopPadding = 0;
                                                    //новый обращается через общий массив ячеек
                                                    //section5.Cell(forCombine4_5, 4).Merge(section5.Cell(forCombine4_5 + d, 4));
                                                    //старый код обращался к ячейке через коллекцию строк 
                                                    //section5.Rows[forCombine4_5].Cells[4].Merge(section5.Rows[forCombine4_5 + d].Cells[4]);
                                                }
                                            }
                                        }
                                        nORs = section5.Rows.Count;
                                        if (k != actionsList[i].SecondLevelList[j].ThirdLevelList.Count - 1)
                                        { section5.Rows.Add(); }
                                    }

                                    

                                    
                                }
                            }
                        }
                    }
                }

                section5.Borders.Enable = 1;

                int uuh = section5.Rows.Count;


                winword.Visible = true;         // Открытие созданного файла
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        partial void CreateWordFile();
        partial void GoToNextPage(Microsoft.Office.Interop.Word.Application app, Microsoft.Office.Interop.Word.Document doc,
            object missing/*, Microsoft.Office.Interop.Word.Paragraph para*/);
        
        /*public void DoWordFile()
        {
            CreateWordFile();
        }*/
        // Запуск создания программы переключений
        private void button_Click_1(object sender, RoutedEventArgs e)
        {
            /*progressBar1.IsIndeterminate = true;  
            // создаем новый поток
            Thread myThread = new Thread(CreateWordFile);
            myThread.Start(); // запускаем поток*/

            //StartCreateWordFileAsync();

            /*System.Threading.Tasks.Task.Factory.StartNew<string>(() => MainWindow.CreateWordFile(),
                                             TaskCreationOptions.LongRunning);*/
            this.Dispatcher.Invoke(() =>
            {
                StartCreateWordFileAsync();
            });

        }


        // Асинхронный метод, позволяющий показывать 
        // ProgressBar во время составления файла ПП
        public async void StartCreateWordFileAsync()
        {
            textBlock7.Visibility = Visibility.Hidden;
            progressBar1.Visibility = Visibility.Visible;
            progressBar1.IsIndeterminate = true; // выполняется синхронно
            await System.Threading.Tasks.Task.Run(() => CreateWordFile());                            // выполняется асинхронно
            progressBar1.IsIndeterminate = false;  // выполняется синхронно
            progressBar1.Visibility = Visibility.Hidden;
            //await System.Threading.Tasks.Task.Run(() => SuccessCreateOfWordFile());
            textBlock7.Visibility = Visibility.Visible;
            //Thread.Sleep(5000);
            //textBlock7.Visibility = Visibility.Hidden;
        }

        public void SuccessCreateOfWordFile()
        {
            textBlock7.Visibility = Visibility.Visible;
            Thread.Sleep(1000);
            textBlock7.Visibility = Visibility.Hidden;
        } // Метод пока что не используется

        private void button2_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application winword =
                    new Microsoft.Office.Interop.Word.Application();

                winword.Visible = false;

                object missing = System.Reflection.Missing.Value;

                //Создание нового документа
                Microsoft.Office.Interop.Word.Document document =
                    winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

                document.Content.SetRange(0, 0);



                //Добавление текста со стилем Заголовок 1
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add(ref missing);
                para1.Range.Text = " ";

                //Создание таблицы 5х5
                Microsoft.Office.Interop.Word.Table firstTable = document.Tables.Add(para1.Range, 1, 5, ref missing, ref missing);


                firstTable.Borders.Enable = 1;
                firstTable.Columns[1].PreferredWidth = 50;
                int ddd = 1;
                firstTable.Rows.Add();                
                firstTable.Rows.Add();
                ddd++;
                firstTable.Rows.Add();
                ddd++;
                firstTable.Cell(ddd - 1, 1).Merge(firstTable.Cell(ddd, 1));
                firstTable.Cell(ddd - 1,1).Range.Text = "wert";
                firstTable.Cell(ddd, 2).Range.Text = "fghj";
                firstTable.Rows.Add();
                ddd++;
                firstTable.Cell(ddd, 1).Range.Text = "PsPs";
                firstTable.Cell(ddd - 2, 1).Merge(firstTable.Cell(ddd, 1));

                /*foreach (Microsoft.Office.Interop.Word.Row row in firstTable.Rows)
                {
                    foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
                    {
                        //Заголовок таблицы
                        if (cell.RowIndex == 1)
                        {
                            cell.Range.Text = "Колонка " + cell.ColumnIndex.ToString();
                            cell.Range.Font.Bold = 1;
                            //Задаем шрифт и размер текста
                            cell.Range.Font.Name = "verdana";
                            cell.Range.Font.Size = 10;
                            cell.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray25;
                            //Выравнивание текста в заголовках столбцов по центру
                            cell.VerticalAlignment =
                            Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment =
                            Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                        //Значения ячеек
                        else
                        {
                            cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                        }
                    }
                }
                int rowNum = 3;
                Microsoft.Office.Interop.Word.Row row1 = firstTable.Rows[rowNum];
                Microsoft.Office.Interop.Word.Cell firstCell = row1.Cells[2];
                foreach (Microsoft.Office.Interop.Word.Cell currCell in row1.Cells)
                {
                    if (currCell.ColumnIndex != firstCell.ColumnIndex && currCell.ColumnIndex != 1) // объединение ячеек только по правую сторону от выбраной
                    {
                        firstCell.Merge(currCell);
                    }
                }
                               
                Microsoft.Office.Interop.Word.Row row2 = firstTable.Rows.Add();
                document.Range(row2.Cells[1].Range.Start, row2.Cells[2].Range.End).Cells.Merge();//Объединение ячеек в третьей снизу строке
                row2.Cells[1].Range.Text = "test ";//Запись текста в первую, уже объединённую ячейку

                row2 = firstTable.Rows.Add();                
                row2.Cells[1].Range.Text = "tes2 " ;//Запись текста в первую, уже объединённую ячейку


                firstTable.Cell(4, 3).Merge(firstTable.Cell(5, 3));

                int sch = firstTable.Rows.Count;
                row2 = firstTable.Rows.Add();
                firstTable.Rows.Add();

                firstTable.Cell(6,2).Range.Text = "asdf";
                firstTable.Rows.Add();*/
                winword.Visible = true;         // Открытие созданного файла
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        // Метод обновления значений основных параметров Программы переключений
        public void RefreshParamsOfSP ()
        {
            mainParamsOfSP.aim = comboBox1.Text;
            mainParamsOfSP.typeDO = comboBox3.Text;
            mainParamsOfSP.dispOffice = textBox1.Text;
            mainParamsOfSP.nameLine = textBox2.Text;
            if ((bool)checkBox4.IsChecked)
            {
                mainParamsOfSP.isACLineSegment = true;
            }
            else mainParamsOfSP.isACLineSegment = false;
            if (mainParamsOfSP.isACLineSegment == true)
            {
                mainParamsOfSP.ACLineSegment = textBox4.Text;
            }
            else
            {
                mainParamsOfSP.ACLineSegment = "non";
            }
            mainParamsOfSP.lineOrganisation = textBox3.Text;
            mainParamsOfSP.actionDate = datePicker1.Text;
            /*if ((bool)checkBox1.IsChecked)
            {
                mainParamsOfSP.inducedVoltage = true;
            }
            else mainParamsOfSP.inducedVoltage = false;
            if ((bool)checkBox2.IsChecked)
            {
                mainParamsOfSP.isUsedARM = true;
            }
            else mainParamsOfSP.isUsedARM = false;
            if ((bool)checkBox2.IsChecked)
            {
                mainParamsOfSP.ferroresonance = true;
            }
            else mainParamsOfSP.ferroresonance = false;*/
        }


        /*private void CheckBox_Checked_1(object sender, RoutedEventArgs e)
        {
            refreshOrgArrs();
        }*/

        // Событие при нажатии на чекбокс в Организационных мероприятиях
        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            listOfOldOrgArrs.Clear();
            for (int i = 0; i < listOfOrgArrs.Count; i++)
            {
                var uuu = new Org_arrangs();
                uuu.NameObj = listOfOrgArrs[i].NameObj;
                uuu.PObject = listOfOrgArrs[i].PObject;
                uuu.isWork = listOfOrgArrs[i].isWork;
                listOfOldOrgArrs.Add(uuu);
            }
        }

        // Событие при нажатии на чекбокс в Списке энергообъектов
        private void CheckBox_Click_1(object sender, RoutedEventArgs e)
        {
            refreshOrgArrs();
        }

       
        // Событие при изменении даты ПП
        private void datePicker1_CalendarClosed(object sender, RoutedEventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении цели ПП
        private void comboBox1_DropDownClosed(object sender, EventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении типа организации от СО
        private void comboBox3_DropDownClosed(object sender, EventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении названия организации от СО
        private void textBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении типа сегмента ВЛ
        private void checkBox4_Click(object sender, RoutedEventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении Названия ВЛ
        private void textBox2_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении Названия сегмента ВЛ
        private void textBox4_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshParamsOfSP();
        }

        // Событие при изменении Названия организации-собственника ВЛ
        private void textBox3_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshParamsOfSP();
        }


        // Кнопка "Персонал"
        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            refreshpObjectsOfSP();

            PersonalOption personal = new PersonalOption();
            personal.Owner = this;
            personal.ShowDialog();            
        }

        private void refreshpObjectsOfSP ()
        {
            listOfPersonalTemp.Clear();
            // Поэлементно присваиваем временному списку все значения постоянного списка
            for (int i = 0; i < listOfPersonal.Count; i++)
            {
                var newpers = new Personal("");
                newpers.organisationOfPersonal = listOfPersonal[i].organisationOfPersonal;
                for (int k = 0; k < listOfPersonal[i].Person.Count; k++)
                {
                    var newpers2 = new PersonalClass("");
                    newpers2.nameOfPerson = listOfPersonal[i].Person[k].nameOfPerson;
                    newpers2.role = listOfPersonal[i].Person[k].role;
                    newpers2.action = listOfPersonal[i].Person[k].action;

                    newpers.Person.Add(newpers2);
                }
                listOfPersonalTemp.Add(newpers);
            }

            listOfPersonal.Clear();

            // В начале добавляются персонал из РДУ и ЦУС
            var newpersSO = new Personal("");            

            newpersSO.organisationOfPersonal = mainParamsOfSP.dispOffice;

            if (newpersSO.organisationOfPersonal == listOfPersonalTemp[0].organisationOfPersonal)
            {
                for (int i = 0; i < listOfPersonalTemp[0].Person.Count; i++)
                {
                    var newpers2 = new PersonalClass("");
                    newpers2.nameOfPerson = listOfPersonalTemp[0].Person[i].nameOfPerson;
                    newpers2.role = listOfPersonalTemp[0].Person[i].role;
                    newpers2.action = listOfPersonalTemp[0].Person[i].action;

                    newpersSO.Person.Add(newpers2);
                }
            }
            else
            {
                var newpers2 = new PersonalClass("");
                newpers2.nameOfPerson = "";
                newpers2.role = "";
                newpers2.action = "";

                newpersSO.Person.Add(newpers2);
            }

            listOfPersonal.Add(newpersSO);

            newpersSO = new Personal("");

            newpersSO.organisationOfPersonal = mainParamsOfSP.lineOrganisation;

            if (newpersSO.organisationOfPersonal == listOfPersonalTemp[1].organisationOfPersonal)
            {
                for (int i = 0; i < listOfPersonalTemp[0].Person.Count; i++)
                {
                    var newpers2 = new PersonalClass("");
                    newpers2.nameOfPerson = listOfPersonalTemp[1].Person[i].nameOfPerson;
                    newpers2.role = listOfPersonalTemp[1].Person[i].role;
                    newpers2.action = listOfPersonalTemp[1].Person[i].action;

                    newpersSO.Person.Add(newpers2);
                }
            }
            else
            {
                var newpers2 = new PersonalClass("");
                newpers2.nameOfPerson = "";
                newpers2.role = "";
                newpers2.action = "";

                newpersSO.Person.Add(newpers2);
            }

            listOfPersonal.Add(newpersSO);

            // Проход по всем ЭО
            for (int i = 0; i < listOfPowerObjects.Count; i++)
            {
                // Если ЭО задействован в работе
                if (listOfPowerObjects[i].isUsed == true)
                {
                    var newpers = new Personal("");

                    newpers.organisationOfPersonal = listOfPowerObjects[i].NamePO;

                    bool find = false;
                    for (int j = 0; j < listOfPersonalTemp.Count; j++)
                    {
                        // Если название ЭО из списка объектов совпадает с назнанием ЭО у персонала, то...
                        if (listOfPersonalTemp[j].organisationOfPersonal == newpers.organisationOfPersonal)
                        {
                            for (int k = 0; k < listOfPersonalTemp[j].Person.Count; k++)
                            {
                                var newpers2 = new PersonalClass("");
                                newpers2.nameOfPerson = listOfPersonalTemp[j].Person[k].nameOfPerson;
                                newpers2.role = listOfPersonalTemp[j].Person[k].role;
                                newpers2.action = listOfPersonalTemp[j].Person[k].action;

                                newpers.Person.Add(newpers2);
                            }
                            find = true;
                        }
                        else
                        {
                            /*var newpers2 = new PersonalClass("");
                            newpers2.nameOfPerson = "";
                            newpers2.role = "";
                            newpers2.action = "";

                            newpers.Person.Add(newpers2);*/
                        }
                    }

                    // Если совпадений не найдено, то добавляется пустое значение
                    if (find == false)
                    {
                        var newpers2 = new PersonalClass("");
                        newpers2.nameOfPerson = "";
                        newpers2.role = "";
                        newpers2.action = "";

                        newpers.Person.Add(newpers2);
                    }

                    listOfPersonal.Add(newpers);
                    /*for (int j = 0; j < listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person.Count; j++)
                    {
                        var newpers2 = new PersonalClass("");
                        newpers2.nameOfPerson = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person[j].nameOfPerson;
                        newpers2.role = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person[j].role;
                        newpers2.action = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person[j].action;

                        newpers.Person.Add(newpers2);
                    }*/
                }
            }

            // В конце добавляется персонал из СО
            newpersSO = new Personal("");
            int lastnum = listOfPersonalTemp.Count - 1;

            newpersSO.organisationOfPersonal = listOfPersonalTemp[lastnum].organisationOfPersonal;

            for (int i = 0; i < listOfPersonalTemp[lastnum].Person.Count; i++)
            {
                var newpers2 = new PersonalClass("");
                newpers2.nameOfPerson = listOfPersonalTemp[lastnum].Person[i].nameOfPerson;
                newpers2.role = listOfPersonalTemp[lastnum].Person[i].role;
                newpers2.action = listOfPersonalTemp[lastnum].Person[i].action;

                newpersSO.Person.Add(newpers2);
            }

            listOfPersonal.Add(newpersSO);


            /*
            int numCheckPO = 0;
            
            // Считается количество задействованных подстанций
            for (int i = 0; i < listOfPowerObjects.Count; i++)
            {
                if (listOfPowerObjects[i].isUsed == true)
                    numCheckPO++;
            }

            // Если количество Энергообъектов в Персонале не равно задействованным объектам
            if (listOfPersonal.Count - 1 != numCheckPO)
            {
                listOfPersonalTemp = listOfPersonal;
                listOfPersonal.Clear();
                var newpers = new Personal("");

                for (int i = 0; i < numCheckPO + 1; i++)
                {
                    if (i == numCheckPO)
                    {
                        newpers.organisationOfPersonal = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].organisationOfPersonal;
                        for (int j = 0; j < listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person.Count; j++)
                        {
                            var newpers2 = new PersonalClass("");
                            newpers2.nameOfPerson = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person[j].nameOfPerson;
                            newpers2.role = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person[j].role;
                            newpers2.action = listOfPersonalTemp[listOfPersonalTemp.Count() - 1].Person[j].action;

                            newpers.Person.Add(newpers2);
                        }
                    }
                    else
                    {

                    }
                }
            }*/
        }

        // Событие при изменении параметра Наведенного напряжения
        private void checkBox1_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)checkBox1.IsChecked)
            {
                mainParamsOfSP.inducedVoltage = true;
            }
            else mainParamsOfSP.inducedVoltage = false;           
        }

        // Событие при изменении параметра использования АРМ
        private void checkBox2_Click(object sender, RoutedEventArgs e)
        {
            if ((bool)checkBox2.IsChecked)
            {
                mainParamsOfSP.isUsedARM = true;
            }
            else mainParamsOfSP.isUsedARM = false;            
        }

        // Событие при изменении параметра наличия феррорезонанса
        private void checkBox3_Click(object sender, RoutedEventArgs e)
        {            
            if ((bool)checkBox3.IsChecked)
            {
                mainParamsOfSP.ferroresonance = true;
            }
            else mainParamsOfSP.ferroresonance = false;
        }

        
    }




    // Класс конвертера для TreeView
    public class tvFontConverter : IValueConverter
    {        
            object IValueConverter.Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
                var tvcontent = Convert.ToString(value);
                var tvdict = MainWindow.tvKeywords;
                var tvoutString = "<TextBlock xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"  xml:space=\"preserve\">";
                foreach (var word in tvcontent.Split(' '))
                {
                    var converted = word;
                    FontWeight fs;
                    if (tvdict.TryGetValue(word, out fs))
                    {
                        var run = new Run(word);
                        run.FontWeight = fs;
                        converted = System.Windows.Markup.XamlWriter.Save(run);
                    }
                    tvoutString += converted + " ";
                }
                tvoutString += "</TextBlock>";
                var prov = System.Windows.Markup.XamlReader.Parse(tvoutString);
                return System.Windows.Markup.XamlReader.Parse(tvoutString);
            }
            object IValueConverter.ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
                throw new NotImplementedException();
            }

    }


    // Класс конвертера для моего TreeView
    public class equipFontConverter : IValueConverter
    {
        object IValueConverter.Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var tvcontent = Convert.ToString(value);
            var tvdict = MainWindow.equipKeywords;
            var tvoutString = "<TextBlock xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"  xml:space=\"preserve\">";
            foreach (var word in tvcontent.Split(' '))
            {
                var converted = word;
                FontWeight fs;
                if (tvdict.TryGetValue(word, out fs))
                {
                    var run = new Run(word);
                    run.FontWeight = fs;
                    converted = System.Windows.Markup.XamlWriter.Save(run);
                }
                tvoutString += converted + " ";
            }
            tvoutString += "</TextBlock>";
            var prov = System.Windows.Markup.XamlReader.Parse(tvoutString);
            return System.Windows.Markup.XamlReader.Parse(tvoutString);
        }
        object IValueConverter.ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }

    }

}
