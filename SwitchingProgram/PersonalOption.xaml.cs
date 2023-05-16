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
    /// Логика взаимодействия для Personal.xaml
    /// </summary>
    public partial class PersonalOption : Window
    {
        List<Personal> listOfPers = new List<Personal>();           //Общий список персонала 
        List<Personal> listOfTablePers = new List<Personal>();      //Список персонала в таблице
        List<PersonalClass> listOfSOPers = new List<PersonalClass>();         //Список персонала от СО 
        List<PersonalClass> listOfSelectedPers = new List<PersonalClass>();

        int position = -1;

        ProgramOptions progOpt = new ProgramOptions();
        //MainWindow main = this.Owner as MainWindow;

        public PersonalOption()
        {
            InitializeComponent();
            MainWindow main = this.Owner as MainWindow;
            //progOpt = main.progOption;
        }

        // Дейстивя при открытии окна
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            listOfPers = main.listOfPersonal;

            for (int i = 0; i < listOfPers.Count; i++)
            {
                if (listOfPers[i].organisationOfPersonal != "SO")
                {
                    var newpers = new Personal("");

                    newpers.organisationOfPersonal = listOfPers[i].organisationOfPersonal;

                    for (int j = 0; j < listOfPers[i].Person.Count; j++)
                    {
                        /*var newpers2 = new PersonalClass("");
                        newpers2.nameOfPerson = listOfPers[i].Person[j].nameOfPerson;
                        newpers2.role = listOfPers[i].Person[j].role;*/

                        newpers.nameOfPerson1 = listOfPers[i].Person[j].nameOfPerson;
                        newpers.role1 = listOfPers[i].Person[j].role;

                        /*newpers.Person.Add(newpers2);*/
                    }

                    listOfTablePers.Add(newpers);
                }
                else
                {
                    //var newpers = new Personal("");

                    //newpers.organisationOfPersonal = listOfPers[i].organisationOfPersonal;

                    for (int j = 0; j < listOfPers[i].Person.Count; j++)
                    {
                        var newpers2 = new PersonalClass("");
                        newpers2.nameOfPerson = listOfPers[i].Person[j].nameOfPerson;
                        newpers2.role = listOfPers[i].Person[j].role;
                        newpers2.action = listOfPers[i].Person[j].action;

                        listOfSOPers.Add(newpers2);
                    }

                    //listOfSOPers.Add(newpers);
                }
            }
            dataGrid1.ItemsSource = listOfTablePers;
            dataGrid2.ItemsSource = listOfSOPers;
        }

        /*private void button1_Click(object sender, RoutedEventArgs e)
        {

            var pers = new Personal("ha");
            pers.organisationOfPersonal = "RDU";            

            var persona = new PersonalClass("");
            persona.nameOfPerson = "Vasiliev R.D.";
            persona.role = "Dispatcher";
            pers.Person.Add(persona);

            listOfPers.Add(pers);

            pers = new Personal("ha");
            pers.organisationOfPersonal = "PMES";            

            persona = new PersonalClass("");
            persona.nameOfPerson = "Ivanov P.O.";
            persona.role = "DEM";
            pers.Person.Add(persona);
            listOfPers.Add(pers);
            
            pers = new Personal("ha");
            pers.organisationOfPersonal = "SO";
            
            persona = new PersonalClass("");
            persona.nameOfPerson = "Kukushin P.O.";
            persona.role = "Duty (Senior) dispatcher Nizhegorodskogo RDU";
            pers.Person.Add(persona);

            persona = new PersonalClass("");
            persona.nameOfPerson = "Pishkin F.P.";
            persona.role = "Duty dispatcher Nizhegorodskogo RDU";
            pers.Person.Add(persona);

            persona = new PersonalClass("");
            persona.nameOfPerson = "Glavniy K.U.";
            persona.role = "Senior dispatcher Nizhegorodskogo RDU";
            pers.Person.Add(persona);

            listOfPers.Add(pers);

            

            dataGrid1.ItemsSource = listOfTablePers;
        }*/

        

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            listOfPers.Clear();

            // Добавление персонала энергообъектов в общий список
            for (int i = 0; i < listOfTablePers.Count; i++)
            {
                var newpers = new Personal("");

                newpers.organisationOfPersonal = listOfTablePers[i].organisationOfPersonal;

                /*for (int j = 0; j < listOfTablePers[i].Person.Count; j++)
                {*/
                    var newpers2 = new PersonalClass("");
                    newpers2.nameOfPerson = listOfTablePers[i].nameOfPerson1;
                    newpers2.role = listOfTablePers[i].role1;
                    newpers2.action = "";

                    newpers.Person.Add(newpers2);
                /*}*/

                listOfPers.Add(newpers);
            }

            // Добавление персонала от СО в общий список
            for (int i = 0; i < 1; i++)
            {
                var newpers = new Personal("");

                newpers.organisationOfPersonal = "SO";

                for (int j = 0; j < listOfSOPers.Count; j++)
                {
                    var newpers2 = new PersonalClass("");
                    newpers2.nameOfPerson = listOfSOPers[j].nameOfPerson;
                    newpers2.role = listOfSOPers[j].role;
                    newpers2.action = listOfSOPers[j].action;

                    newpers.Person.Add(newpers2);
                }
                listOfPers.Add(newpers);
            }

            main.listOfPersonal = listOfPers;

            this.Close();
        }
    }
}
