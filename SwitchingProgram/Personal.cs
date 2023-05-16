using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;

namespace WpfApplication1
{
    public class Personal
    {
        public string organisationOfPersonal { get; set; }

        /*public string OrganisationOfPersonal
        {
                get { return organisationOfPersonal; }
                set { organisationOfPersonal = value; }
        }*/

        public string nameOfPerson1 { get; set; }   
        public string role1 { get; set; }

        //список входящих в класс первого уровня экземпляров класса второго уровня
        private ObservableCollection<PersonalClass> person;

        public ObservableCollection<PersonalClass> Person
        {
                get { return person; }
                set { person = value; }
        }

        //класс второго уровня
        public Personal(string _organisationOfPersonal)
        {
                //OrganisationOfPersonal = _organisationOfPersonal;
                Person = new ObservableCollection<PersonalClass>();
        }
    }


    public class PersonalClass
    {
        public string nameOfPerson { get; set; }
        public string role { get; set; }
        public string action { get; set; }

        /*public string Person
        {
                get { return person; }
                set { person = value; }
        }*/

        //класс третьего уровня
        public PersonalClass(string _person)
        {
            nameOfPerson = _person;
        }
    }      
}
