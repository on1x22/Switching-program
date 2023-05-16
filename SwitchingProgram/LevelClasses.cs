using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;

namespace WpfApplication1
{
    //класс первого уровня
    public class FirstLevelClass
    {
        private string flName;

        public string flCommand;

        public string FlName
        {
            get { return flName; }
            set { flName = value; }
        }        

        //список входящих в класс первого уровня экземпляров класса второго уровня
        private ObservableCollection<SecondLevelClass> secondLevelList;
        //private string v;

        /*public FirstLevelClass(string v)
        {
            this.v = v;
        }*/

        public ObservableCollection<SecondLevelClass> SecondLevelList
        {
            get { return secondLevelList; }
            set { secondLevelList = value; }
        }

        //класс второго уровня
        public FirstLevelClass( string _flName)
        {            
            FlName = _flName;
            SecondLevelList = new ObservableCollection<SecondLevelClass>();            
        }
    }


    public class SecondLevelClass
    {
        string slName;
        public string slCommand;

        public string SlName
        {
            get { return slName; }
            set { slName = value; }
        }

        //public int Level = 2;

        //список входящих в класс второго уровня экземпляров класса третьего уровня
        private ObservableCollection<ThirdLevelClass> thirdLevelList;

        public ObservableCollection<ThirdLevelClass> ThirdLevelList
        {
            get { return thirdLevelList; }
            set { thirdLevelList = value; }
        }

        //класс третьего уровня
        public SecondLevelClass(string _slName)
        {
            SlName = _slName;
            ThirdLevelList = new ObservableCollection<ThirdLevelClass>();
        }
    }


    public class ThirdLevelClass
    {
        string tlName;
        public string tlCommand;
        public string equipmentName;
        public bool isNumerated;
        public string itemNumber;
        public bool isConsistEquip;

        public string TlName
        {
            get { return tlName; }
            set { tlName = value; }
        }

        //список входящих в класс третьего уровня экземпляров класса четвертого уровня
        private ObservableCollection<FourthLevelClass> fourthLevelList;

        public ObservableCollection<FourthLevelClass> FourthLevelList
        {
            get { return fourthLevelList; }
            set { fourthLevelList = value; }
        }

        //класс четвертого уровня
        public ThirdLevelClass(string _tlName)
        {
            TlName = _tlName;
            FourthLevelList = new ObservableCollection<FourthLevelClass>();
        }
    }


    public class FourthLevelClass
    {
        string flName;
        public string fourthlCommand;
        public string equipmentName;
        public bool isNumerated;
        public string itemNumber;
        public bool isConsistEquip;

        public string FlName
        {
            get { return flName; }
            set { flName = value; }
        }

        public FourthLevelClass(string _flName)
        {
            FlName = _flName;
        }
    }
}
