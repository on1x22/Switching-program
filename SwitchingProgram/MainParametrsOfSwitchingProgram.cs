using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication1
{
    public class MainParametrsOfSwitchingProgram
    {
        public string aim;
        public string typeDO;
        public string dispOffice;
        public string nameLine;
        public bool isACLineSegment;
        public string ACLineSegment;
        public string lineOrganisation;
        public string actionDate;

        // по п.п. 3.2. - 3.4. 
        public bool inducedVoltage;// { get; set; }
        public bool isUsedARM; // { get; set; }
        public bool ferroresonance;// { get; set; }


        public string Aim
        {
            get { return aim; }
            set { aim = value; }
        }
    }
}
