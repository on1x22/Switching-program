using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfApplication1
{
    public class PowerObject
    {
        public string NamePO { get; set; }
        public bool isUsed { get; set; }
        public string organisationPO { get; set; }
        public void ups()
        { }

        public static implicit operator bool(PowerObject v)
        {
            throw new NotImplementedException();
        }

        public class Equipment : PowerObject
        {
            public string nameEquip { get; set; }
            public bool stateEquip { get; set; }
            public string typeEquip { get; set; }
        }
        


    }
}
