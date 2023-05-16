using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;

namespace WpfApplication1
{
    public class Node
    {
        public String Name { get; set; }
        public ObservableCollection<Node> Nodes { get; set; }
    }
}
