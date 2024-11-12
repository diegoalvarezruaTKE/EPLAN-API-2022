using EPLAN_API.SAP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API_2022.Forms
{
    public partial class PanelGEC : Form
    {
        public List<Electric> oElectricList;

        public PanelGEC(List<Electric> electricList)
        {
            oElectricList = electricList;

            InitializeComponent();
        }
    }
}
