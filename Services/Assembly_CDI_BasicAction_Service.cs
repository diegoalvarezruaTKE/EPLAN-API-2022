using Eplan.EplAddin.Actions;
using Eplan.EplApi.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPLAN_API_2022.Forms;
using EPLAN_API.User;

namespace EPLAN_API_2022
{
    class Assembly_CDI_BasicAction_Service : IActionService
    {
        public Assembly_CDI_BasicAction_Service()
        {
        }
        public void Execute()
        {
            new DrawCDIBasicEN();
        }
    }
}
