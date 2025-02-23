using Eplan.EplAddin.Actions;
using Eplan.EplApi.Base;
using EPLAN_API.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API.Services
{
    public class MaterialesActionService : IActionService
    {

        public MaterialesActionService()
        {

        }

        #region IActionService Implementation

        public void Execute()
        {
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Materials());
        }



        #endregion
    }

}
