﻿using Eplan.EplAddin.Actions;
using Eplan.EplApi.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using EPLAN_API_2024.Forms;

namespace EPLAN_API_2024
{
    class ConfigAction_2024_Service : IActionService
    {
        public ConfigAction_2024_Service()
        {
        }
        public void Execute()
        {
            Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            Application.Run(new Configurador());
        }
    }
}
