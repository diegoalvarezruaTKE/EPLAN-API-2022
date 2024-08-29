using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using static System.Net.Mime.MediaTypeNames;
using System;
using Eplan.EplApi.Gui;

namespace EplanAddIn
{
    public class AddInModule : IEplAddIn
    {
        public bool OnRegister(ref System.Boolean bLoadOnStart)
        {
            bLoadOnStart = true;
            return true;
        }
        public bool OnUnregister()
        {
            return true;
        }
        public bool OnInit()
        {
            return true;
        }
        public bool OnInitGui()
        {
            var ribbonBar = new RibbonBar();
            var ribbonTab = ribbonBar.AddTab("API 2022");
            var ribbonCommandGroup = ribbonTab.AddCommandGroup("Config");
            ribbonCommandGroup.AddCommand("Configurador", "ConfigAction");
            ribbonCommandGroup.AddCommand("Lista de Materiales", "MaterialesAction");
            return true;
        }
        public bool OnExit()
        {
            return true;
        }
    }


}
