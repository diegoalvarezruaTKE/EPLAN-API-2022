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
            var ribbonTab = ribbonBar.AddTab("API 2024");
            var ribbonCommandGroupConfig = ribbonTab.AddCommandGroup("Config");
            ribbonCommandGroupConfig.AddCommand("Configurador 2024", "ConfigAction_2024");
            ribbonCommandGroupConfig.AddCommand("Lista de Materiales", "MaterialesAction2024");
            var ribbonCommandGroupBasic = ribbonTab.AddCommandGroup("Basic");
            ribbonCommandGroupBasic.AddCommand("Ensambla Armario Basic", "Assembly_Main_Cab_Basic");
            ribbonCommandGroupBasic.AddCommand("Ensambla CDS Basic", "Assembly_CDS_Basic");
            ribbonCommandGroupBasic.AddCommand("Ensambla CDI Basic", "Assembly_CDI_Basic");
            return true;
        }
        public bool OnExit()
        {
            return true;
        }
    }


}
