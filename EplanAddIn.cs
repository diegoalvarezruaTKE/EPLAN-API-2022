using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using static System.Net.Mime.MediaTypeNames;
using System;

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
            return true;
        }
        public bool OnExit()
        {
            return true;
        }
    }


}
