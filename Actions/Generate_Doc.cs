using Eplan.EplAddin.Actions;
using Eplan.EplApi.ApplicationFramework;
using EPLAN_API_2024;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.Actions
{
    public class Generate_Doc : IEplAction
    {
        ///<summary>
        ///This function is called when executing the action.
        ///</summary>
        ///<returns>true, if the action performed successfully</returns>
        public bool Execute(ActionCallingContext ctx)
        {
            IActionService service = new Generate_Doc_Service();

            service.Execute();

            return true;
        }
        ///<summary>
        ///This function is called by the application framework, when registering the add-in.
        ///</summary>
        ///<param name="Name">The action is registered in EPLAN under this name</param>
        ///<param name="Ordinal">The action is registered with this overload priority</param>
        ///<returns>true, if OnRegister succeeds</returns>
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "Generate_Doc";
            Ordinal = 23;
            return true;
        }
        ///<summary>
        /// Documentation function for the action, which is called by EPLAN on demand
        /// returns the descriptive text for the action itself and if the action takes string parameters
        /// (command line), it also provides the name and description of each parameter
        ///</summary>
        ///<param name="actionProperties"> This object needs to be filled with information about the action
        ///</param>
        public void GetActionProperties(ref ActionProperties actionProperties)
        {

        }

        public bool Enabled(string strActionName, ActionCallingContext actionContext)
        {
            if (strActionName == "TESTACTION")
                return false;
            else
                return true;
        }
    }

}
