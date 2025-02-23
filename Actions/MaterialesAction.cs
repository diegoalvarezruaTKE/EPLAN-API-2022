using Eplan.EplAddin.Actions;
using Eplan.EplApi.ApplicationFramework;
using EPLAN_API.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.Actions
{
    class MaterialesAction : IEplAction, IEplActionEnable
    {
        #region IEplAction Members

        /// <summary>
        /// Execution of the Action.  
        /// </summary>
        /// <returns>True:  Execution of the Action was successful</returns>
        public bool Execute(ActionCallingContext ctx)
        {
            IActionService service = new MaterialesActionService();

            service.Execute();

            return true;
        }
        /// <summary>
        /// Function is called through the ApplicationFramework
        /// </summary>
        /// <param name="Name">Under this name, this Action in the system is registered</param>
        /// <param name="Ordinal">With this overload priority, this Action is registered</param>
        /// <returns>true: the return parameters are valid</returns>
        public bool OnRegister(ref string Name, ref int Ordinal)
        {
            Name = "MaterialesAction2024";
            Ordinal = 20;
            return true;
        }

        /// <summary>
        /// Documentation function for the Action; is called of the system as required 
        /// Bescheibungstext delivers for the Action itself and if the Action String-parameters ("Kommandozeile")
        /// also name and description of the single parameters evaluates
        /// </summary>
        /// <param name="actionProperties"> This object must be filled with the information of the Action.</param>
        public void GetActionProperties(ref ActionProperties actionProperties)
        {

        }

        #endregion

        #region IEplActionEnable Members

        public bool Enabled(string strActionName, ActionCallingContext actionContext)
        {
            if (strActionName == "TESTACTION")
                return false;
            else
                return true;
        }

        #endregion

    }

}
