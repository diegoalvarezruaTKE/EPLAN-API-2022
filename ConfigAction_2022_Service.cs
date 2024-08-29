using Eplan.EplAddin.Actions;
using Eplan.EplApi.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API_2022
{
    class ConfigAction_2022_Service : IActionService
    {
        public ConfigAction_2022_Service()
        {
        }
        public void Execute()
        {
            new Decider().Decide(EnumDecisionType.eOkDecision, "CSharpAction was called!", "Eplan.EplAddIn.Demo1", EnumDecisionReturn.eOK, EnumDecisionReturn.eOK);
            // TODO: Add your Code here
        }
    }
}
