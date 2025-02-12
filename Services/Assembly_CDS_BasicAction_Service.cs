using Eplan.EplAddin.Actions;
using EPLAN_API.API_Basic;

namespace EPLAN_API_2022
{
    class Assembly_CDS_BasicAction_Service : IActionService
    {
        public Assembly_CDS_BasicAction_Service()
        {
        }
        public void Execute()
        {
            new DrawCDSBasicEN();
        }
    }
}
