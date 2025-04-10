using Eplan.EplAddin.Actions;
using EPLAN_API.API_Basic;


namespace EPLAN_API_2024
{
    class Assembly_Main_Cab_BasicAction_Service : IActionService
    {
        public Assembly_Main_Cab_BasicAction_Service()
        {
        }
        public void Execute()
        {
            new DrawMainCabinetBasicEN();
        }
    }
}
