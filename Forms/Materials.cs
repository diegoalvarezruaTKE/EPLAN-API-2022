using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using EPLAN_API.SAP;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;

namespace EPLAN_API.Forms
{
    public partial class Materials : Form
    {
        private Project oProject;
        private Page[] oPages;
        private SortedList Caracteristics;

        public Materials()
        {
            InitializeComponent();
            oProject = new ProjectManager().CurrentProject;
            //String [] locations= oProject.GetLocations();
            cB_Locations.Items.AddRange(oProject.GetLocations(Project.Hierarchy.Location));
            cB_Locations.SelectedIndex = 0;

        }

        private void execute(string location, int cablemultiplier)
        {
            oProject = new ProjectManager().CurrentProject;
            oPages = oProject.Pages;

            FunctionsFilter functionsFilter = new FunctionsFilter();
            //functionsFilter.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.DeviceEndTerminal;
            FunctionPropertyList functionsPropertyList = new FunctionPropertyList();
            functionsPropertyList.FUNC_MAINFUNCTION = true;
            functionsPropertyList.DESIGNATION_LOCATION = location;
            functionsFilter.SetFilteredPropertyList(functionsPropertyList);

            DMObjectsFinder DMObjectsFinder = new DMObjectsFinder(oProject);

            Function[] functions = DMObjectsFinder.GetFunctions(functionsFilter);

            FunctionsFilter functionsFilter1 = new FunctionsFilter();
            //functionsFilter.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.DeviceEndTerminal;
            FunctionPropertyList functionsPropertyList1 = new FunctionPropertyList();
            functionsPropertyList1.FUNC_MAINFUNCTION = true;
            functionsFilter1.SetFilteredPropertyList(functionsPropertyList1);

            DMObjectsFinder DMObjectsFinder1 = new DMObjectsFinder(oProject);

            Function[] functions1 = DMObjectsFinder1.GetFunctions(functionsFilter1);

        }

        private void b_List_Click(object sender, EventArgs e)
        {
            ProjectSettings projectSettings = new ProjectSettings(oProject);
            int unidadMedida = projectSettings.GetNumericSetting("CableLog.CableManuell.UnitOfCableLength", 0);
            if (unidadMedida == 0)
                //0=m
                execute(cB_Locations.SelectedItem.ToString(), 100);
            else
                //5=cm
                execute(cB_Locations.SelectedItem.ToString(), 1);
        }
    }

}
