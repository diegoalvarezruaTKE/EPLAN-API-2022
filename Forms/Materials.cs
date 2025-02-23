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
            FunctionPropertyList functionsPropertyList = new FunctionPropertyList();
            functionsPropertyList.FUNC_MAINFUNCTION = true;
            functionsPropertyList.DESIGNATION_LOCATION = location;
            functionsFilter.SetFilteredPropertyList(functionsPropertyList);
            DMObjectsFinder DMObjectsFinder = new DMObjectsFinder(oProject);
            Function[] functions = DMObjectsFinder.GetFunctions(functionsFilter);

            List<SAPMaterial> materials = new List<SAPMaterial>();

            foreach (Function function in functions) 
            {
                foreach (ArticleReference articleReference in function.ArticleReferences)
                {
                    SAPMaterial material = new SAPMaterial();
                    //SAP Code
                    material.SAPCode = articleReference.Properties.ARTICLE_ERPNR.ToString();
                    //SAP Name
                    LanguageList SAPNamelanguageList = new LanguageList();
                    articleReference.Properties.ARTICLE_DESCR2.ToMultiLangString().GetLanguageList(ref SAPNamelanguageList);
                    material.SAPName = articleReference.Properties.ARTICLE_DESCR2.ToMultiLangString().GetStringToDisplay(SAPNamelanguageList.get_Language(0));
                    if (SAPNamelanguageList.Count > 1)
                        ;
                    //SAP Description L1
                    material.SAPDescriptionL1 = function.VisibleName;
                    //SAP Description L2
                    LanguageList SAPDescriptionL2languageList = new LanguageList();
                    function.Properties.FUNC_TEXT.ToMultiLangString().GetLanguageList(ref SAPDescriptionL2languageList);
                    material.SAPDescriptionL2 = function.Properties.FUNC_TEXT.ToMultiLangString().GetStringToDisplay(ISOCode.Language.L_en_US);
                    if (SAPDescriptionL2languageList.Count > 1)
                        ;
                    //Aporte
                    if (articleReference.Properties.ARTICLE_SUPPLIER.ToString().Equals("PROV"))
                        material.Aporte = "L";
                    //Fabricante
                    //Referencia Fabricante



                    materials.Add(material);
                }

            }

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

    /*
     *                 oSheet.Cells[1, 3] = "SAP Code";
                oSheet.Cells[1, 4] = "Description L1";
                oSheet.Cells[1, 5] = "Description L2";
                oSheet.Cells[1, 6] = "Count";
                oSheet.Cells[1, 7] = "Aporte";
                oSheet.Cells[1, 8] = "Relevancia Fab.";
                oSheet.Cells[1, 9] = "Calculo Coste";
                oSheet.Cells[1, 10] = "Nombre SAP";
                oSheet.Cells[1, 11] = "Fabricante";
                oSheet.Cells[1, 12] = "Referencia Fabricante";
    */

    public class SAPMaterial
    {
        public string SAPCode { get; set; }

        public string SAPName { get; set; }

        public string SAPDescriptionL1 { get; set; }

        public string SAPDescriptionL2 { get; set; }

        public double Count { get; set; }

        public string Aporte { get; set; }

        public string Fabricante { get; set; }

        public string RefFabricante { get; set; }

        public SAPMaterial() { }

    }

}
