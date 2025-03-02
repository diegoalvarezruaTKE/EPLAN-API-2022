using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using EPLAN_API.SAP;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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
                    if (SAPNamelanguageList.Count > 0)
                    {
                        if (SAPNamelanguageList.Contains(ISOCode.Language.L_en_US))
                            material.SAPName = articleReference.Properties.ARTICLE_DESCR2.ToMultiLangString().GetStringToDisplay(ISOCode.Language.L_en_US);
                        else
                            material.SAPName = articleReference.Properties.ARTICLE_DESCR2.ToMultiLangString().GetStringToDisplay(SAPNamelanguageList.get_Language(0));
                    }

                    //SAP Description L1
                    material.SAPDescriptionL1 = function.VisibleName;

                    //SAP Description L2
                    try
                    {
                        LanguageList SAPDescriptionL2languageList = new LanguageList();
                        function.Properties.FUNC_TEXT.ToMultiLangString().GetLanguageList(ref SAPDescriptionL2languageList);
                        if (SAPDescriptionL2languageList.Count > 0)
                        {
                            if (SAPDescriptionL2languageList.Contains(ISOCode.Language.L_en_US))
                                material.SAPDescriptionL2 = function.Properties.FUNC_TEXT.ToMultiLangString().GetStringToDisplay(ISOCode.Language.L_en_US);
                            else
                                material.SAPDescriptionL2 = function.Properties.FUNC_TEXT.ToMultiLangString().GetStringToDisplay(SAPDescriptionL2languageList.get_Language(0));

                        }
                    }
                    catch (Exception ex)
                    {
                        material.SAPDescriptionL2 = "";
                    }

                    //Aporte
                    if (articleReference.Properties.ARTICLE_SUPPLIER.ToString().Equals("PROV"))
                        material.Aporte = "L";
                    else
                        material.Aporte = "";

                    //Calculo de coste
                    if (articleReference.Properties.ARTICLE_SUPPLIER.ToString().Equals("PROV"))
                        material.CalculoCoste = "";
                    else
                        material.CalculoCoste= "X";

                    //Fabricante
                    try
                    {
                        material.Fabricante = articleReference.Article.Properties.ARTICLE_MANUFACTURER_NAME;
                    }
                    catch
                    {
                        material.Fabricante = "";
                    }

                    try
                    {
                        //Referencia Fabricante
                        material.RefFabricante = articleReference.Properties.ARTICLE_ORDERNR.ToString();
                    }
                    catch
                    {
                        material.RefFabricante = "";
                    }


                    materials.Add(material);
                }

            }

            //Escribe en excel
            // Escribir diferencias en Excel
            string filePath = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "MAT\\Materiales_", location, ".xls");

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial; // Para uso no comercial
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Diferencias");

                // Escribir encabezado
                worksheet.Cells[1, 1].Value = "Position";
                worksheet.Cells[1, 2].Value = "Position Type";
                worksheet.Cells[1, 3].Value = "SAP Code";
                worksheet.Cells[1, 4].Value = "Description L1";
                worksheet.Cells[1, 5].Value = "Description L2";
                worksheet.Cells[1, 6].Value = "Count";
                worksheet.Cells[1, 7].Value = "Aporte";
                worksheet.Cells[1, 8].Value = "Relevancia Fab.";
                worksheet.Cells[1, 9].Value = "Calculo Coste";
                worksheet.Cells[1, 10].Value = "Nombre SAP";
                worksheet.Cells[1, 11].Value = "Fabricante";
                worksheet.Cells[1, 12].Value = "Referencia Fabricante";

                int row = 2;
                foreach (SAPMaterial material in materials)
                {
                    //Position
                    worksheet.Cells[row, 1].Value = row-1;

                    //Position Type
                    worksheet.Cells[row, 2].Value = "L";

                    //SAP Code
                    worksheet.Cells[row, 3].Value = material.SAPCode;

                    //Description L1
                    worksheet.Cells[row, 4].Value = material.SAPDescriptionL1;

                    //Description L2
                    worksheet.Cells[row, 5].Value = material.SAPDescriptionL2;

                    //Count
                    worksheet.Cells[row, 6].Value = material.Count;

                    //Aporte
                    worksheet.Cells[row, 7].Value = material.Aporte;

                    //Relevancia Fab.
                    worksheet.Cells[row, 8].Value = "X";

                    //Calculo Coste
                    worksheet.Cells[row, 9].Value = material.CalculoCoste;

                    //Nombre SAP
                    worksheet.Cells[row, 10].Value = material.SAPName;

                    //Fabricante
                    worksheet.Cells[row, 11].Value = material.Fabricante;

                    //Referencia Fabricante
                    worksheet.Cells[row, 12].Value = material.RefFabricante;

                    row++;

                }

                // Guardar archivo
                package.SaveAs(new FileInfo(filePath));

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
     * 
     * //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Position";
                oSheet.Cells[1, 2] = "Position Type";
                oSheet.Cells[1, 3] = "SAP Code";
                oSheet.Cells[1, 4] = "Description L1";
                oSheet.Cells[1, 5] = "Description L2";
                oSheet.Cells[1, 6] = "Count";
                oSheet.Cells[1, 7] = "Aporte";
                oSheet.Cells[1, 8] = "Relevancia Fab.";
                oSheet.Cells[1, 9] = "Calculo Coste";
                oSheet.Cells[1, 10] = "Nombre SAP";
                oSheet.Cells[1, 11] = "Fabricante";
                oSheet.Cells[1, 12] = "Referencia Fabricante";
     * 
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

        public string CalculoCoste { get; set; }

        public SAPMaterial() { }

    }

}
