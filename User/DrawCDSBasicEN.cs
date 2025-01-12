using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API.User
{
    public class DrawCDSBasicEN:DrawTools
    {
        Project oProject;

        public DrawCDSBasicEN()
        {

            Restore oRestore = new Restore();
            PathInfo EPLANPaths = new PathInfo();
            StringCollection oArchives = new StringCollection();
            string[] fileTemplatePaths = Directory.GetFiles(EPLANPaths.Templates);
            string ProjectPath;

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select a folder";
                folderDialog.ShowNewFolderButton = true;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    ProjectPath = folderDialog.SelectedPath;
                }
                else
                {
                    return;
                }
            }

            //Restore base project
            foreach (string file in fileTemplatePaths)
            {
                if (file.Contains("GEC_EN115_Base_Basic"))
                {
                    oArchives.Add(file);
                    oRestore.Project(oArchives,
                                    ProjectPath,
                                    "11200004366200_ENSAMBLAJE_CAJA_DERIV_SUP_BASIC",
                                    false,
                                    false);
                    break;
                }
            }


            oProject = new ProjectManager().OpenProject(string.Concat(ProjectPath, "\\11200004366200_ENSAMBLAJE_CAJA_DERIV_SUP_BASIC"));

            //Draw Basic Macros
            Insert oInsert = new Insert();
            PageMacro oPageMacro = new PageMacro();

            //Main Cabinet
            oPageMacro.Open("$(MD_MACROS)\\_Esquema\\3_Basic\\200004271100-CAJA DERIV SUP. BASIC GEC EN_ES_EXT.emp", oProject);
            oInsert.PageMacro(oPageMacro, oProject, null, PageMacro.Enums.NumerationMode.Ignore);
            //draw_Main_Cab_3D();
            draw_CDS_Basic_Cables();
        }

        public void draw_CDS_Basic_Cables()
        {
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'A', "Upper Power Supply", 76, 252);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'B', "Upper Diagnostic Inputs I", 4, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'C', "Upper Diagnostic Inputs II", 0, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'D', "Upper Diagnostic Inputs III", 168, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'E', "Upper Diagnostic Inputs IV", 68, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'F', "Upper Sensors I", 24, 244);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'G', "Motor Sensor I", 188, 212);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'H', "Brake I", 64, 176);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'I', "Brake II", 64, 176);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'J', "Upper Keys", 32, 144);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'K', "Upper Lighting I", 32, 224);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'L', "Upper Maintenance", 52, 272);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'M', "Interconnection Terminals", 28, 212);

        }
    }
}
