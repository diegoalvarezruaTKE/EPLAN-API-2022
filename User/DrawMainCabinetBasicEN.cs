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
    public class DrawMainCabinetBasicEN:DrawTools
    {
        Project oProject;

        public DrawMainCabinetBasicEN()
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
                                    "11200004366300_ENSAMBLAJE_ARMARIO_INT_BASIC",
                                    false,
                                    false);
                    break;
                }
            }


            oProject = new ProjectManager().OpenProject(string.Concat(ProjectPath, "\\11200004366300_ENSAMBLAJE_ARMARIO_INT_BASIC"));

            //Draw Basic Macros
            Insert oInsert = new Insert();
            PageMacro oPageMacro = new PageMacro();

            //Main Cabinet
            oPageMacro.Open("$(MD_MACROS)\\_Esquema\\3_Basic\\200004271300-ARMARIO INTERIOR BASIC GEC EN_ES.emp", oProject);
            oInsert.PageMacro(oPageMacro, oProject, null, PageMacro.Enums.NumerationMode.Ignore);
            //draw_Main_Cab_3D();
            draw_Main_Cab_Cables();
        }

        public void draw_Main_Cab_Cables()
        {
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'A', "External Feed Wiring", 76, 256);
            deleteArea(oProject, "VVF Power", 84, 216, 108, 216);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'B', "Safety Inputs I", 164, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'C', "Motor", 20, 100);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'D', "Control I", 148, 56);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'E', "Motor Sensors", 28, 264);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'F', "Control Inputs I", 356, 124);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'G', "Communication", 148, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'H', "Display", 124, 148);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'I', "Safety Inputs I", 228, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'J', "Control I", 252, 56);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'K', "Safety Inputs II", 128, 168);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'L', "Safety Pulse Inputs", 96, 168);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'M', "Safety Inputs I", 356, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'N', "Control II", 176, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'O', "Control I", 44, 52);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'P', "Communication", 280, 172);

        }
    }
}
