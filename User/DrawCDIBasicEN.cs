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
    public class DrawCDIBasicEN:DrawTools
    {
        Project oProject;

        public DrawCDIBasicEN()
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
                                    "11200004366100_ENSAMBLAJE_CAJA_DERIV_INF_BASIC",
                                    false,
                                    false);
                    break;
                }
            }


            oProject = new ProjectManager().OpenProject(string.Concat(ProjectPath, "\\11200004366100_ENSAMBLAJE_CAJA_DERIV_INF_BASIC"));

            //Draw Basic Macros
            Insert oInsert = new Insert();
            PageMacro oPageMacro = new PageMacro();

            //Main Cabinet
            oPageMacro.Open("$(MD_MACROS)\\_Esquema\\3_Basic\\200004271200-CAJA DERIV INF. BASIC GEC EN_ES_EXT.emp", oProject);
            oInsert.PageMacro(oPageMacro, oProject, null, PageMacro.Enums.NumerationMode.Ignore);
            //draw_Main_Cab_3D();
            draw_CDI_Basic_Cables();
        }

        public void draw_CDI_Basic_Cables()
        {
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'A', "Lower Power Supply", 40, 208);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'B', "Lower Diagnostic Inputs I", 32, 188);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'C', "Lower Diagnostic Inputs II", 20, 108);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'D', "Lower Diagnostic Inputs IV", 72, 164);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'E', "Lower Sensors I", 16, 264);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'F', "Lower Keys", 4, 244);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'G', "Lower Lighting I", 40, 232);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'H', "Lower Maintenance", 120, 188);

        }
    }
}
