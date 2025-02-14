using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;
using EPLAN_API.User;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;

using System.Windows.Forms;

namespace EPLAN_API.API_Basic
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
            draw_CDS_3D();
            draw_CDS_Basic_Cables();

            Reports report = new Reports();
            report.GenerateProject(oProject);

            //Redraw
            Edit edit = new Edit();
            edit.RedrawGed();

            ReleaseBasic releaseBasic = new ReleaseBasic(oProject, "CDS");
        }

        private void draw_CDS_Basic_Cables()
        {
            double particiones = 1;

            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'B', "Upper Diagnostic Inputs I", 4, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'C', "Upper Diagnostic Inputs II", 0, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'E', "Upper Diagnostic Inputs IV", 68, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'F', "Upper Sensors I", 24, 244);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'G', "Motor Sensor I", 60, 212);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'H', "Brake I", 64, 176);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'I', "Brake II", 64, 176);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'J', "Upper Keys", 32, 144);

            if (particiones == 0)
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'N', "Upper Power Supply", 76, 252);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'O', "Upper Diagnostic Inputs III", 168, 220);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'P', "Upper Lighting I", 32, 224);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'Q', "Upper Maintenance", 52, 272);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'R', "Interconnection Terminals", 28, 212);
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'A', "Upper Power Supply", 76, 252);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'D', "Upper Diagnostic Inputs III", 168, 220);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'K', "Upper Lighting I", 32, 224);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'L', "Upper Maintenance", 52, 272);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359400_CABLEADO_CDS_BASIC.ema", 'M', "Interconnection Terminals", 28, 212);
            }

        }


        public void draw_CDS_3D()
        {
            InstallationSpace installationSpace = new InstallationSpace();
            foreach (InstallationSpace iSpace in oProject.InstallationSpaces)
            {
                if (iSpace.ToString().Equals("2 - CDS_600"))
                    installationSpace = iSpace;
            }
            StorableObject[] storableObjects = insert3DMacro(oProject, "$(MD_MACROS)\\_Esquema\\3_Basic\\2 - CDS_600.ema", 'A', installationSpace, 1, 1, 1, 0.2);

            //Find Mounting Panel and Cabinet
            MountingPanel MountingPanel_CDS = new MountingPanel();
            Cabinet CDS_cabinet = new Cabinet();
            Cabinet Housings = new Cabinet();
            Plane CoverOutside = new Plane();
            Plane SidePanelLeftOutside = new Plane();
            Plane SidePanelRightOutside = new Plane();
            MountingRail MR_U7 = new MountingRail();
            MountingRail MR_U8 = new MountingRail();

            foreach (StorableObject StorableObject in storableObjects)
            {
                MountingPanel MP = StorableObject as MountingPanel;
                Cabinet cab = StorableObject as Cabinet;
                Plane plane = StorableObject as Plane;
                MountingRail MR = StorableObject as MountingRail;

                if (cab != null)
                {
                    if (cab.Parent.ToString().Equals(installationSpace.ToString()))
                        CDS_cabinet = cab;
                    else
                        Housings = cab;
                }
                if (MP != null)
                {
                    if (MP.Name.Contains("MP1_CDS"))
                        MountingPanel_CDS = MP;

                }
                if (plane != null)
                {
                    PropertyValue propertyValue = plane.Properties.FUNCTION3D_DESIGNATION;
                    PropertyValue propertyValue1 = plane.Parent.Properties.FUNCTION3D_DESIGNATION;
                    if (propertyValue.ToString().Equals("Cover outside"))
                    {
                        CoverOutside = plane;
                        //Housings = plane.Parent as Cabinet;
                    }
                    if (propertyValue.ToString().Equals("Side panel left outside"))
                    {
                        SidePanelLeftOutside = plane;
                        //Housings = plane.Parent as Cabinet;
                    }
                    if (propertyValue.ToString().Equals("Side panel right outside"))
                    {
                        SidePanelRightOutside = plane;
                        //Housings = plane.Parent as Cabinet;
                    }

                }
                if (MR != null)
                {
                    PropertyValue propertyValue = MR.Properties.FUNCTION3D_DESIGNATION;
                    if (MR.Name.Contains("U7"))
                        MR_U7 = MR;
                    if (MR.Name.Contains("U8"))
                        MR_U8 = MR;
                }
            }

            Dictionary<int, string> dictPages = GetPageTable(oProject);
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Layout CDS");
            Placement[] Placements = oProject.Pages[key].AllFirstLevelPlacements;
            Placements = Placements.Concat(oProject.Pages[key + 1].AllFirstLevelPlacements).ToArray();
            Placements = Placements.Concat(oProject.Pages[key + 2].AllFirstLevelPlacements).ToArray();
            Placements = Placements.Concat(oProject.Pages[key + 3].AllFirstLevelPlacements).ToArray();
            Placements = Placements.Concat(oProject.Pages[key + 4].AllFirstLevelPlacements).ToArray();

            foreach (Placement Placement in Placements)
            {
                ViewPlacement viewPlacement = Placement as ViewPlacement;
                if (viewPlacement != null)
                {
                    PropertyValue propertyValue = viewPlacement.Properties.DMG_VIEWPLACEMENT_DESIGNATION;
                    viewPlacement.InstallationSpace = installationSpace;

                    //Assign root element
                    switch (propertyValue.ToString())
                    {
                        case "10":
                            viewPlacement.RootElements = new Placement3D[] { CDS_cabinet };
                            viewPlacement.ViewDirection = ViewDirectionType.IsoSouthEast;
                            break;
                        case "11":
                            viewPlacement.RootElements = new Placement3D[] { MountingPanel_CDS };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "12":
                            viewPlacement.IncludedElements = new Placement3D[] { CoverOutside, Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromTop;
                            break;
                        case "13":
                            viewPlacement.RootElements = new Placement3D[] { MountingPanel_CDS };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "14":
                            viewPlacement.IncludedElements = new Placement3D[] { CoverOutside, Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromTop;
                            break;
                        case "15":
                            viewPlacement.IncludedElements = new Placement3D[] { SidePanelLeftOutside, Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromLeft;
                            break;
                        case "16":
                            viewPlacement.IncludedElements = new Placement3D[] { SidePanelRightOutside, Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromRight;
                            break;
                        case "17":
                            viewPlacement.RootElements = new Placement3D[] { MR_U7 };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "18":
                            viewPlacement.RootElements = new Placement3D[] { MR_U8 };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;

                    }

                    //Update view
                    viewPlacement.Update();
                }
            }


        }

    }
}
