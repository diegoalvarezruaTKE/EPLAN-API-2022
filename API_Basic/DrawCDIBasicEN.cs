using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;
using EPLAN_API.User;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace EPLAN_API.API_Basic
{
    public class DrawCDIBasicEN : DrawTools
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
            draw_CDI_3D();
            draw_CDI_Basic_Cables();

            Reports report = new Reports();
            report.GenerateProject(oProject);

            //Redraw
            Edit edit = new Edit();
            edit.RedrawGed();

            ReleaseBasic releaseBasic = new ReleaseBasic(oProject, "CDI");
        }

        private void draw_CDI_Basic_Cables()
        {
            double particiones = 1;

            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'C', "Lower Diagnostic Inputs II", 20, 108);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'D', "Lower Diagnostic Inputs IV", 72, 164);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'F', "Lower Keys", 4, 244);

            if (particiones == 0)
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'I', "Lower Power Supply", 60, 176);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'J', "Lower Diagnostic Inputs I", 32, 188);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'K', "Lower Sensors I", 16, 264);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'L', "Lower Lighting I", 40, 232);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'M', "Lower Maintenance", 120, 188);
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'A', "Lower Power Supply", 40, 208);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'B', "Lower Diagnostic Inputs I", 32, 188);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'E', "Lower Sensors I", 16, 264);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'G', "Lower Lighting I", 40, 232);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359500_CABLEADO_CDI_BASIC.ema", 'H', "Lower Maintenance", 120, 188);
            }

        }


        public void draw_CDI_3D()
        {
            InstallationSpace installationSpace = new InstallationSpace();
            foreach (InstallationSpace iSpace in oProject.InstallationSpaces)
            {
                if (iSpace.ToString().Equals("3 - CDI_600"))
                    installationSpace = iSpace;
            }
            StorableObject[] storableObjects = insert3DMacro(oProject, "$(MD_MACROS)\\_Esquema\\3_Basic\\3 - CDI_600.ema", 'A', installationSpace, 1, 1, 1, 0.2);

            //Find Mounting Panel and Cabinet
            MountingPanel MountingPanel_CDI = new MountingPanel();
            Cabinet CDI_cabinet = new Cabinet();
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
                        CDI_cabinet = cab;
                    else
                        Housings = cab;
                }
                if (MP != null)
                {
                    if (MP.Name.Contains("MP1_CDI"))
                        MountingPanel_CDI = MP;

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
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Layout CDI");
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
                        case "19":
                            viewPlacement.RootElements = new Placement3D[] { CDI_cabinet };
                            viewPlacement.ViewDirection = ViewDirectionType.IsoSouthEast;
                            break;
                        case "20":
                            viewPlacement.RootElements = new Placement3D[] { CoverOutside };
                            viewPlacement.IncludedElements = new Placement3D[] { Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromTop;
                            break;
                        case "21":
                            viewPlacement.RootElements = new Placement3D[] { MountingPanel_CDI };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "22":
                            viewPlacement.RootElements = new Placement3D[] { MountingPanel_CDI };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "23":
                            viewPlacement.RootElements = new Placement3D[] { CoverOutside };
                            viewPlacement.IncludedElements = new Placement3D[] { Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromTop;
                            break;
                        case "24":
                            viewPlacement.IncludedElements = new Placement3D[] { SidePanelLeftOutside, Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromLeft;
                            break;
                        case "25":
                            viewPlacement.IncludedElements = new Placement3D[] { SidePanelRightOutside, Housings };
                            viewPlacement.ViewDirection = ViewDirectionType.FromRight;
                            break;
                        case "26":
                            viewPlacement.RootElements = new Placement3D[] { MR_U7 };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "27":
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
