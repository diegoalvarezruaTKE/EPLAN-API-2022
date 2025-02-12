using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;
using EPLAN_API.User;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace EPLAN_API.API_Basic
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
            draw_Main_Cab_3D();
            draw_Main_Basic_Cables();
        }

        private void draw_Main_Basic_Cables()
        {
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'A', "External Feed Wiring", 76, 256);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'B', "Safety Inputs I", 84, 156);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'C', "Motor", 20, 100);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'D', "Control I", 148, 56);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'E', "Motor Sensors", 28, 264);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'F', "Control Inputs I", 356, 124);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'G', "Communication", 148, 172);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'H', "Display", 124, 148);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'I', "External Feed Wiring", 268, 48);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'J', "Control I", 252, 56);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'K', "Safety Inputs II", 128, 168);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'L', "Safety Pulse Inputs", 96, 168);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'N', "Control II", 176, 220);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'O', "Control I", 44, 52);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'P', "Communication", 280, 172);
        }

        public void draw_Main_Cab_3D()
        {
            InstallationSpace installationSpace = new InstallationSpace();
            foreach (InstallationSpace iSpace in oProject.InstallationSpaces)
            {
                if (iSpace.ToString().Equals("1 - MAIN"))
                    installationSpace = iSpace;
            }
            StorableObject[] storableObjects = insert3DMacro(oProject, "$(MD_MACROS)\\_Esquema\\3_Basic\\1 - MAIN.ema", 'A', installationSpace, 1, 1, 1, 0.2);

            //Find Mounting Panel and Cabinet
            MountingPanel FrontMountingPanel = new MountingPanel();
            MountingPanel DoorMountingPanel = new MountingPanel();
            Cabinet Cabinet = new Cabinet();
            Cabinet Enclosure = new Cabinet();
            Plane FlangePlateInside = new Plane();
            Plane FlangePlateOutside = new Plane();
            Plane SidePanelRightOutside = new Plane();
            Placement3D[] Planearr = { };

            foreach (StorableObject StorableObject in storableObjects)
            {
                MountingPanel MP = StorableObject as MountingPanel;
                Cabinet cab = StorableObject as Cabinet;
                Plane plane = StorableObject as Plane;

                if (cab != null)
                {
                    if (cab.Parent.ToString().Equals(installationSpace.ToString()))
                        Cabinet = cab;
                    else
                        Enclosure = cab;

                }
                if (MP != null)
                {
                    if (MP.Name.Contains("MP1_MAIN"))
                        FrontMountingPanel = MP;

                    if (MP.Name.Contains("MP2_MAIN"))
                        DoorMountingPanel = MP;
                }
                if (plane != null)
                {
                    PropertyValue propertyValue = plane.Properties.FUNCTION3D_DESIGNATION;
                    PropertyValue propertyValue1 = plane.Parent.Properties.FUNCTION3D_DESIGNATION;
                    if (propertyValue.ToString().Equals("Flange plate inside"))
                        FlangePlateInside = plane;
                    if (propertyValue.ToString().Equals("Flange plate outside"))
                        FlangePlateOutside = plane;
                    if (propertyValue.ToString().Equals("Side panel right outside"))
                        SidePanelRightOutside = plane;
                    if (propertyValue1.ToString().Equals("Housings"))
                    {
                        if (!propertyValue.ToString().Equals("Side panel right outside"))
                            Planearr = Planearr.Append(plane).ToArray();
                    }

                }
            }

            Placement[] Placements = oProject.Pages[0].AllFirstLevelPlacements;
            Placements = Placements.Concat(oProject.Pages[1].AllFirstLevelPlacements).ToArray();
            Placements = Placements.Concat(oProject.Pages[2].AllFirstLevelPlacements).ToArray();
            Placements = Placements.Concat(oProject.Pages[3].AllFirstLevelPlacements).ToArray();
            Placements = Placements.Concat(oProject.Pages[4].AllFirstLevelPlacements).ToArray();

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
                        case "1":
                            viewPlacement.RootElements = new Placement3D[] { Cabinet };
                            viewPlacement.ViewDirection = ViewDirectionType.IsoSouthEast;
                            break;
                        case "2":
                            viewPlacement.RootElements = new Placement3D[] { Enclosure };
                            viewPlacement.ViewDirection = ViewDirectionType.FromTop;
                            break;
                        case "3":
                            viewPlacement.RootElements = new Placement3D[] { FrontMountingPanel };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "4":
                            viewPlacement.RootElements = new Placement3D[] { DoorMountingPanel };
                            viewPlacement.ViewDirection = ViewDirectionType.FromBack;
                            break;
                        case "5":
                            viewPlacement.RootElements = new Placement3D[] { FrontMountingPanel };
                            viewPlacement.ViewDirection = ViewDirectionType.FromFront;
                            break;
                        case "6":
                            viewPlacement.RootElements = new Placement3D[] { DoorMountingPanel };
                            viewPlacement.ViewDirection = ViewDirectionType.FromBack;
                            break;
                        case "7":
                            viewPlacement.RootElements = Planearr;
                            viewPlacement.IncludedElements = new Placement3D[] { Enclosure };
                            viewPlacement.ViewDirection = ViewDirectionType.FromTop;
                            break;
                        case "8":
                            viewPlacement.IncludedElements = new Placement3D[] { Enclosure as Placement3D, SidePanelRightOutside as Placement3D };
                            viewPlacement.ViewDirection = ViewDirectionType.FromRight;
                            break;
                        case "9":
                            viewPlacement.RootElements = new Placement3D[] { FlangePlateInside, FlangePlateOutside };
                            viewPlacement.ViewDirection = ViewDirectionType.FromBelow;
                            break;

                    }

                    //Update view
                    viewPlacement.Update();
                }
            }

        }


    }
}
