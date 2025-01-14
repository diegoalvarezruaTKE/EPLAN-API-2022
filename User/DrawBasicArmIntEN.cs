using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;
using Eplan.EplApi.DataModel.EObjects;
using System.Globalization;
using Eplan.EplApi.DataModel.E3D;
using EPLAN_API.SAP;
//using System.Windows.Media.Media3D;
//using static System.Windows.Forms.VisualStyles.VisualStyleElement;
//using System.Runtime.InteropServices.ComTypes;


namespace EPLAN_API.User
{
    public class DrawBasicArmIntEN:DrawTools
    {
        private Project oProject;
        private Page[] oPages;
        private Function[] oCablesFunctions;
        private Cable[] oCables;
        private Hashtable oHPages;
        Dictionary<int, string> dictPages;
        private Cable[] sCables;
        private Electric oElectric;
        private DrawingService oDs;
        private string log;
        private ProgressBar oProgressBar;
        private String OE;
        StorableObject[] oInsertedObjects;
        private TextBox stateTexbox;

        public DrawBasicArmIntEN(Project project, Electric electric)
        {
            oProject = project;
            oElectric = electric;
            DrawMacro();
        }

        public void DrawMacro()
        {

            //Draw Basic Macros
            Insert oInsert = new Insert();
            PageMacro oPageMacro = new PageMacro();

            /*********************/
            /****MAIN CABINET*****/
            /*********************/

            //Main Cabinet
            setStatusText("Insertando macro Armario Basic");
            oPageMacro.Open("$(MD_MACROS)\\_Esquema\\3_Basic\\200004271300-ARMARIO INTERIOR BASIC GEC EN_ES.emp", oProject);
            oInsertedObjects = oInsert.PageMacro(oPageMacro, oProject, null, PageMacro.Enums.NumerationMode.Ignore);
            //draw_Main_Cab_3D();
            draw_Main_Basic_Cables();

            /*********************/
            /*********CDS*********/
            /*********************/
            //Upper Box
            setStatusText("Insertando macro Caja Derivación Superior Basic");
            oPageMacro.Open("$(MD_MACROS)\\_Esquema\\3_Basic\\200004271100-CAJA DERIV SUP. BASIC GEC EN_ES_EXT.emp", oProject);
            oInsertedObjects.Concat(oInsert.PageMacro(oPageMacro, oProject, null, PageMacro.Enums.NumerationMode.Ignore));
            //draw_CDS_3D();
            draw_CDS_Basic_Cables();

            /*********************/
            /*********CDI*********/
            /*********************/
            //Lower Box
            setStatusText("Insertando macro Caja Derivación Inferior Basic");
            oPageMacro.Open("$(MD_MACROS)\\_Esquema\\3_Basic\\200004271200-CAJA DERIV INF. BASIC GEC EN_ES_EXT.emp", oProject);
            oInsertedObjects.Concat(oInsert.PageMacro(oPageMacro, oProject, null, PageMacro.Enums.NumerationMode.Ignore));
            //draw_CDI_3D();
            draw_CDI_Basic_Cables();

            /*********************/
            /******ADVANCED*******/
            /*********************/
            //11200004299200  RELÉ TÉRMICO MOTOR_BAS
            draw_termico();

            //11200004443200  CABLEADO VARIADOR ADV.
            draw_VVF();

            //11200004443300  MAX INSTALADO ADV.
            draw_MAX();

            //11200004443400	PAQUETE MERCADONA ADV.


            //1	GENERAL
            //1.6.1 Stop para Carritos
            //11200004441900  STOP CARRITOS SUP.ADV.
            //11200004444100  STOP CARRITOS INF.ADV.
            draw_stop_carritos();


            //9	SAFETY DEVICES

            //9.2  Sensor de sincronismo de pasamanos
            //11200004441700  SINCRONISMO PASAMANOS ADV.
            draw_Speed_Handrail_Monitoring();

            //9.4	Seguridad funcionamiento de frenos
            //11200004323000  FRENO 1 MOTOR 1 - FINAL DE CARRERA_ADV
            //11200004323200  FRENO 1 MOTOR 1 - INDUCTIVO_ADV
            draw_Freno_1_M1();
            //11200004323100  FRENO 2 MOTOR 1 - FINAL DE CARRERA_ADV
            //11200004323300  FRENO 2 MOTOR 1 - INDUCTIVO_ADV
            draw_Freno_2_M1();

            //9.26	Traffic lights
            //11200004441600  SEMAFORO SUP. ADV.
            //11200004443900  SEMAFORO INF. ADV.
            draw_Traffic_Lights();

            //9.31	Cerrojo mantenimiento en eje 
            //11200004441800  CERROJO MANTENIMIENTO ADV
            draw_Main_shaft_maintenance_lock();

            //9.5	Brake wear indicator
            //11200004442000  DESGASTE FRENOS ADV.
            draw_Brake_wear_assembly();

            //9.14  Comb plate safety devices
            //11200004442100  SEG.VERT.PEINES CL. SUP.ADV.
            //11200004442200  SEG.VERT.PEINES SUP. ADV.
            //11200004444400  SEG.VERT.PEINES CL. INF.ADV.
            //11200004444300  SEG.VERT.PEINES INF. ADV.
            draw_Vertical_Combplate();

            //9.20	Buggy device
            //11200004442300  BUGGY CLASSIC SUP.ADV.
            //11200004442400  BUGGY SUP. ADV.
            //11200004444500  BUGGY CLASSIC INF.ADV.
            //11200004444600  BUGGY INF. ADV.
            draw_Buggy();

            //9.21	Drive chain safety devices
            //11200004442900  CADENA MOTRIZ ADV.
            //draw_Drive_Chain_Safety();

            //9.3   Handrail breakage sensor
            //11200004444200  ROTURA PASAMANOS ADV.
            draw_Handrail_breakage_sensor();

            //9.30	Level of water on pit detector
            //11200004444000  NIVEL DE AGUA ADV.
            draw_Level_Water();


            //10    DRIVE
            //10.2.2	Tipo detección - Fotocélula
            //11200004442500  FOTOCELULA CL. SUP.ADV.
            //11200004442600  FOTOCELULA SUP. ADV.
            //11200004444700  FOTOCELULA CL. INF.ADV.
            //11200004444800  FOTOCELULA INF. ADV.
            draw_fotocélula_VVF();
            //11200004442700  RADAR CL. SUP.ADV.
            //11200004442800  RADAR SUP. ADV.
            //11200004444900  RADAR CL. INF.ADV.
            //11200004445000  RADAR INF. ADV.
            draw_radar();

            //10.11	Automatic lubrication
            //draw_automatic_lubrication();


            //11200004445200  MULTICAB.INF.CENT.ADV PARTIDO





            deleteAllDummyConnections(oProject);
            return;

            /****ADVANCED*****/
            
            

            //9	SAFETY DEVICES

            

            

            

            

            //9.22	Skirting microswitches
            //9.23	Fire/smoke detectors
            //9.25	Additional user stop (both heads)     

            //9.27	Failure display -- Always included
            //9.28	Passenger detection system
            //9.29	Eject device for the controller

            

            //9.32	Temperature sensor
            //9.33	LHD (linear heat detection)
            //9.34  Oil lever switch in gear
            //9.35  Fire contact in controller
            //9.36	Sismic contact in controller
            //9.37	People counter
            //9.38	Run time meter
            //9.43	Cables in conduits

            

            //10.13 Auxiliary brake on the main shaft
            draw_auxiliary_brake();

            //Packages
            draw_Paquete_Mercadona();


            //Get BOMs
            //getBOM("MAIN");
            //getBOM("CDS");
            //getBOM("CDI");
            //calcParmGEC_Basic();
            paramGEC();


            Reports report = new Reports();
            report.GenerateProject(oProject);

            //Redraw
            Edit edit = new Edit();
            edit.RedrawGed();

            setStatusText("Finalizado");
            oProgressBar.Style = ProgressBarStyle.Continuous;
            oProgressBar.MarqueeAnimationSpeed = 0;

        }

        public void draw_Main_Cab_3D()
        {
            InstallationSpace installationSpace = new InstallationSpace();
            foreach (InstallationSpace iSpace in oProject.InstallationSpaces)
            {
                if (iSpace.ToString().Equals("1 - MAIN"))
                    installationSpace = iSpace;
            }
            StorableObject[] storableObjects = insert3DMacro(oProject,"$(MD_MACROS)\\_Esquema\\3_Basic\\1 - MAIN.ema", 'A', installationSpace, 1, 1, 1, 0.2);

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

        public void draw_Central_Cables()
        {

            if (((Caracteristic)oElectric.CaractComercial["FEFL0"]).NumVal <= 0)
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'A', "Upper Power Supply", 172, 160);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'B', "Upper Diagnostic Inputs III", 160, 104);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'D', "Lower Power Supply", 60, 188);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'E', "Lower Diagnostic Inputs I", 24, 192);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'F', "Lower Sensors I", 56, 96);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'G', "Lower Maintenance", 120, 248);

                //Configure cable length
                updateCableLengthProperty("Cable.length.W1", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDS");
                updateCableLengthProperty("Cable.length.W1", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDI");
                updateCableLengthProperty("Cable.length.W100", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDS");
                updateCableLengthProperty("Cable.length.WC1", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDS");
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'I', "Upper Power Supply", 172, 160);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'O', "Upper Diagnostic Inputs III", 160, 104);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'L', "Lower Power Supply", 60, 188);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'M', "Lower Diagnostic Inputs I", 24, 184);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'K', "Lower Sensors I", 56, 124);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'J', "Lower Maintenance", 120, 216);

                //Change cable gland name in upper and lower boxes
                FunctionsFilter ff = new FunctionsFilter();
                ff.Category = Function.Enums.Category.CabMechConduit;
                FunctionBasePropertyList functionBasePropertyList = new FunctionBasePropertyList();
                /*No funciona el filtrado
                //functionBasePropertyList.DESIGNATION_LOCATION.Set("CDI");
                //ff.SetFilteredPropertyList(functionBasePropertyList);
                */
                DMObjectsFinder objFinder = new DMObjectsFinder(oProject);

                Function[] ofunctions = objFinder.GetFunctions(ff);

                foreach (Function f in ofunctions)
                {
                    String location = f.NameParts.DESIGNATION_LOCATION.ToString();
                    String name = f.VisibleName;

                    if (name.Equals("W1"))
                    {
                        if (location.Equals("CDS"))
                        {
                            f.Name = "=ESC+CDS-W1.U";
                            f.VisibleName = "W1.U";
                        }
                        if (location.Equals("CDI"))
                        {
                            f.Name = "=ESC+CDI-W1.L";
                            f.VisibleName = "W1.L";
                        }
                    }
                }

                //Configure cable length
                updateCableLengthProperty("Cable.length.W1.U", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDS");
                updateCableLengthProperty("Cable.length.W1.L", 2.1, "CDI");
                updateCableLengthProperty("Cable.length.W100.U", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDS");
                updateCableLengthProperty("Cable.length.W100.L", 2.1, "CDI");
                updateCableLengthProperty("Cable.length.WC1.U", ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal, "CDS");
                updateCableLengthProperty("Cable.length.WC1.L", 2.1, "CDI");

            }

            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004028900_00_GEC_CENTRAL_CABLES.ema", 'C', "Upper Maintenance", 80, 232);

        }

        public void draw_Main_Basic_Cables()
        {
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004359300_CABLEADO_ARMARIO_BASIC.ema", 'A', "External Feed Wiring", 76, 256);
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

        public void draw_termico()
        {
            setStatusText("Insertando Termico");

            //Delete XAUX
            deleteArea(oProject, "Motor", 28, 152, 68, 152);

            Caracteristic iTermico = ((Caracteristic)oElectric.CaractIng["ITERMICO"]);

            StorableObject[] storableObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004299200 -  RELÉ TÉRMICO MOTOR_BAS.ema", 'A', "Motor", 52, 156);

            foreach (Placement placement in storableObjects)
            {
                PlaceHolder oPlaceHolder = placement as PlaceHolder;
                if (oPlaceHolder != null)
                {
                    if (oPlaceHolder.Name == "Termico")
                    {
                        if (iTermico.NumVal <= 7)
                        {
                            oPlaceHolder.ApplyRecord("0<I<7");
                        }
                        if (iTermico.NumVal > 7 && iTermico.NumVal <= 9)
                        {
                            oPlaceHolder.ApplyRecord("7<I<9");
                        }
                        if (iTermico.NumVal > 9 && iTermico.NumVal <= 13)
                        {
                            oPlaceHolder.ApplyRecord("9<I<13");
                        }
                        else if (iTermico.NumVal > 13 && iTermico.NumVal <= 18.0)
                        {
                            oPlaceHolder.ApplyRecord("13<=I<18");
                        }
                        else if (iTermico.NumVal > 18.0 && iTermico.NumVal <= 25)
                        {
                            oPlaceHolder.ApplyRecord("18<=I<25");
                        }
                        else if (iTermico.NumVal > 25.0 && iTermico.NumVal <= 32)
                        {
                            oPlaceHolder.ApplyRecord("25<=I<32");
                        }
                        else if (iTermico.NumVal > 32.0 && iTermico.NumVal <= 40)
                        {
                            oPlaceHolder.ApplyRecord("32<=I<40");
                        }
                        else if (iTermico.NumVal > 40.0 && iTermico.NumVal <= 50)
                        {
                            oPlaceHolder.ApplyRecord("40<=I<50");
                        }
                        break;
                    }
                }
            }
        }

        public void draw_VVF()
        {
            if (((Caracteristic)oElectric.CaractComercial["TNCR_SD_SIST_AHORRO"]).CurrentReference.Equals("VA"))
            {
                deleteArea(oProject, "VVF Power", 84, 216, 108, 216);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004443200 - CABLEADO VARIADOR ADV.ema", 'A', "VVF Power", 36, 220);
                deleteArea(oProject, "VVF Control", 132, 216, 132, 216);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004443200 - CABLEADO VARIADOR ADV.ema", 'B', "VVF Control", 48, 212);

                double cableLength = ((Caracteristic)oElectric.CaractComercial["TNCR_OT_DESARROLLO"]).NumVal + 5;
                SetCableLength(oProject, "VVF Power", "WP20", cableLength);
                SetCableLength(oProject, "VVF Power", "WP21", cableLength);
                SetCableLength(oProject, "VVF Control", "W106", cableLength);

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "I5", (uint)GEC.Param.VFD_EEC, true);
            }
        }

        public void draw_MAX()
        {
            if (((Caracteristic)oElectric.CaractComercial["TNCR_S_MAX"]).CurrentReference.Equals("A"))
            {
                setStatusText("Insertando MAX");
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004443300 - MAX INSTALADO ADV.ema", 'A', "Display", 124, 148);
            }
        }

        public void draw_Freno_1_M1()
        {

            if (((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference.Equals("FJ") ||
                ((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference.Equals("FTJ+FJ") ||
                ((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference.Equals("QC"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004323000 - FRENO 1 MOTOR 1 - FINAL DE CARRERA_ADV.ema", 'A', "Brake I", 64, 76);
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004323200 - FRENO 1 MOTOR 1 - INDUCTIVO_ADV.ema", 'A', "Brake I", 64, 76);
            }
        }

        public void draw_Freno_2_M1()
        {

            if (((Caracteristic)oElectric.CaractComercial["FBREMSE2"]).CurrentReference.Equals("4/4"))
            {
                if (((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference.Equals("FJ") ||
                    ((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference.Equals("FTJ+FJ") ||
                    ((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference.Equals("QC"))
                {
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004323100 - FRENO 2 MOTOR 1 - FINAL DE CARRERA_ADV.ema", 'A', "Brake II", 64, 76);
                }
                else
                {
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004323300 - FRENO 2 MOTOR 1 - INDUCTIVO_ADV.ema", 'A', "Brake II", 64, 76);
                }

                SetGECParameter(oProject, oElectric, "SI15", (uint)GEC.Param.Brake_function_brake_3_mot_1, true);
                SetGECParameter(oProject, oElectric, "SI16", (uint)GEC.Param.Brake_function_brake_4_mot_1, true);
            }
        }

        public void draw_stop_carritos()
        {
            if (!((Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"]).CurrentReference.Equals("KEINE"))
            {
                setStatusText("Insertando Stop de carritos");

                //Insert graphical macro
                //Upper box
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004441900 - STOP CARRITOS SUP.ADV.ema", 'A', "Upper Diagnostic Inputs II", 168, 112);
                //Lower box
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444100 - STOP CARRITOS INF.ADV.ema", 'A', "Lower Diagnostic Inputs III", 164, 112);

                //Configure GEC parameters
                //Upper box
                SetGECParameter(oProject, oElectric, "UI11", (uint)GEC.Param.Top_emergency_stop_trolley_SS, true);
                //Lower box
                SetGECParameter(oProject, oElectric, "LI18", (uint)GEC.Param.Bottom_emergency_stop_trolley_SS, true);
            }
        }

        public void draw_Speed_Handrail_Monitoring()
        {
            setStatusText("Insertando control velocidad pasamanos");

            if (((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
            {
                /*CABEZA INFERIOR*/

                //Insert graphical macro
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004307800 - Seguridad sincro. de pasamanos advance_ADV.ema", 'B', "Lower Sensors I", 256, 172);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004307800 - Seguridad sincro. de pasamanos advance_ADV.ema", 'C', "Upper Sensors I", 256, 204);
            }
            else
            {
                /*CABEZA SUPERIOR*/

                //Insert graphical macro
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004307800 - Seguridad sincro. de pasamanos advance_ADV.ema", 'A', "Upper Sensors I", 256, 140);
            }

        }

        public void draw_Handrail_breakage_sensor()
        {

            if (((Caracteristic)oElectric.CaractComercial["F09ZUB1"]).CurrentReference.Equals("BRUCHSCHALT"))
            {
                setStatusText("Insertando Rotura Pasamanos");
                //Insert graphical macro
                insertNewPage(oProject,"Lower Sensors II", "Lower Sensors I");
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444200 - ROTURA PASAMANOS ADV.ema", 'A', "Lower Sensors II", 20, 260);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444200 - ROTURA PASAMANOS ADV.ema", 'B', "Lower Diagnostic Inputs IV", 144, 164);

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "LI25", (uint)GEC.Param.Broken_handrail_L, true);
                SetGECParameter(oProject, oElectric, "LI26", (uint)GEC.Param.Broken_handrail_R, true);
            }

        }

        public void draw_Brake_wear_assembly()
        {
            if (((Caracteristic)oElectric.CaractComercial["F01ZUB"]).CurrentReference.Equals("INDUCTIVO"))
            {
                setStatusText("Insertando Desgaste Freno");

                //Insert graphical macro
                insertNewPage(oProject, "Upper Sensors II", "Upper Sensors I");
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442000 - DESGASTE FRENOS ADV.ema", 'A', "Upper Sensors II", 20, 264);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442000 - DESGASTE FRENOS ADV.ema", 'B', "Upper Diagnostic Inputs IV", 212, 172);

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "UI27", (uint)GEC.Param.Brake_wear_brake_1_M1, true);
                SetGECParameter(oProject, oElectric, "UI28", (uint)GEC.Param.Brake_wear_brake_2_M1, true);
            }

        }
        
        public void draw_Vertical_Combplate()
        {
            if (((Caracteristic)oElectric.CaractComercial["FKAMMPLHK"]).CurrentReference.Equals("INDEPENDIENTE"))
            {
                setStatusText("Insertando Seguridad Vertical de Peines");

                if (((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                {
                    //Insert graphical macro
                    //Upper box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442100 - SEG.VERT.PEINES CL. SUP.ADV.ema", 'A', "Upper Diagnostic Inputs II", 224, 104);
                    //Lower box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444400 - SEG.VERT.PEINES CL. INF.ADV.ema", 'A', "Lower Diagnostic Inputs II", 288, 112);
                }
                else
                {
                    //Insert graphical macro
                    //Upper box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442200 - SEG.VERT.PEINES SUP. ADV.ema", 'A', "Upper Diagnostic Inputs II", 224, 104);
                    //Lower box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444300 - SEG.VERT.PEINES INF. ADV.ema", 'A', "Lower Diagnostic Inputs II", 288, 112);
                }

                //Configure GEC parameters
                //Upper box
                SetGECParameter(oProject, oElectric, "UI12", (uint)GEC.Param.Top_vertical_comb_plate_right_SS, true);
                SetGECParameter(oProject, oElectric, "UI13", (uint)GEC.Param.Top_vertical_comb_plate_left_SS, true);
                ////Lower box
                SetGECParameter(oProject, oElectric, "LI13", (uint)GEC.Param.Bottom_vertical_comb_plate_right_SS, true);
                SetGECParameter(oProject, oElectric, "LI14", (uint)GEC.Param.Bottom_vertical_comb_plate_left_SS, true);

            }
        }

        public void draw_Buggy()
        {
            if (((Caracteristic)oElectric.CaractComercial["F04ZUB"]).CurrentReference.Equals("BUGGY") ||
                ((Caracteristic)oElectric.CaractComercial["F04ZUB"]).CurrentReference.Equals("BUGGYOT"))
            {

                setStatusText("Insertando Buggy");

                if (((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                {
                    //Insert graphical macro
                    //Upper box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442300 - BUGGY CLASSIC SUP.ADV.ema", 'A', "Upper Diagnostic Inputs II", 348, 112);
                    deleteSymbol(oProject, "Upper Diagnostic Inputs III", "=ESC-SL", "InterruptionPoint");
                    deleteSymbol(oProject, "Upper Diagnostic Inputs III", "=ESC-SL", "SymbolReference", 20, 112);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442300 - BUGGY CLASSIC SUP.ADV.ema", 'B', "Upper Diagnostic Inputs III", 20, 108);
                    //Lower box
                    deleteSymbol(oProject, "Lower Diagnostic Inputs III", "=ESC-SL", "InterruptionPoint");
                    deleteSymbol(oProject, "Lower Diagnostic Inputs III", "=ESC-SL", "SymbolReference", 20, 112);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444500 - BUGGY CLASSIC INF.ADV.ema", 'A', "Lower Diagnostic Inputs III", 20, 112);
                }
                else
                {
                    //Insert graphical macro
                    //Upper box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442400 - BUGGY SUP. ADV.ema", 'A', "Upper Diagnostic Inputs II", 348, 112);
                    deleteSymbol(oProject, "Upper Diagnostic Inputs III", "=ESC-SL", "InterruptionPoint");
                    deleteSymbol(oProject, "Upper Diagnostic Inputs III", "=ESC-SL", "SymbolReference", 20, 112);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442400 - BUGGY SUP. ADV.ema", 'B', "Upper Diagnostic Inputs III", 20, 108);
                    //Lower box
                    deleteSymbol(oProject, "Lower Diagnostic Inputs III", "=ESC-SL", "InterruptionPoint");
                    deleteSymbol(oProject, "Lower Diagnostic Inputs III", "=ESC-SL", "SymbolReference", 20, 112);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444600 - BUGGY INF. ADV.ema", 'A', "Lower Diagnostic Inputs III", 20, 112);
                }
                //Configure GEC parameters
                ////Upper box
                SetGECParameter(oProject, oElectric, "UI14", (uint)GEC.Param.Top_buggy_right_SS, true);
                SetGECParameter(oProject, oElectric, "UI15", (uint)GEC.Param.Top_buggy_left_SS, true);
                ////Lower box
                SetGECParameter(oProject, oElectric, "LI15", (uint)GEC.Param.Bottom_buggy_right_SS, true);
                SetGECParameter(oProject, oElectric, "LI16", (uint)GEC.Param.Bottom_buggy_left_SS, true);
            }
        }

        public void draw_Drive_Chain_Safety()
        {
            if (((Caracteristic)oElectric.CaractComercial["TNCR_S_DRIVE_CHAIN"]).CurrentReference.Equals("SI"))
            {
                setStatusText("Insertando Seguridad Cadena Principal");
                ;
            }
        }

        public void draw_Traffic_Lights()
        {
            if (((Caracteristic)oElectric.CaractComercial["FAMPELSYM"]).CurrentReference.Equals("BICOLOR"))
            {
                setStatusText("Insertando Semáforo Bicolor");
                insertNewPage(oProject, "Upper Traffic Lights & Oil Pump", "Upper People Detection");
                insertNewPage(oProject, "Lower Traffic Lights", "Lower People Detection");

                //Upper box
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004441600 - SEMAFORO SUP. ADV.ema", 'C', "Upper Diagnostic Outputs I", 24, 136);
                if (((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004441600 - SEMAFORO SUP. ADV.ema", 'B', "Upper Traffic Lights & Oil Pump", 48, 140);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004441600 - SEMAFORO SUP. ADV.ema", 'A', "Upper Traffic Lights & Oil Pump", 48, 140);

                //Lower box
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004443900 - SEMAFORO INF. ADV.ema", 'C', "Lower Diagnostic Outputs I", 36, 136);
                if (((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004443900 - SEMAFORO INF. ADV.ema", 'B', "Lower Traffic Lights", 48, 140);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004443900 - SEMAFORO INF. ADV.ema", 'A', "Lower Traffic Lights", 48, 140);

                //Configure GEC parameters
                //Upper box
                //oElectric.GECParameterList["UO1"].value = (uint)GEC.Param.Top_traffic_light_red;
                //oElectric.GECParameterList["UO2"].value = (uint)GEC.Param.Top_traffic_light_green;
                //changeFunctionTextPLCInput("UO1", oElectric.IDFunctions[oElectric.GECParameterList["UO1"].value]);
                //changeFunctionTextPLCInput("UO2", oElectric.IDFunctions[oElectric.GECParameterList["UO2"].value]);

                ////Lower box
                //oElectric.GECParameterList["LO1"].value = (uint)GEC.Param.Bottom_traffic_light_red;
                //oElectric.GECParameterList["LO2"].value = (uint)GEC.Param.Bottom_traffic_light_green;
                //changeFunctionTextPLCInput("LO1", oElectric.IDFunctions[oElectric.GECParameterList["LO1"].value]);
                //changeFunctionTextPLCInput("LO2", oElectric.IDFunctions[oElectric.GECParameterList["LO2"].value]);
            }
        }

        public void draw_Level_Water()
        {
            if (((Caracteristic)oElectric.CaractComercial["TNCR_OT_NIVEL_AGUA"]).CurrentReference.Equals("S"))
            {
                setStatusText("Insertando Sensor Nivel de Agua");

                //Insert graphical macro
                insertNewPage(oProject, "Lower Sensors II", "Lower Sensors I");
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444000 - NIVEL DE AGUA ADV.ema", 'A', "Lower Sensors II", 180, 260);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444000 - NIVEL DE AGUA ADV.ema", 'B', "Lower Diagnostic Inputs IV", 392, 164);

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "LI32", (uint)GEC.Param.Water_detection_bottom, true);
            }
        }

        public void draw_Main_shaft_maintenance_lock()
        {
            if (!((Caracteristic)oElectric.CaractComercial["TNCR_OT_CERROJO_MANTENIMIENTO"]).CurrentReference.Equals("N"))
            {
                setStatusText("Insertando Cerrojo en eje principal");

                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004441800 - CERROJO MANTENIMIENTO ADV.ema", 'A', "Upper Diagnostic Inputs III", 52, 112);

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "UI16", (uint)GEC.Param.Chain_locking_device_SS, true);
            }
        }

        public void draw_fotocélula_VVF()
        {
            string deteccion = ((Caracteristic)oElectric.CaractComercial["FLICHTINT"]).CurrentReference;
            string modoFuncionamiento = ((Caracteristic)oElectric.CaractComercial["FBETRART"]).CurrentReference;
            if (deteccion.Equals("LICHTINT") ||
                (deteccion.Equals("RADAR") && 
                (modoFuncionamiento.Equals("INTERM") || modoFuncionamiento.Equals("SG") || modoFuncionamiento.Equals("SGBV"))))
            {
                insertNewPage(oProject, "Upper People Detection", "Upper Diagnostic Outputs I");
                insertNewPage(oProject, "Lower People Detection", "Lower Diagnostic Outputs I");
                if (!((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                {
                    setStatusText("Insertando Fotocélulas modelo Legacy");
                    //Upper Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442600 - FOTOCELULA SUP. ADV.ema", 'A', "Upper People Detection", 16, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442600 - FOTOCELULA SUP. ADV.ema", 'B', "Upper Diagnostic Inputs III", 400, 168);

                    //Lower Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444800 - FOTOCELULA INF. ADV.ema", 'A', "Lower People Detection", 16, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444800 - FOTOCELULA INF. ADV.ema", 'B', "Lower Diagnostic Inputs III", 396, 152);
                }
                else
                {
                    setStatusText("Insertando Fotocélulas modelo Classic");
                    //Upper Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442500 - FOTOCELULA CL. SUP.ADV.ema", 'A', "Upper People Detection", 16, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442500 - FOTOCELULA CL. SUP.ADV.ema", 'B', "Upper Diagnostic Inputs III", 400, 168);

                    //Lower Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444700 - FOTOCELULA CL. INF.ADV.ema", 'A', "Lower People Detection", 16, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444700 - FOTOCELULA CL. INF.ADV.ema", 'B', "Lower Diagnostic Inputs III", 396, 152);
                }

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "UI21", (uint)GEC.Param.Top_light_barrier_comb_plate_NC, true);
                SetGECParameter(oProject, oElectric, "LI21", (uint)GEC.Param.Bottom_light_barrier_comb_plate_NC, true);
            }
        }

        public void draw_radar()
        {
            if (((Caracteristic)oElectric.CaractComercial["FLICHTINT"]).CurrentReference.Equals("RADAR"))
            {
                insertNewPage(oProject, "Upper People Detection", "Upper Diagnostic Outputs I");
                insertNewPage(oProject, "Lower People Detection", "Lower Diagnostic Outputs I");
                if (!((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                {
                    setStatusText("Insertando radares modelo Legacy");
                    //Upper Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442800 - RADAR SUP. ADV.ema", 'A', "Upper People Detection", 240, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442800 - RADAR SUP. ADV.ema", 'B', "Upper Diagnostic Inputs III", 280, 168);

                    //Lower Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004445000 - RADAR INF. ADV.ema", 'A', "Lower People Detection", 240, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004445000 - RADAR INF. ADV.ema", 'B', "Lower Diagnostic Inputs III", 276, 152);
                }
                else
                {
                    setStatusText("Insertando radares modelo Classic");
                    //Upper Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442700 - RADAR CL. SUP. ADV.ema", 'A', "Upper People Detection", 240, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004442700 - RADAR CL. SUP. ADV.ema", 'B', "Upper Diagnostic Inputs III", 280, 168);

                    //Lower Box
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444900 - RADAR CL. INF. ADV.ema", 'A', "Lower People Detection", 240, 256);
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004444900 - RADAR CL. INF. ADV.ema", 'B', "Lower Diagnostic Inputs III", 276, 152);
                }

                //Configure GEC parameters
                SetGECParameter(oProject, oElectric, "UI19", (uint)GEC.Param.Top_radar_right_NC, true);
                SetGECParameter(oProject, oElectric, "UI20", (uint)GEC.Param.Top_radar_left_NC, true);
                SetGECParameter(oProject, oElectric, "LI19", (uint)GEC.Param.Bottom_radar_right_NC, true);
                SetGECParameter(oProject, oElectric, "LI20", (uint)GEC.Param.Bottom_radar_left_NC, true);

            }
        }

        public void draw_automatic_lubrication()
        {

            if (((Caracteristic)oElectric.CaractComercial["TNCR_ENGRASE_AUTOMATICO"]).CurrentReference.Equals("S"))
            {
                setStatusText("Insertando Bomba de lubricacion");

                insertNewPage(oProject, "Upper Traffic Lights & Oil Pump", "Upper People Detection");
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004320300 - Instalacion Bomba Engrase_ADV.ema", 'A', "Upper Traffic Lights & Oil Pump", 288, 108);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004320300 - Instalacion Bomba Engrase_ADV.ema", 'B', "Control Outputs III", 276, 76);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004320300 - Instalacion Bomba Engrase_ADV.ema", 'C', "Control Inputs I", 168, 220);

                ////Configure GEC parameters
                //oElectric.GECParameterList["I4"].setValue((uint)GEC.Param.Oil_level_in_pump_1);
                //changeFunctionTextPLCInput("I4", oElectric.IDFunctions[oElectric.GECParameterList["I4"].value]);

                //if (((Caracteristic)oElectric.CaractComercial["FMODELL"]).CurrentReference.Contains("CLASSIC"))
                //    oElectric.GECParameterList["O9"].setValue((uint)GEC.Param.Oil_pump_control_1);
                //else
                //    oElectric.GECParameterList["O9"].setValue((uint)GEC.Param.Oil_pump_activation);
                //changeFunctionTextPLCInput("O9", oElectric.IDFunctions[oElectric.GECParameterList["O9"].value]);

                ////C6	OIL_PUMP_CONTROL
                //oElectric.GECParameterList["C6"].setValue((uint)GEC.Active.Enable);

                ////C7	OIL_PUMP1_TIMER_ON
                //oElectric.GECParameterList["C7"].setValue(Configurador_Form.Calc_OIL_PUMP1_TIMER_ON());

                ////C8	OIL_PUMP1_CYCLE_TIME
                //oElectric.GECParameterList["C8"].setValue(Configurador_Form.Calc_OIL_PUMP1_CYCLE_TIME());
            }
        }

        public void draw_auxiliary_brake()
        {


        }
        public void draw_Paquete_Mercadona()
        {
            if (((Caracteristic)oElectric.CaractIng["PAQUETE_ESP"]).CurrentReference.Equals("MERCADONA"))
            {
                setStatusText("Insertando Paquete Mercadona");

                //Rolling door
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004308000 - Paquete Mercadona_ADV.ema", 'A', "Upper Diagnostic Inputs III", 108, 112);
                //oElectric.GECParameterList["UI17"].value = (uint)GEC.Param.Shutter_rolling_door_SS;
                //changeFunctionTextPLCInput("UI17", oElectric.IDFunctions[oElectric.GECParameterList["UI17"].value]);

                //Rele temporizado
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004308000 - Paquete Mercadona_ADV.ema", 'B', "Control I", 304, 204);
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200004308000 - Paquete Mercadona_ADV.ema", 'C', "Control I", 44, 164);
                moveSymbol(oProject, "Control I", 8, 0, 8, 174, 88, 288);
            }
        }

        #region Metodos auxiliares
       
        public void updateCableLengthProperty(string cableName, double length, string location)
        {
            try
            {
                string auxstr;
                if (cableName.Split('.').Length > 3)
                    auxstr = String.Concat(cableName.Split('.')[0], ".", cableName.Split('.')[1], ".", cableName.Split('.')[2], "_", cableName.Split('.')[3]);

                else
                    auxstr = cableName;

                //Update property
                AnyPropertyId propertyId = new AnyPropertyId(oProject, auxstr);
                PropertyValue propertyValue = oProject.Properties[propertyId];

                CultureInfo esES = CultureInfo.CreateSpecificCulture("es-ES");
                propertyValue.Set(String.Format(esES, "{0:0.00}", length));

                //Update Cable Length
                FunctionsFilter ff = new FunctionsFilter();
                ff.Category = Function.Enums.Category.Cable;
                FunctionBasePropertyList functionBasePropertyList = new FunctionBasePropertyList();
                functionBasePropertyList.DESIGNATION_LOCATION.Set(location);
                ff.SetFilteredPropertyList(functionBasePropertyList);
                DMObjectsFinder objFinder = new DMObjectsFinder(oProject);

                Function[] ofunctions = objFinder.GetFunctions(ff);

                foreach (Function f in ofunctions)
                {
                    if (f.VisibleName.Trim('-').Equals(auxstr.Split('.')[2].Replace("_", ".")))
                        f.Properties[Eplan.EplApi.DataModel.Properties.Function.FUNC_CABLELENGTH_VALUE] = length;
                }
            }
            catch (Exception ex)
            {
                ;
            }
        }

        public void placeCableGland(string cableName, string boxName)
        {
            //Find cable glands placed in diagram
            FunctionsFilter ffCableGlands = new FunctionsFilter();
            ffCableGlands.Category = Function.Enums.Category.CabMechConduit;
            FunctionBasePropertyList fBPCableGland = new FunctionBasePropertyList();
            fBPCableGland.DESIGNATION_LOCATION.Set(boxName);
            DMObjectsFinder cableGlandsFinder = new DMObjectsFinder(oProject);
            //ffCableGlands.SetFilteredPropertyList(fBPCableGland);

            //Find cables placed in diagram
            FunctionsFilter ffCable = new FunctionsFilter();
            ffCable.Category = Function.Enums.Category.Cable;
            FunctionBasePropertyList fBPCable = new FunctionBasePropertyList();
            fBPCable.DESIGNATION_LOCATION.Set(boxName);
            fBPCable.FUNC_VISIBLENAME.Set(cableName);
            DMObjectsFinder cableFinder = new DMObjectsFinder(oProject);
            ffCable.SetFilteredPropertyList(fBPCable);

            Function[] fCableGlands = cableGlandsFinder.GetFunctions(ffCableGlands);
            Function[] fCables = cableFinder.GetFunctions(ffCable);

            AnyPropertyId propertyId = new AnyPropertyId(oProject, "Cable.Hand");
            PropertyValue propertyValue = fCables[0].Properties[propertyId];

            double Row, Col, Row1x, Row1y, Row2x, Row2y, Rowx, Rowy;
            Row = 0;
            Col = 0;
            Row1x = 0;
            Row1y = 0;
            Row2x = 0;
            Row2y = 0;
            Rowx = 0;
            Row1y = 0;

            if (propertyValue.ToMultiLangString().GetString(ISOCode.Language.L___).Equals("Derecha"))
            {
                foreach (Function fCableGland in fCableGlands)
                {
                    if (fCableGland.Location.X < 500 &&
                        fCableGland.NameParts.DESIGNATION_LOCATION.ToString().Equals(boxName))
                    {
                        AnyPropertyId propertyCableGlandId = new AnyPropertyId(oProject, "Cable.CableGlandPos");
                        PropertyValue propertyValueCableGlandId = fCableGland.Properties[propertyCableGlandId];

                        double auxRow, auxCol;

                        double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[0], out auxRow);
                        double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[1], out auxCol);

                        if (auxRow > Row && auxRow < 50)
                            Row = auxRow;

                        if (auxCol == 1 && auxRow == 1)
                        {
                            Row1x = fCableGland.Location.X;
                            Row1y = fCableGland.Location.Y;
                        }

                        if (auxCol == 2 && auxRow == 2)
                        {
                            Row2x = fCableGland.Location.X;
                            Row2y = fCableGland.Location.Y;
                        }
                    }

                }

                foreach (Function fCableGland in fCableGlands)
                {
                    AnyPropertyId propertyCableGlandId = new AnyPropertyId(oProject, "Cable.CableGlandPos");
                    PropertyValue propertyValueCableGlandId = fCableGland.Properties[propertyCableGlandId];

                    double auxRow, auxCol;

                    double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[0], out auxRow);
                    double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[1], out auxCol);

                    if (fCableGland.Location.X > 500 &&
                        fCableGland.NameParts.DESIGNATION_LOCATION.ToString().Equals(boxName))
                    {
                        if (auxRow == Row)
                        {
                            if (auxCol > Col)
                                Col = auxCol;
                        }
                    }
                }

                if (Col == 1)
                {
                    Rowx = Row2x;
                    Rowy = Row1y - (Row - 1) * (Row1y - Row2y);
                    long oident = insertSymbol(oProject, @"GR_PG13", "IEC_ED_ESS", 'A', String.Concat("Cable Entry ", boxName), Rowx, Rowy);

                }
                else
                {
                    Rowx = Row1x;
                    Rowy = Row1y - (Row) * (Row1y - Row2y);
                    long oident = insertSymbol(oProject, @"GR_PG13", "IEC_ED_ESS", 'A', String.Concat("Cable Entry ", boxName), Rowx, Rowy);
                }
            }

            if (propertyValue.ToMultiLangString().GetString(ISOCode.Language.L___).Equals("Izquierda"))
            {
                foreach (Function fCableGland in fCableGlands)
                {
                    if (fCableGland.Location.X > 500 &&
                        fCableGland.NameParts.DESIGNATION_LOCATION.ToString().Equals(boxName))
                    {
                        AnyPropertyId propertyCableGlandId = new AnyPropertyId(oProject, "Cable.CableGlandPos");
                        PropertyValue propertyValueCableGlandId = fCableGland.Properties[propertyCableGlandId];

                        double auxRow, auxCol;

                        double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[0], out auxRow);
                        double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[1], out auxCol);

                        if (auxRow > Row && auxRow < 50)
                            Row = auxRow;

                        if (auxCol == 1 && auxRow == 1)
                        {
                            Row1x = fCableGland.Location.X;
                            Row1y = fCableGland.Location.Y;
                        }

                        if (auxCol == 2 && auxRow == 2)
                        {
                            Row2x = fCableGland.Location.X;
                            Row2y = fCableGland.Location.Y;
                        }
                    }

                }

                foreach (Function fCableGland in fCableGlands)
                {
                    AnyPropertyId propertyCableGlandId = new AnyPropertyId(oProject, "Cable.CableGlandPos");
                    PropertyValue propertyValueCableGlandId = fCableGland.Properties[propertyCableGlandId];

                    double auxRow, auxCol;

                    double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[0], out auxRow);
                    double.TryParse(propertyValueCableGlandId.ToMultiLangString().GetString(ISOCode.Language.L_en_US).Split('.')[1], out auxCol);

                    if (fCableGland.Location.X > 500 &&
                        fCableGland.NameParts.DESIGNATION_LOCATION.ToString().Equals(boxName))
                    {
                        if (auxRow == Row)
                        {
                            if (auxCol > Col)
                                Col = auxCol;
                        }
                    }
                }

                if (Col == 1)
                {
                    Rowx = Row2x;
                    Rowy = Row1y - (Row - 1) * (Row1y - Row2y);
                    long oident = insertSymbol(oProject, @"GR_PG13", "IEC_ED_ESS", 'A', String.Concat("Cable Entry ", boxName), Rowx, Rowy);

                }
                else
                {
                    Rowx = Row1x;
                    Rowy = Row1y - (Row) * (Row1y - Row2y);
                    long oident = insertSymbol(oProject, @"GR_PG13", "IEC_ED_ESS", 'A', String.Concat("Cable Entry ", boxName), Rowx, Rowy);
                }
            }

            Function oInsertedCableGland = cableGlandsFinder.GetFunctions(ffCableGlands)[cableGlandsFinder.GetFunctions(ffCableGlands).Length - 1];
            //Set Name
            oInsertedCableGland.NameParts = fCables[0].Page.Functions[0].NameParts;
            oInsertedCableGland.Name = String.Concat(oProject, "=ESC+", boxName, cableName);
            oInsertedCableGland.VisibleName = cableName.Trim('-');
            //Set properties
            AnyPropertyId propertyoInsertedCableGlandId = new AnyPropertyId(oProject, "Cable.CableGlandPos");
            PropertyValue propertyValueoInsertedCableGlandId = oInsertedCableGland.Properties[propertyoInsertedCableGlandId];
            MultiLangString multiLangString = new MultiLangString();
            if (Col == 1)
                multiLangString.AddString(ISOCode.Language.L_en_US, String.Concat(Row.ToString(), ".", 2));
            else
                multiLangString.AddString(ISOCode.Language.L_en_US, String.Concat((Row + 1).ToString(), ".", 1));
            propertyValueoInsertedCableGlandId.Set(multiLangString);

            //Set view schema
            oInsertedCableGland.PropertyPlacementsSchemas.Selected = oInsertedCableGland.PropertyPlacementsSchemas.All[2];

        }

        public void paramGEC()
        {
            //String path = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "GEC\\GEC.csv");
            //String data = "";
            //foreach (GEC gEC in oElectric.GECParameterList.Values)
            //{
            //    if (gEC.active)
            //        data = String.Concat(data, "GEC_", gEC.ID, ";", gEC.getValue(), "\r\n");
            //}
            //File.WriteAllText(path, data);
        }

        public void setStatusText(string text)
        {
            //stateTexbox.ResetText();
            //stateTexbox.AppendText(text);
        }


        #endregion

    }
}
