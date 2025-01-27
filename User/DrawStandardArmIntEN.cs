using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.HEServices;
using EPLAN_API.SAP;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API.User
{
    public class DrawStandardArmIntEN
    {
        private Project oProject;
        private Page[] oPages;
        private Function[] oCablesFunctions;
        //private Cable[] oCables;
        private Hashtable oHPages;
        Dictionary<int, string> dictPages;
        //private Cable[] sCables;
        private Electric oElectric;
        private DrawingService oDs;
        private string log;
        private ProgressBar oProgressBar;
        private String OE;

        public DrawStandardArmIntEN(Project project, Electric electric) 
        {
            oProject = project;
            oElectric = electric;
            DrawMacro();
        }

        public void DrawMacro()
        {
            Caracteristic c, c2, c3, c4;
            String refVal, refVal2, refVal3, refVal4;


            //Sensores de Freno
            Draw_Freno();

            //Sensores sincronismo
            Draw_Sincronismo();

            //Display
            Draw_Display();


            //Termico Motor
            c = (Caracteristic)oElectric.CaractComercial["FANTREHT"];
            Draw_Termico(c.CurrentReference);

            //PLC
            c = (Caracteristic)oElectric.CaractIng["TNCR_DO_CONTROL"];
            if (c.CurrentReference.Equals("GEC+PLC"))
            {
                Draw_PLC();
                log = String.Concat(log, "\nIncluido PLC");
            }


            //oProgressBar.Maximum = oElectric.Caracteristics.Count;
            oProgressBar.Maximum = 28;
            oProgressBar.Step = 1;

            //Segundo freno
            c = (Caracteristic)oElectric.CaractComercial["FBREMSE2"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("4/4"))
                {
                    c2 = (Caracteristic)oElectric.CaractComercial["FANTREHT"];
                    refVal2 = c2.CurrentReference;
                    Draw_Freno_adicional(refVal2);
                    log = String.Concat(log, "\nIncluido segundo freno");
                }
            }
            oProgressBar.Value++;

            //Bomba de lubricación
            c = (Caracteristic)oElectric.CaractComercial["TNCR_ENGRASE_AUTOMATICO"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("S") ||
                refVal.Equals("C"))
                {
                    Draw_Lubricacion_auto();
                    log = String.Concat(log, "\nIncluida Bomba de lubricacion");
                }
            }
            oProgressBar.Value++;

            //Variador
            c = (Caracteristic)oElectric.CaractComercial["TNCR_SD_SIST_AHORRO"];


            refVal = c.CurrentReference;

            if (refVal != null)
            {
                if (refVal.Equals("VA"))
                {
                    c2 = (Caracteristic)oElectric.CaractComercial["TNPOTENCIAMOTOR"];
                    Draw_VVF(c2.NumVal);
                    log = String.Concat(log, "\nIncluido VVF");

                    c3 = (Caracteristic)oElectric.CaractComercial["TNCR_OT_BYPASS_VARIADOR"];
                    if (c3.CurrentReference.Equals("N"))
                    {
                        Draw_DeleteBypass();
                        log = String.Concat(log, "\nEliminado bypass de VVF");
                    }
                }
                else if (refVal.Equals("NO"))
                {
                    Draw_NoVVF();
                }
            }
            oProgressBar.Value++;

            //Micros de Zócalo
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            if (c.NumVal >= 4)
            {
                Draw_Micros_Zocalo(Convert.ToInt16(c.NumVal));
                log = String.Concat(log, "\nIncluidos ", Convert.ToInt16(c.NumVal).ToString(), " micros de zócalo");
            }
            oProgressBar.Value++;

            //Contacto de fuego
            c = (Caracteristic)oElectric.CaractComercial["TNCR_CONTACTO_FUEGO"];
            if (c.CurrentReference.Equals("S"))
            {
                Draw_Contacto_Fuego();
                log = String.Concat(log, "\nIncluido contacto de fuego");
            }
            oProgressBar.Value++;

            //Stop carrito superior
            c = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (!refVal.Equals("KEINE"))
                {
                    Draw_StopCarritoSuperior();
                    log = String.Concat(log, "\nIncluido Stop carrito superior");
                }
            }
            oProgressBar.Value++;

            //Stop carrito inferior
            c = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (!refVal.Equals("KEINE"))
                {
                    Draw_StopCarritoInferior();
                    log = String.Concat(log, "\nIncluido Stop carrito inferior");
                }
            }
            oProgressBar.Value++;

            //Stop adicional superior
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("S") ||
               refVal.Equals("2"))
                {
                    Draw_StopAdicionalSuperior();
                    log = String.Concat(log, "\nIncluido Stop adicional superior");
                }
            }

            //Stop adicional inferior
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("2"))
                {
                    Draw_StopAdicionalInferior();
                    log = String.Concat(log, "\nIncluido Stop adicional inferior");
                }
            }
            oProgressBar.Value++;


            //Seguridad vertical de peines
            c = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("INDEPENDIENTE"))
                {
                    Draw_SeguridadVerticalPeines();
                    log = String.Concat(log, "\nIncluida seguridad vertical de peines");
                }
            }
            oProgressBar.Value++;

            //Seguridad de buggy inferior
            c = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("BUGGY") ||
                refVal.Equals("BUGGYUT"))
                {
                    Draw_BuggyInferior();
                    log = String.Concat(log, "\nIncluida seguridad buggy inferior");
                }
            }
            oProgressBar.Value++;

            //Seguridad de buggy superior
            c = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("BUGGY") ||
                refVal.Equals("BUGGYOT"))
                {
                    Draw_BuggySuperior();
                    log = String.Concat(log, "\nIncluida seguridad buggy superior");
                }
            }
            oProgressBar.Value++;

            //Control de desgaste de frenos
            c = (Caracteristic)oElectric.CaractComercial["F01ZUB"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("INDUCTIVO"))
                {
                    Draw_ControlDesgasteFrenos();
                    log = String.Concat(log, "\nIncluido control de desgaste de frenos");
                }
            }
            oProgressBar.Value++;

            //Rotura pasamanos
            c = (Caracteristic)oElectric.CaractComercial["F09ZUB1"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("BRUCHSCHALT"))
                {
                    Draw_RoturaPasamanos();
                    log = String.Concat(log, "\nIncluida seguridad de rotura de pasamanos");
                }
            }
            oProgressBar.Value++;

            //Seguridad de cadena principal
            c = (Caracteristic)oElectric.CaractComercial["TNCR_S_DRIVE_CHAIN"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("SI"))
                {
                    Draw_RoturaCadenaPrincipal();
                    log = String.Concat(log, "\nIncluida seguridad de cadena principal");
                }
            }
            oProgressBar.Value++;

            //Trinquete Magnetico
            c = (Caracteristic)oElectric.CaractComercial["FZUSBREMSE"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("HWSPERRKMAGN") || refVal.Equals("NAB"))
                {
                    Draw_Trinquete_Magnetico(refVal);
                    log = String.Concat(log, "\nIncluido Trinquete Magnetico");
                }
            }
            oProgressBar.Value++;

            //Trinquete Mecánico
            c = (Caracteristic)oElectric.CaractComercial["FZUSBREMSE"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("HWSPERRKMECH"))
                {
                    Draw_Trinquete_Mecanico();
                    log = String.Concat(log, "\nIncluido Trinquete Mecanico");
                }
            }
            oProgressBar.Value++;

            //Cerrojo mantenimiento eje ppal
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_CERROJO_MANTENIMIENTO"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("S") || refVal.Equals("P"))
                {
                    Draw_Cerrojo();
                    log = String.Concat(log, "\nIncluido Cerrojo de mantenimiento");
                }
            }
            oProgressBar.Value++;

            //Detección de personas por radar
            c = (Caracteristic)oElectric.CaractComercial["FLICHTINT"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("RADAR"))
                {
                    Draw_Radar();
                    log = String.Concat(log, "\nIncluidos Radares");
                }
            }
            oProgressBar.Value++;

            //Detección de personas por fotocélulas
            c = (Caracteristic)oElectric.CaractComercial["FLICHTINT"];
            refVal = c.CurrentReference;
            c2 = (Caracteristic)oElectric.CaractComercial["FBETRART"];
            refVal2 = c2.CurrentReference;
            if (refVal != null && refVal2 != null)
            {
                if (refVal.Equals("LICHTINT") ||
                (refVal.Equals("RADAR") &&
                (refVal2.Equals("INTERM") || refVal2.Equals("SG") || refVal2.Equals("SGBV"))))
                {
                    Draw_Fotocelulas();
                    log = String.Concat(log, "\nIncluidas fotocelulas de peines");
                }
            }
            oProgressBar.Value++;

            //Llavin Automatico/Continuo
            c = (Caracteristic)oElectric.CaractComercial["LLAVES_AUT_CONT"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("E") ||
                refVal.Equals("P") ||
                refVal.Equals("B"))
                {
                    Draw_Llavin_Auto_Cont();
                    log = String.Concat(log, "\nIncluido Llavin de Aut/Cont");
                }
            }
            oProgressBar.Value++;

            //Llavin Local/Remoto
            c = (Caracteristic)oElectric.CaractComercial["LLAVES_LOCAL_REM"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("E") ||
                refVal.Equals("P") ||
                refVal.Equals("B") ||
                refVal.Equals("S"))
                {
                    Draw_Llavin_Local_Remoto();
                    log = String.Concat(log, "\nIncluido Llavin de Local/Remoto");
                }
            }
            oProgressBar.Value++;

            //Semáforos
            c = (Caracteristic)oElectric.CaractComercial["FAMPELSYM"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (!refVal.Equals("NINGUNO"))
                {
                    Draw_Semaforo(refVal);
                }
            }
            oProgressBar.Value++;

            //Luz de foso
            c = (Caracteristic)oElectric.CaractComercial["FBELANTRST"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (!refVal.Equals("KEINE"))
                {
                    Draw_LuzFoso(refVal);
                }
            }
            oProgressBar.Value++;

            //Luz Estroboscopica
            c = (Caracteristic)oElectric.CaractComercial["FSHEITSBEL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (!refVal.Equals("KEINE"))
                {
                    Draw_LuzEstroboscopica(refVal);
                }
            }

            oProgressBar.Value++;

            //Luz de peines
            c = (Caracteristic)oElectric.CaractComercial["ILUPEI"];
            c2 = (Caracteristic)oElectric.CaractComercial["FSHEITSBEL"];
            refVal = c.CurrentReference;
            refVal2 = c2.CurrentReference;
            if (refVal != null && refVal2 != null)
            {
                if (!refVal.Equals("NO"))
                {
                    Draw_LuzPeines(refVal, refVal2);
                }
            }

            oProgressBar.Value++;

            //Luz bajopasamanos
            c = (Caracteristic)oElectric.CaractComercial["FBALUBEL"];
            c2 = (Caracteristic)oElectric.CaractComercial["FSHEITSBEL"];
            c3 = (Caracteristic)oElectric.CaractComercial["ILUPEI"];
            refVal = c.CurrentReference;
            refVal2 = c2.CurrentReference;
            refVal3 = c3.CurrentReference;
            if (refVal != null && refVal2 != null && refVal3 != null)
            {
                if (!refVal.Equals("BELOHNE"))
                {
                    Draw_LuzBajopasamanos(refVal, refVal2, refVal3);
                }
            }

            oProgressBar.Value++;

            //Luz Zocalos
            c = (Caracteristic)oElectric.CaractComercial["FSOCKELBEL"];
            c2 = (Caracteristic)oElectric.CaractComercial["FBALUBEL"];
            c3 = (Caracteristic)oElectric.CaractComercial["FSHEITSBEL"];
            c4 = (Caracteristic)oElectric.CaractComercial["ILUPEI"];
            refVal = c.CurrentReference;
            refVal2 = c2.CurrentReference;
            refVal3 = c3.CurrentReference;
            refVal4 = c4.CurrentReference;
            if (refVal != null && refVal2 != null && refVal3 != null && refVal4 != null)
            {
                if (!refVal.Equals("BELOHNE"))
                {
                    Draw_LuzZocalos(refVal, refVal2, refVal3, refVal4);
                }
            }

            oProgressBar.Value++;

            //Sistema Andén
            c = (Caracteristic)oElectric.CaractComercial["FWIEDERB"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (!refVal.Equals("KEINE"))
                {
                    Draw_Safety_Curtain();
                    log = String.Concat(log, "\nIncluido Sistema Anden");
                }
            }

            oProgressBar.Value++;

            //Paquete Francia
            //Paquetes especiales
            c = (Caracteristic)oElectric.CaractIng["PAQUETE_ESP"];
            refVal = c.CurrentReference;
            if (refVal.Equals("FRANCIA"))
                drawfrenchpakage();
            else if (refVal.Equals("MERCADONA"))
            {
                drawMercadonapakage();
            }

            Reports report = new Reports();
            report.GenerateProject(oProject);

            //Redraw
            Edit edit = new Edit();
            edit.RedrawGed();

            MessageBox.Show(log);
            String path = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "Log\\Log_Draw.txt");
            File.WriteAllText(path, log);


        }

        public void RenumeraPaginas()
        {
            //Renumeramos páginas
            PagePropertyList pagePropList = new PagePropertyList();

            for (int i = oPages.Length - 1; i >= 0; i--)
            {
                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                pagePropList.PAGE_COUNTER = new PagePropertyList(oPages[i]).PAGE_COUNTER + 1;
                oPages[i].SetName(pagePropList);
            }


            for (int i = 0; i < oPages.Length; i++)
            {
                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                pagePropList.PAGE_COUNTER = i + 1;
                oPages[i].SetName(pagePropList);
            }

            GetPageTable();
        }

        public void GetPageTable()
        {
            oPages = oProject.Pages;
            dictPages = new Dictionary<int, string>();


            for (int i = 0; i < oPages.Length; i++)
            {
                dictPages.Add(i, oProject.Pages[i].Properties.PAGE_NOMINATIOMN.ToMultiLangString().GetStringToDisplay(ISOCode.Language.L_en_US));
            }

            oHPages = new Hashtable(dictPages);
        }

        #region Metodos de dibujo_
        public void Draw_Freno()
        {
            string unidadAccionamiento = ((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference;

            insertNewPage("Brake Sensors", "Motor Sensors");

            log = "Incluidos sensores de freno: ";

            if (unidadAccionamiento.Equals("QC") ||
                unidadAccionamiento.Equals("FJ"))
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\200002262200_00_BRAKE_ELECTRICAL_ASSEMBLY_UE.ema", 'P', "Brake Sensors", 44, 244);
                log = String.Concat(log, "Finales de carrera");
            }
            else
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\200002262200_00_BRAKE_ELECTRICAL_ASSEMBLY_UE.ema", 'O', "Brake Sensors", 44, 244);
                log = String.Concat(log, "Inductivos");
            }

        }
        
        public void Draw_Sincronismo()
        {
            string Modelo = (oElectric.CaractComercial["FMODELL"] as Caracteristic).CurrentReference;

            //Sensores superiores
            if (Modelo.Contains("CLASSIC"))
            {
                insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_I_VC3_0.emp", "Upper Diagnostic Outputs I", "Upper Sensors I");
            }
            else
            {
                insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_I_VE.emp", "Upper Diagnostic Outputs I", "Upper Sensors I");
            }

            //Sensores superiores
            if (Modelo.Contains("CLASSIC"))
            {
                insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_I_VC3_0.emp", "Lower Diagnostic Outputs I", "Lower Sensors I");
            }
            else
            {
                insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_I_VE.emp", "Lower Diagnostic Outputs I", "Lower Sensors I");
            }

        }

        public void Draw_Display()
        {
            string tipoDisplay = (oElectric.CaractIng["TNCR_DO_DISPLAY_TYPE"] as Caracteristic).CurrentReference;

            if (tipoDisplay.Equals("DDU"))
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'B', "Display", 272.0, 252.0);
            else if (tipoDisplay.Equals("ESCATRONIC"))
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'A', "Display", 272.0, 252.0); ;
        }

        #endregion

        #region Metodos de dibujo

        public void Draw_Contacto_Fuego()
        {
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Fire_Contact.ema", 'C', "Control Inputs I", 224, 268);
        }

        
        public void Draw_Llavin_Auto_Cont()
        {
            int key;
            Insert oInsert = new Insert();

            //Upper Keys"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Keys");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 0, oProject.Pages[key], new PointD(24.0, 256.0), Insert.MoveKind.Absolute);

            //Upper Diagnostic Inputs IV"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs IV");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 1, oProject.Pages[key], new PointD(220.0, 168.0), Insert.MoveKind.Absolute);
        }

        public void Draw_Llavin_Local_Remoto()
        {
            int key;
            Insert oInsert = new Insert();

            //Upper Keys"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Keys");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 2, oProject.Pages[key], new PointD(300.0, 256.0), Insert.MoveKind.Absolute);

            //Upper Diagnostic Inputs IV
            //Compruebo si ya esta insertada la página "Upper Diagnostic Inputs V"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs V");

            if (key == 0)
            {

                //Despues de la página de "Upper Diagnostic Inputs IV"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs IV");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Insertamos página
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Diagnostic_Inputs_V.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            //Upper Diagnostic Inputs V"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs V");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 3, oProject.Pages[key], new PointD(248.0, 168.0), Insert.MoveKind.Absolute);
        }

        public void Draw_LuzEstroboscopica(string type)
        {
            int key;
            Insert oInsert = new Insert();

            switch (type)
            {
                //Una tira LED
                case "LED":
                    //en página de "Upper Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 5, oProject.Pages[key], new PointD(236.0, 140.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 0, oProject.Pages[key], new PointD(168.0, 108.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 1 Tira LED");
                    break;

                //Dos tiras LED
                case "2LED":
                    //en página de "Upper Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 6, oProject.Pages[key], new PointD(236.0, 140.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 1, oProject.Pages[key], new PointD(168.0, 108.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 2 Tira LED");
                    break;

                //Tres tiras LED
                case "3LED":
                    //en página de "Upper Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 7, oProject.Pages[key], new PointD(236.0, 140.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 2, oProject.Pages[key], new PointD(168.0, 108.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 3 Tira LED");
                    break;

                //Una lampara
                case "STFSPALTBEL":
                    //en página de "Upper Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 8, oProject.Pages[key], new PointD(32.0, 176.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 3, oProject.Pages[key], new PointD(168.0, 136.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 1 Lampara");
                    break;

                //Dos lamparas
                case "STFSPALTBEL2":
                    //en página de "Upper Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 9, oProject.Pages[key], new PointD(32.0, 176.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 4, oProject.Pages[key], new PointD(168.0, 136.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 2 Lamparas");
                    break;
            }
        }

        public void Draw_LuzPeines(string type, string luzEstro)
        {
            int key;
            Insert oInsert = new Insert();

            if (type.Equals("DI"))
            {
                //en página de "Upper Lighting I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Lighting I");

                if (luzEstro.Equals("STFSPALTBEL") || luzEstro.Equals("STFSPALTBEL2"))
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 0, oProject.Pages[key], new PointD(80.0, 176.0), Insert.MoveKind.Absolute);
                else
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 1, oProject.Pages[key], new PointD(32.0, 176.0), Insert.MoveKind.Absolute);

                //en página de "Lower Lighting I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");

                if (luzEstro.Equals("STFSPALTBEL") || luzEstro.Equals("STFSPALTBEL2"))
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 3, oProject.Pages[key], new PointD(176.0, 180.0), Insert.MoveKind.Absolute);
                else if (luzEstro.Contains("LED"))
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 4, oProject.Pages[key], new PointD(176.0, 180.0), Insert.MoveKind.Absolute);
                else
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 2, oProject.Pages[key], new PointD(176.0, 180.0), Insert.MoveKind.Absolute);

                log = String.Concat(log, "\nIncluida luz de peines LED 230V");
            }
        }

        public void Draw_LuzBajopasamanos(string type, string luzEstro, string luzPeines)
        {
            int key;
            Insert oInsert = new Insert();

            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO"))
                {
                    //en página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 1, oProject.Pages[key], new PointD(156.0, 112.0), Insert.MoveKind.Absolute);
                }
                else
                {
                    //Despues de la página de "Lower Lighting I"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");

                    //Renumeramos páginas
                    RenumeraPaginas();

                    PagePropertyList pagePropList = new PagePropertyList();

                    for (int i = oPages.Length - 1; i > key; i--)
                    {
                        pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                        pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                        pagePropList.PAGE_COUNTER = i + 2;
                        oPages[i].SetName(pagePropList);
                    }
                    //Finales de carrera
                    oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", oPages[key], oProject, false);

                    GetPageTable();

                    //en página de "Lower Lighting II"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting II");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 0, oProject.Pages[key], new PointD(40.0, 188.0), Insert.MoveKind.Absolute);

                    if (!luzEstro.Equals("KEINE") || !luzPeines.Equals("NO"))
                    {
                        //en página de "Lower Lighting I"
                        key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                        oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 2, oProject.Pages[key], new PointD(176, 200.0), Insert.MoveKind.Absolute);

                    }

                }


                log = String.Concat(log, "\nIncluida luz bajpasamanos LED 230V");
            }

        }

        public void Draw_LuzZocalos(string type, string luzBajopasamanos, string luzEstro, string luzPeines)
        {
            int key;
            Insert oInsert = new Insert();

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (modelo.CurrentReference.Contains("CLASSIC"))
                {
                    insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'E', "Lower Lighting I", 176, 196.0);
                    insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");
                    insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'I', "Lower Lighting II", 36.0, 272.0);

                }
                else
                {
                    if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && luzBajopasamanos.Equals("BELOHNE"))
                    {
                        //en página de "Lower Lighting I"
                        key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                        oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 0, oProject.Pages[key], new PointD(156.0, 112.0), Insert.MoveKind.Absolute);
                    }
                    else if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && !luzBajopasamanos.Equals("BELOHNE"))
                    {
                        //en página de "Lower Lighting I"
                        key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                        oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 1, oProject.Pages[key], new PointD(228.0, 172.0), Insert.MoveKind.Absolute);
                    }
                    else
                    {
                        if (!luzBajopasamanos.Equals("BELOHNE"))
                        {
                            //Despues de la página de "Lower Lighting I"
                            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");

                            //Renumeramos páginas
                            RenumeraPaginas();

                            PagePropertyList pagePropList = new PagePropertyList();

                            for (int i = oPages.Length - 1; i > key; i--)
                            {
                                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                                pagePropList.PAGE_COUNTER = i + 2;
                                oPages[i].SetName(pagePropList);
                            }
                            //Finales de carrera
                            oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", oPages[key], oProject, false);

                            GetPageTable();

                            //en página de "Lower Lighting II"
                            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting II");
                            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 2, oProject.Pages[key], new PointD(132.0, 188.0), Insert.MoveKind.Absolute);
                        }
                        else
                        {
                            //Despues de la página de "Lower Lighting I"
                            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");

                            //Renumeramos páginas
                            RenumeraPaginas();

                            PagePropertyList pagePropList = new PagePropertyList();

                            for (int i = oPages.Length - 1; i > key; i--)
                            {
                                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                                pagePropList.PAGE_COUNTER = i + 2;
                                oPages[i].SetName(pagePropList);
                            }
                            //Lower Lighting
                            oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", oPages[key], oProject, false);

                            GetPageTable();

                            //en página de "Lower Lighting II"
                            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting II");
                            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 3, oProject.Pages[key], new PointD(40.0, 188.0), Insert.MoveKind.Absolute);

                            //en página de "Lower Lighting I"
                            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 4, oProject.Pages[key], new PointD(176, 200.0), Insert.MoveKind.Absolute);
                        }
                    }
                }

                log = String.Concat(log, "\nIncluida luz zócalo LED 230V");
            }

        }

        public void Draw_LuzFoso(string type)
        {
            int key;
            Insert oInsert = new Insert();

            if (type.Equals("HANDL") ||
                type.Equals("OVAL"))
            {
                //en página de "Main Power Supply"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Main Power Supply");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 0, oProject.Pages[key], new PointD(308.0, 60.0), Insert.MoveKind.Absolute);

                //en página de "LLower Lighting I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Lighting I");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 1, oProject.Pages[key], new PointD(24.0, 172.0), Insert.MoveKind.Absolute);

                log = String.Concat(log, "\nIncluida luz de foso manual");
            }
        }

        public void Draw_Semaforo(string type)
        {
            int key;
            Insert oInsert = new Insert();

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Semaforos superiores
            //Compruebo si ya esta insertada la página "Upper Traffic Lights"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Traffic Lights");

            if (key == 0)
            {

                //Despues de la página de "Upper Keys"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Keys");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Traffic_Lights.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            //Semaforos inferiores
            //Compruebo si ya esta insertada la página "Lower Traffic Lights"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Traffic Lights");

            if (key == 0)
            {

                //Despues de la página de "Lower Keys"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Keys");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Traffic_Lights.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            if (modelo.CurrentReference.Contains("CLASSIC"))
            {
                //en página de "Upper Traffic Lights"
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'M', "Upper Traffic Lights", 128.0, 148.0);
                //en página de "Lower Traffic Lights"
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'N', "Lower Traffic Lights", 128.0, 148.0);
                log = String.Concat(log, "\nIncluidos Semáforos Chinos de VC3.0");
            }
            else
            {
                if (type.Equals("BICOLOR"))
                {
                    //en página de "Upper Traffic Lights"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Traffic Lights");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 0, oProject.Pages[key], new PointD(128.0, 148.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Traffic Lights"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Traffic Lights");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 2, oProject.Pages[key], new PointD(128.0, 148.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluidos Semáforos Bicolor");
                }
                else if (type.Equals("ROTGRUEN") ||
                        type.Equals("F6NEINB") ||
                        type.Equals("PROHI_VERDE"))
                {
                    //en página de "Upper Traffic Lights"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Traffic Lights");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 1, oProject.Pages[key], new PointD(128.0, 148.0), Insert.MoveKind.Absolute);

                    //en página de "Lower Traffic Lights"
                    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Traffic Lights");
                    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 3, oProject.Pages[key], new PointD(128.0, 148.0), Insert.MoveKind.Absolute);

                    log = String.Concat(log, "\nIncluidos Semáforos Flecha/Prohibido");
                }
                else
                {
                    log = String.Concat(log, "\nNo existe macro para el tipo de semaforo seleccionado");
                }
            }


            //en página de "Upper Diagnostic Outputs I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Outputs I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 15, oProject.Pages[key], new PointD(36.0, 124.0), Insert.MoveKind.Absolute);

            //en página de "Lower Diagnostic Outputs I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Outputs I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 14, oProject.Pages[key], new PointD(36.0, 124.0), Insert.MoveKind.Absolute);



        }

        public void Draw_Radar()
        {
            int key;
            Insert oInsert = new Insert();
            Caracteristic producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Radares Superiores
            //Compruebo si ya esta insertada la página "Upper People Detection"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper People Detection");

            if (key == 0)
            {

                //Despues de la página de "Upper Diagnostic Outputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Outputs I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            //en página de "Upper People Detection"
            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'I', "Upper People Detection", 200.0, 256.0);
            else
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'A', "Upper People Detection", 200.0, 256.0);


            //en página de "Upper Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 15, oProject.Pages[key], new PointD(280.0, 168.0), Insert.MoveKind.Absolute);

            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Radares Inferiores
            //Compruebo si ya esta insertada la página "Lower People Detection"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower People Detection");

            if (key == 0)
            {

                //Despues de la página de "Lower Diagnostic Outputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Outputs I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            //en página de "Lower People Detection"
            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'J', "Lower People Detection", 200.0, 256.0);
            else
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'B', "Lower People Detection", 200.0, 256.0);


            //en página de "Lower Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 14, oProject.Pages[key], new PointD(280.0, 168.0), Insert.MoveKind.Absolute);

        }

        public void Draw_Fotocelulas()
        {
            int key;
            Insert oInsert = new Insert();

            Caracteristic producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Fotocélulas superiores
            //Compruebo si ya esta insertada la página "Upper People Detection"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper People Detection");

            if (key == 0)
            {

                //Despues de la página de "Upper Diagnostic Outputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Outputs I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'H', "Upper People Detection", 12.0, 256.0);
            else
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'B', "Upper People Detection", 12.0, 256.0);


            //en página de "Upper Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 15, oProject.Pages[key], new PointD(400.0, 168.0), Insert.MoveKind.Absolute);


            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Fotocélulas inferiores
            //Compruebo si ya esta insertada la página "Lower People Detection"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower People Detection");

            if (key == 0)
            {

                //Despues de la página de "Lower Diagnostic Outputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Outputs I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'G', "Lower People Detection", 12.0, 256.0);
            else
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'A', "Lower People Detection", 12.0, 256.0);


            //en página de "Upper Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 14, oProject.Pages[key], new PointD(400.0, 168.0), Insert.MoveKind.Absolute);



        }

        public void Draw_Trinquete_Magnetico(string type)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Motor Sensors"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Motor Sensors");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 3, oProject.Pages[key], new PointD(320.0, 260.0), Insert.MoveKind.Absolute);

            //en página de "Control I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 7, oProject.Pages[key], new PointD(332.0, 212.0), Insert.MoveKind.Absolute);


            //en página de "Safety Pulse Inputs"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Safety Pulse Inputs");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 4, oProject.Pages[key], new PointD(288.0, 124.0), Insert.MoveKind.Absolute);

            if (type.Equals("HWSPERRKMAGN"))
            {
                //en página de "Brake Sensors"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Brake Sensors");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 0, oProject.Pages[key], new PointD(336.0, 244.0), Insert.MoveKind.Absolute);
            }
            else if (type.Equals("NAB"))
            {
                //en página de "Brake Sensors"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Brake Sensors");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 1, oProject.Pages[key], new PointD(336.0, 244.0), Insert.MoveKind.Absolute);

            }

            //en página de "Safety Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Safety Inputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 2, oProject.Pages[key], new PointD(256.0, 164.0), Insert.MoveKind.Absolute);

            //en página de "Control Outputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Outputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 3, oProject.Pages[key], new PointD(28.0, 104.0), Insert.MoveKind.Absolute);

            //en página de "Control I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 4, oProject.Pages[key], new PointD(20.0, 192.0), Insert.MoveKind.Absolute);

            //quitamos puntos de conexión
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.TypeIdentifier == 42 &&
                    placement.Location.Y >= 204.00 &&
                    placement.Location.Y <= 212.00 &&
                    placement.Location.X >= 72 &&
                    placement.Location.X <= 88)
                {
                    placement.Remove();
                }
            }


            //Anulado por rele termico chino

            //Caracteristic motorType = (Caracteristic)oElectric.CaractComercial["FANTREHT"];

            //if (motorType.CurrentReference.Equals("QC") ||
            //   motorType.CurrentReference.Equals("FJ"))
            //{
            //    //en página de "Control I"
            //    key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");
            //    oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 10, oProject.Pages[key], new PointD(72.0, 184.0), Insert.MoveKind.Absolute);

            //    placements = oProject.Pages[key].AllPlacements;

            //    foreach (Placement placement in placements)
            //    {
            //        if (placement.TypeIdentifier == 42 &&
            //            placement.Location.Y >= 204.00 &&
            //            placement.Location.Y <= 212.00 &&
            //            placement.Location.X >= 124 &&
            //            placement.Location.X <= 140)
            //        {
            //            placement.Remove();
            //        }
            //    }
            //}

            //SAI
            //Compruebo si ya esta insertada la página "Control II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control II");

            if (key == 0)
            {

                //Despues de la página de "Control I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Control_II_Arm_Int.emp", oPages[key], oProject, false);

                GetPageTable();
            }

        }

        public void Draw_Cerrojo()
        {
            Caracteristic sStopAdicional = (Caracteristic)oElectric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            Caracteristic sStopCarritos = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic sMicrosZocalo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            Caracteristic sPeines = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];

            if (sStopCarritos.CurrentReference.Equals("KEINE") || sStopAdicional.CurrentReference.Equals("N"))
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs III", "UI17", "Chain locking device");
            }
            else if (sMicrosZocalo.NumVal <= 6)
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs III", 76.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs III", "UI16", "Chain locking device");

            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {

                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs II", 372.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs II", "UI14", "Chain locking device");

            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs II", 256.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs II", "UI12", "Chain locking device");
            }
        }

        public void Draw_Trinquete_Mecanico()
        {
            Caracteristic sStopAdicional = (Caracteristic)oElectric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            Caracteristic sStopCarritos = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic sMicrosZocalo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            Caracteristic sPeines = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];
            Caracteristic cerrojo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_CERROJO_MANTENIMIENTO"];

            if ((sStopCarritos.CurrentReference.Equals("KEINE") || sStopAdicional.CurrentReference.Equals("N")) &&
                cerrojo.CurrentReference.Equals("N"))
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs III", "UI17", "Mechanical Pawl Brake");
            }
            else if (sMicrosZocalo.NumVal <= 6 &&
                cerrojo.CurrentReference.Equals("N"))
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 76.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs III", "UI16", "Mechanical Pawl Brake");

            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {

                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 16.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs III", "UI15", "Mechanical Pawl Brake");

            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs II", 312.0, 156.0);
                changeFunctionTextPLCInput("Upper Diagnostic Inputs II", "UI13", "Mechanical Pawl Brake");
            }
        }

        public void Draw_RoturaCadenaPrincipal()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Motor Sensors"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Motor Sensors");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 0, oProject.Pages[key], new PointD(208.0, 260.0), Insert.MoveKind.Absolute);


            //en página de "Safety Inputs I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Safety Inputs I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 15, oProject.Pages[key], new PointD(292.0, 184.0), Insert.MoveKind.Absolute);
        }

        public void Draw_RoturaPasamanos()
        {

            int key;
            Insert oInsert = new Insert();

            //Compruebo si ya esta insertada la página "Lower Sensors II""
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Sensors II");

            if (key == 0)
            {

                //Despues de la página de "Lower Sensors I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Sensors I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_II.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            //en página de "Lower Sensors II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Sensors II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 0, oProject.Pages[key], new PointD(36.0, 260.0), Insert.MoveKind.Absolute);


            //en página de "Lower Diagnostic Inputs IV"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs IV");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 15, oProject.Pages[key], new PointD(216.0, 164.0), Insert.MoveKind.Absolute);
        }

        public void Draw_ControlDesgasteFrenos()
        {
            int key;
            Insert oInsert = new Insert();


            //Compruebo si ya esta insertada la página "Upper Sensors II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Sensors II");

            if (key == 0)
            {

                //Despues de la página de "Upper Sensors I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Sensors I");

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                //Finales de carrera
                oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_II.emp", oPages[key], oProject, false);

                GetPageTable();
            }

            //en página de "Upper Sensors II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Sensors II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 0, oProject.Pages[key], new PointD(28.0, 264.0), Insert.MoveKind.Absolute);


            //en página de "Upper Diagnostic Inputs IV"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs IV");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 15, oProject.Pages[key], new PointD(340.0, 168.0), Insert.MoveKind.Absolute);
        }

        public void Draw_BuggyInferior()
        {

            int key;
            Insert oInsert = new Insert();

            //en página de "Lower Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 1, oProject.Pages[key], new PointD(16.0, 156.0), Insert.MoveKind.Absolute);
        }

        public void Draw_BuggySuperior()
        {

            int key;
            Insert oInsert = new Insert();

            //en página de "Upper Diagnostic Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 12, oProject.Pages[key], new PointD(368.0, 156.0), Insert.MoveKind.Absolute);

            //en página de "Upper Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 13, oProject.Pages[key], new PointD(16.0, 156.0), Insert.MoveKind.Absolute);
        }

        public void Draw_SeguridadVerticalPeines()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Upper Diagnostic Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 0, oProject.Pages[key], new PointD(252.0, 156.0), Insert.MoveKind.Absolute);

            //en página de "Lower Diagnostic Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 1, oProject.Pages[key], new PointD(312.0, 156.0), Insert.MoveKind.Absolute);
        }

        public void Draw_StopAdicionalSuperior()
        {
            int key;
            Insert oInsert = new Insert();

            Caracteristic sStopCarritos = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic sMicrosZocalo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            Caracteristic sPeines = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];

            if (sStopCarritos.CurrentReference.Equals("KEINE"))
            {
                //en página de "Upper Diagnostic Inputs II"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs II");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 0, oProject.Pages[key], new PointD(192.0, 156.0), Insert.MoveKind.Absolute);
            }
            else if (sMicrosZocalo.NumVal <= 6)
            {
                //en página de "Upper Diagnostic Inputs III"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs III");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 0, oProject.Pages[key], new PointD(132.0, 156.0), Insert.MoveKind.Absolute);
                changeFunctionTextPLCInput(oProject.Pages[key], "UI17", "Top External Emergency Stop");
            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {
                //en página de "Upper Diagnostic Inputs III"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs III");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 0, oProject.Pages[key], new PointD(16.0, 156.0), Insert.MoveKind.Absolute);
                changeFunctionTextPLCInput(oProject.Pages[key], "UI15", "Top External Emergency Stop");
            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                //en página de "Upper Diagnostic Inputs II"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs II");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 0, oProject.Pages[key], new PointD(312.0, 156.0), Insert.MoveKind.Absolute);
                changeFunctionTextPLCInput(oProject.Pages[key], "UI13", "Top External Emergency Stop");
            }

        }

        public void Draw_StopCarritoSuperior()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Upper Diagnostic Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_Trolley.ema", 0, oProject.Pages[key], new PointD(192.0, 156.0), Insert.MoveKind.Absolute);

            changeFunctionTextPLCInput(oProject.Pages[key], "UI11", "Top_Emergency STOP_Trolley");

        }

        public void Draw_StopAdicionalInferior()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Lower Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 1, oProject.Pages[key], new PointD(132.0, 156.0), Insert.MoveKind.Absolute);
        }

        public void Draw_StopCarritoInferior()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Lower Diagnostic Inputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs III");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_Trolley.ema", 1, oProject.Pages[key], new PointD(192.0, 156.0), Insert.MoveKind.Absolute);

            changeFunctionTextPLCInput(oProject.Pages[key], "LI18", "Bottom_Emergency STOP_Trolley");
        }

        public void Draw_Micros_Zocalo(int nMicros)
        {
            int key;
            Insert oInsert = new Insert();

            if (nMicros >= 4)
            {
                //en página de "Upper Diagnostic Inputs II"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs II");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Microswitch_1.ema", 0, oProject.Pages[key], new PointD(72.0, 156.0), Insert.MoveKind.Absolute);

                //en página de "Lower Diagnostic Inputs II"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs II");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Microswitch_1.ema", 1, oProject.Pages[key], new PointD(192.0, 156.0), Insert.MoveKind.Absolute);
            }

        }

        public void Draw_Freno(string motorType)
        {
            int key;
            Insert oInsert = new Insert();

            //Compruebo si ya esta insertada la página "Brake Sensors"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Brake Sensors");

            if (key == 0)
            {
                log = "Incluidos sensores de freno: ";

                if (motorType.Equals("QC") ||
                    motorType.Equals("FJ"))
                {
                    //Finales de carrera
                    insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_FC.emp", "Motor Sensors", "Brake Sensors");
                    log = String.Concat(log, "Finales de carrera");
                }
                else
                {
                    //Inductivos
                    insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_Sensors.emp", "Motor Sensors", "Brake Sensors");
                    log = String.Concat(log, "Inductivos");
                }

                GetPageTable();
            }
        }

        public void Draw_Termico(string motorType)
        {
            int key;
            Insert oInsert = new Insert();

            if (motorType.Equals("QC") ||
               motorType.Equals("FJ"))
            {
                //Termistor
                //en página de "Motor"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Motor");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 3, oProject.Pages[key], new PointD(296.0, 144.0), Insert.MoveKind.Absolute);

                //en página de "Control Inputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs I");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 2, oProject.Pages[key], new PointD(348.0, 268.0), Insert.MoveKind.Absolute);

                //en página de "Control Inputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs I");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 4, oProject.Pages[key], new PointD(96.0, 188.0), Insert.MoveKind.Absolute);

                log = String.Concat(log, "\nIncluido relé termico");
            }
            else
            {
                //Bimetal
                //en página de "Motor"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Motor");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 0, oProject.Pages[key], new PointD(296.0, 112.0), Insert.MoveKind.Absolute);

                //en página de "Control Inputs I"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs I");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 1, oProject.Pages[key], new PointD(96.0, 176.0), Insert.MoveKind.Absolute);

                log = String.Concat(log, "\nIncluido bimetal");
            }
        }

        public void Draw_Freno_adicional(string motorType)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Control I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 14, oProject.Pages[key], new PointD(208.0, 160.0), Insert.MoveKind.Absolute);

            //en página de "Oil pump & Brake"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Oil pump & Brake");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 6, oProject.Pages[key], new PointD(316.0, 84.0), Insert.MoveKind.Absolute);

            if (motorType.Equals("QC") ||
                motorType.Equals("FJ"))
            {
                //Finales de carrera
                //en página de "Brake Sensors"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Brake Sensors");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 1, oProject.Pages[key], new PointD(188.0, 244.0), Insert.MoveKind.Absolute);
            }
            else
            {
                //Inductivos
                //en página de "Brake Sensors"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Brake Sensors");
                oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 2, oProject.Pages[key], new PointD(176.0, 248.0), Insert.MoveKind.Absolute);
            }



            //en página de "Safety Inputs I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Safety Inputs I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 11, oProject.Pages[key], new PointD(228.0, 184.0), Insert.MoveKind.Absolute);

        }

        public void Draw_Lubricacion_auto()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Oil pump & Brake"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Oil pump & Brake");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 0, oProject.Pages[key], new PointD(44.0, 176.0), Insert.MoveKind.Absolute);

            //en página de "Control Outputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Outputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 15, oProject.Pages[key], new PointD(276.0, 168.0), Insert.MoveKind.Absolute);

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            if (modelo.CurrentReference.Contains("CLASSIC"))
                changeFunctionTextPLCInput(oProject.Pages[key], "Q9", "Oil Pump Control");

            //en página de "Control I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 1, oProject.Pages[key], new PointD(384.0, 212.0), Insert.MoveKind.Absolute);

            //en página de "Control Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs II");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 14, oProject.Pages[key], new PointD(128.0, 240.0), Insert.MoveKind.Absolute);


        }

        public void Draw_VVF(Double power)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "VVF Power"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "VVF Power");
            StorableObject[] oInsertedObjects = oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 8, oProject.Pages[key], new PointD(44.0, 208.0), Insert.MoveKind.Absolute);

            foreach (StorableObject oSOTemp in oInsertedObjects)
            {
                //we are searching for PlaceHolder 'Three-Phase' in the results
                PlaceHolder oPlaceHoldeThreePhase = oSOTemp as Eplan.EplApi.DataModel.Graphics.PlaceHolder;
                if ((oPlaceHoldeThreePhase != null)
                    && (oPlaceHoldeThreePhase.Name == "VVF"))
                {
                    if (power <= 9.0)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("9kW");
                    }
                    else if (power > 9.0 && power <= 15.0)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("15kW");
                    }
                    else if (power > 15.0 && power <= 23.5)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("23,5kW");
                    }

                }
            }

            //quitamos puntos de interrupcion
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.TypeIdentifier == 42 &&
                    placement.Location.Y == 208.00 &&
                    placement.Location.X >= 92 &&
                    placement.Location.X <= 116)
                {
                    placement.Remove();
                }
            }

            //Despues de la página de "VVF"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "VVF Power");

            //Renumeramos páginas
            RenumeraPaginas();

            PagePropertyList pagePropList = new PagePropertyList();

            for (int i = oPages.Length - 1; i > key; i--)
            {
                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                pagePropList.PAGE_COUNTER = i + 2;
                oPages[i].SetName(pagePropList);
            }

            oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\VVF_Control.emp", oPages[key], oProject, false);

            GetPageTable();

            //en página de "Control Outputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Outputs I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 14, oProject.Pages[key], new PointD(236.0, 128.0), Insert.MoveKind.Absolute);

            //en página de "Control Inputs I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs I");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 15, oProject.Pages[key], new PointD(192.0, 160.0), Insert.MoveKind.Absolute);


        }

        public void Draw_NoVVF()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Control II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs II");

            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.Y >= 184 &&
                    placement.Location.Y <= 184 &&
                    placement.Location.X >= 64 &&
                    placement.Location.X <= 64)
                {
                    placement.Remove();
                }
            }
        }

        public void Draw_DeleteBypass()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Control II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs II");

            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.Y >= 184 &&
                    placement.Location.Y <= 268 &&
                    placement.Location.X >= 64 &&
                    placement.Location.X <= 64)
                {
                    placement.Remove();
                }
            }
        }

        public void Bypass_VVF()
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Control Inputs II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs II");

            //quitamos selector de bypass si no lo hay
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.Y == 184 &&
                    placement.Location.X == 64)
                {
                    placement.Remove();
                }
            }
        }

        public void Draw_Safety_Curtain()
        {

            //int key;
            //Insert oInsert = new Insert();

            ////Despues de la página de "Display"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Display");

            ////Renumeramos páginas
            //RenumeraPaginas();

            //PagePropertyList pagePropList = new PagePropertyList();

            //for (int i = oPages.Length - 1; i > key; i--)
            //{
            //    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
            //    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
            //    pagePropList.PAGE_COUNTER = i + 2;
            //    oPages[i].SetName(pagePropList);
            //}

            //oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Curtain.emp", oPages[key], oProject, false);
            insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Curtain.emp", "Display", "Safety Curtain");

            //GetPageTable();

            //en página de "Display"
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'P', "Display", 124.0, 272.0);

            //en página de "Safety Curtain"
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'A', "Safety Curtain", 52.0, 224.0);

            //en página de "Safety Curtain"
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'B', "Safety Curtain", 52.0, 120.0);

            //en página de "Safety Curtain"
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'O', "Safety Curtain", 16.0, 288.0);


            insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Extension_Inputs.emp", "Safety Inputs II", "Safety Extension Inputs");





            ////en página de "PLC Output I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "PLC Output I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 13, oProject.Pages[key], new PointD(96.0, 120.0), Insert.MoveKind.Absolute);

            ////en página de "PLC Output I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "PLC Output I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 8, oProject.Pages[key], new PointD(132.0, 200.0), Insert.MoveKind.Absolute);

            ////en página de "Control Inputs I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Inputs I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 9, oProject.Pages[key], new PointD(160.0, 160.0), Insert.MoveKind.Absolute);

            //changeFunctionTextPLCInput(oProject.Pages[key], "I4", "Lightbarrier");

            ////en página de "Safety Inputs II"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Safety Inputs II");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 10, oProject.Pages[key], new PointD(256.0, 164.0), Insert.MoveKind.Absolute);

            //changeFunctionTextPLCInput(oProject.Pages[key], "SI27", "Top Up Key Order");
            //changeFunctionTextPLCInput(oProject.Pages[key], "SI28", "Top Up Down Order");

            ////en página de "PLC Input I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "PLC Input I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 3, oProject.Pages[key], new PointD(124.0, 176.0), Insert.MoveKind.Absolute);

            ////en página de "PLC Input I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "PLC Input I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 2, oProject.Pages[key], new PointD(84.0, 128.0), Insert.MoveKind.Absolute);

            ////en página de "Upper Diagnostic Inputs III"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Inputs III");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 6, oProject.Pages[key], new PointD(260.0, 212.0), Insert.MoveKind.Absolute);

            ////en página de "Lower Diagnostic Inputs III"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Inputs III");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 7, oProject.Pages[key], new PointD(260.0, 216.0), Insert.MoveKind.Absolute);

            ////Despues de la página de "Upper People Detection"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper People Detection");

            //Renumeramos páginas
            //RenumeraPaginas();

            //pagePropList = new PagePropertyList();

            //for (int i = oPages.Length - 1; i > key; i--)
            //{
            //    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
            //    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
            //    pagePropList.PAGE_COUNTER = i + 2;
            //    oPages[i].SetName(pagePropList);
            //}

            //oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper People Detection II.emp", oPages[key], oProject, false);

            //GetPageTable();

            ////en página de "Upper People Detection II"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper People Detection II");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 4, oProject.Pages[key], new PointD(52.0, 184.0), Insert.MoveKind.Absolute);

            ////en página de "Upper Diagnostic Outputs I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Upper Diagnostic Outputs I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 11, oProject.Pages[key], new PointD(24.0, 196.0), Insert.MoveKind.Absolute);

            ////Despues de la página de "Lower People Detection"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower People Detection");

            ////Renumeramos páginas
            //RenumeraPaginas();

            //pagePropList = new PagePropertyList();

            //for (int i = oPages.Length - 1; i > key; i--)
            //{
            //    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
            //    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
            //    pagePropList.PAGE_COUNTER = i + 2;
            //    oPages[i].SetName(pagePropList);
            //}

            //oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower People Detection II.emp", oPages[key], oProject, false);

            //GetPageTable();

            ////en página de "Lower People Detection II"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower People Detection II");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 5, oProject.Pages[key], new PointD(52.0, 184.0), Insert.MoveKind.Absolute);

            ////en página de "Lower Diagnostic Outputs I"
            //key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Lower Diagnostic Outputs I");
            //oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 12, oProject.Pages[key], new PointD(24.0, 196.0), Insert.MoveKind.Absolute);
        }

        public void Draw_PLC()
        {
            int key;
            Insert oInsert = new Insert();

            //Despues de la página de "Control Outputs III"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control Outputs III");

            //Renumeramos páginas
            RenumeraPaginas();

            PagePropertyList pagePropList = new PagePropertyList();

            for (int i = oPages.Length - 1; i > key; i--)
            {
                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                pagePropList.PAGE_COUNTER = i + 2;
                oPages[i].SetName(pagePropList);
            }

            //PLC Input
            oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\PLC_Input_I.emp", oPages[key], oProject, false);

            GetPageTable();

            //Despues de la página de "PLC Input I"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "PLC Input I");

            //Renumeramos páginas
            RenumeraPaginas();

            pagePropList = new PagePropertyList();

            for (int i = oPages.Length - 1; i > key; i--)
            {
                pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                pagePropList.PAGE_COUNTER = i + 2;
                oPages[i].SetName(pagePropList);
            }

            //PLC Output
            oInsert.PageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\PLC_Output_I.emp", oPages[key], oProject, false);

            GetPageTable();

            //en página de "Display"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Display");
            oInsert.WindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\PLC.ema", 0, oProject.Pages[key], new PointD(216.0, 252.0), Insert.MoveKind.Absolute);

            //en página de "Communication"
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\PLC.ema", 'B', "Communication", 68.0, 180.0);


        }

        public void drawfrenchpakage()
        {
            Caracteristic Pais = (Caracteristic)oElectric.CaractComercial["FLAND"];
            Caracteristic Producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            if (Pais.CurrentReference.Equals("1") ||
                Pais.CurrentReference.Equals("2"))
            {
                //en página de "Main Power Supply"
                int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Main Power Supply");

                Placement[] placements = oProject.Pages[key].AllPlacements;

                foreach (Placement placement in placements)
                {
                    if (placement.Location.Y >= 228 &&
                        placement.Location.Y <= 240 &&
                        placement.Location.X >= 56 &&
                        placement.Location.X <= 120)
                    {
                        placement.Remove();
                    }

                }

                placements = oProject.Pages[key].AllPlacements;

                foreach (Placement placement in placements)
                {
                    if (placement.Location.Y >= 208 &&
                        placement.Location.Y <= 220 &&
                        placement.Location.X >= 244 &&
                        placement.Location.X <= 280)
                    {
                        placement.Remove();
                    }
                }
                placements = oProject.Pages[key].AllPlacements;

                foreach (Placement placement in placements)
                {
                    if (placement.Location.Y >= 176 &&
                        placement.Location.Y <= 192 &&
                        placement.Location.X >= 84 &&
                        placement.Location.X <= 276)
                    {
                        placement.Remove();
                    }
                }

                insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'A', "Main Power Supply", 60.0, 264.0);

                if (Producto.CurrentReference.Equals("ORINOCO"))
                {
                    insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper People Detection II.emp", "Upper People Detection", "Upper People Detection II");
                    insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'C', "Upper People Detection II", 156.0, 240.0);
                    insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'B', "Upper Diagnostic Outputs I", 52.0, 124.0);
                    insertPageMacro("$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower People Detection II.emp", "Lower People Detection", "Lower People Detection II");
                    insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'E', "Lower People Detection II", 156.0, 240.0);
                    insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'D', "Lower Diagnostic Outputs I", 52.0, 124.0);
                }

                log = String.Concat(log, "\nIncluido Paquete Francia");

            }
        }

        public void drawMercadonapakage()
        {
            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Emergency open door.ema", 'B', "Upper Diagnostic Inputs III", 148, 200);

            changeFunctionTextPLCInput("Upper Diagnostic Inputs III", "UI17", "shutter rolling door (SS)");


            //en página de "Main Power Supply"
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control I");
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {

                if (placement.Location.X > 160 && placement.Location.X < 392)
                {
                    placement.Location = new PointD(placement.Location.X + 44, placement.Location.Y);
                }

            }

            placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.X == 72 && placement.Location.Y == 212)
                {
                    placement.Remove();
                }
            }

            placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {

                if (placement.Location.X == 88 && placement.Location.Y == 204)
                {
                    placement.Remove();
                }
            }

            insertWindowMacro("$(MD_MACROS)\\_Esquema\\2_Ventana\\Mercadona_Package.ema", 'A', "Control I", 72, 176);


        }

        #endregion

        #region auxiliary functions

        public void insertNewPage(string pageName, string pageBefore)
        {
            //Compruebo si ya esta insertada
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageName);

            if (key == 0)
            {
                //Despues de la página de
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageBefore);

                //Renumeramos páginas
                RenumeraPaginas();

                PagePropertyList pagePropList = new PagePropertyList();

                for (int i = oProject.Pages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oProject.Pages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oProject.Pages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oProject.Pages[i].SetName(pagePropList);
                }

                oProject.Pages[key].Properties.CopyTo(pagePropList);
                pagePropList.PAGE_COUNTER = oProject.Pages[key].Properties.PAGE_COUNTER + 1;

                Page page = new Page();
                page.Create(oProject, DocumentTypeManager.DocumentType.Circuit, pagePropList);

                oProject.Pages[key + 1].Properties.PAGE_NOMINATIOMN = oProject.Pages[key].Properties.PAGE_NOMINATIOMN;
                PropertyValue pageNameProperty = oProject.Pages[key + 1].Properties.PAGE_NOMINATIOMN;
                MultiLangString langString = new MultiLangString();
                langString = pageNameProperty.ToMultiLangString();
                langString.AddString(ISOCode.Language.L_en_US, pageName);
                pageNameProperty.Set(langString);

                GetPageTable();

            }

        }

        public void changeFunctionTextPLCInput(Page page, string address, string newText)
        {
            page.Filter.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.PLCTerminal;

            //now we have all functions having category 'MOTOR' placed on page p
            Function[] functions = page.Functions;

            //other way to do the same:
            FunctionsFilter ff = new FunctionsFilter();
            ff.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.PLCTerminal;
            ff.Page = page;
            DMObjectsFinder objFinder = new DMObjectsFinder(oProject);

            //now we have all functions having category 'MOTOR' placed on page p
            functions = objFinder.GetFunctions(ff);

            foreach (Function f in functions)
            {
                PropertyValue PLCAdress = f.Properties.FUNC_GEDNAMEWITHPLCADRESS;
                if (PLCAdress.ToString().Equals(address))
                {
                    PropertyValue functionText = f.Properties.FUNC_TEXT;
                    MultiLangString langString = new MultiLangString();
                    langString.AddString(ISOCode.Language.L_en_US, newText);
                    functionText.Set(langString);
                }
            }
        }

        public void changeFunctionTextPLCInput(string spage, string address, string newText)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "External Feed Wiring"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == spage);

            Page page = oProject.Pages[key];

            page.Filter.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.PLCTerminal;

            //now we have all functions having category 'MOTOR' placed on page p
            Function[] functions = page.Functions;

            //other way to do the same:
            FunctionsFilter ff = new FunctionsFilter();
            ff.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.PLCTerminal;
            ff.Page = page;
            DMObjectsFinder objFinder = new DMObjectsFinder(oProject);

            //now we have all functions having category 'MOTOR' placed on page p
            functions = objFinder.GetFunctions(ff);

            foreach (Function f in functions)
            {
                PropertyValue PLCAdress = f.Properties.FUNC_GEDNAMEWITHPLCADRESS;
                if (PLCAdress.ToString().Equals(address))
                {
                    PropertyValue functionText = f.Properties.FUNC_TEXT;
                    MultiLangString langString = new MultiLangString();
                    langString.AddString(ISOCode.Language.L_en_US, newText);
                    functionText.Set(langString);
                }
            }
        }

        public void insertWindowMacro(string pathMacro, char variante, string page, double x, double y)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "External Feed Wiring"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);

            int nVariante = -1;
            switch (variante)
            {
                case 'A':
                    nVariante = 0;
                    break;

                case 'B':
                    nVariante = 1;
                    break;

                case 'C':
                    nVariante = 2;
                    break;

                case 'D':
                    nVariante = 3;
                    break;

                case 'E':
                    nVariante = 4;
                    break;

                case 'F':
                    nVariante = 5;
                    break;

                case 'G':
                    nVariante = 6;
                    break;

                case 'H':
                    nVariante = 7;
                    break;

                case 'I':
                    nVariante = 8;
                    break;

                case 'J':
                    nVariante = 9;
                    break;

                case 'K':
                    nVariante = 10;
                    break;

                case 'L':
                    nVariante = 11;
                    break;

                case 'M':
                    nVariante = 12;
                    break;

                case 'N':
                    nVariante = 13;
                    break;

                case 'O':
                    nVariante = 14;
                    break;

                case 'P':
                    nVariante = 15;
                    break;

            }

            oInsert.WindowMacro(pathMacro, nVariante, oProject.Pages[key], new PointD(x, y), Insert.MoveKind.Absolute);
        }

        public StorableObject[] insertWindowMacro_ObjCont(string pathMacro, char variante, string page, double x, double y)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "External Feed Wiring"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);

            int nVariante = -1;
            switch (variante)
            {
                case 'A':
                    nVariante = 0;
                    break;

                case 'B':
                    nVariante = 1;
                    break;

                case 'C':
                    nVariante = 2;
                    break;

                case 'D':
                    nVariante = 3;
                    break;

                case 'E':
                    nVariante = 4;
                    break;

                case 'F':
                    nVariante = 5;
                    break;

                case 'G':
                    nVariante = 6;
                    break;

                case 'H':
                    nVariante = 7;
                    break;

                case 'I':
                    nVariante = 8;
                    break;

                case 'J':
                    nVariante = 9;
                    break;

                case 'K':
                    nVariante = 10;
                    break;

                case 'L':
                    nVariante = 11;
                    break;

                case 'M':
                    nVariante = 12;
                    break;

                case 'N':
                    nVariante = 13;
                    break;

                case 'O':
                    nVariante = 14;
                    break;

                case 'P':
                    nVariante = 15;
                    break;

            }

            StorableObject[] oInsertedObjects = oInsert.WindowMacro(pathMacro, nVariante, oProject.Pages[key], new PointD(x, y), Insert.MoveKind.Absolute);

            return oInsertedObjects;
        }

        public void insertPageMacro(string pageMacroPath, string pageBefore, string pageName)
        {
            int key;
            Insert oInsert = new Insert();
            PagePropertyList pagePropList;

            //Compruebo si ya esta insertada la página
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageName);

            if (key == 0)
            {
                //Despues de la página de "Control II"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageBefore);

                Renumber renumber = new Renumber();
                renumber.Pages(oPages, false, 1, 1, false, false, Renumber.Enums.SubPages.ConsecutiveNumbering);

                pagePropList = new PagePropertyList();

                for (int i = oPages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oPages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oPages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oPages[i].SetName(pagePropList);
                }

                oInsert.PageMacro(pageMacroPath, oPages[key], oProject, false);

                GetPageTable();

            }
        }

        public void insertLayoutMacro(string pathMacro, char variante, string page, double x, double y, string IME)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "External Feed Wiring"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);

            int nVariante = -1;
            switch (variante)
            {
                case 'A':
                    nVariante = 0;
                    break;

                case 'B':
                    nVariante = 1;
                    break;

                case 'C':
                    nVariante = 2;
                    break;

                case 'D':
                    nVariante = 3;
                    break;

                case 'E':
                    nVariante = 4;
                    break;

                case 'F':
                    nVariante = 5;
                    break;

                case 'G':
                    nVariante = 6;
                    break;

                case 'H':
                    nVariante = 7;
                    break;

                case 'I':
                    nVariante = 8;
                    break;

                case 'J':
                    nVariante = 9;
                    break;

                case 'K':
                    nVariante = 10;
                    break;

                case 'L':
                    nVariante = 11;
                    break;

                case 'M':
                    nVariante = 12;
                    break;

                case 'N':
                    nVariante = 13;
                    break;

                case 'O':
                    nVariante = 14;
                    break;

                case 'P':
                    nVariante = 15;
                    break;

            }

            StorableObject[] oInsertedObjects = oInsert.WindowMacro(pathMacro, nVariante, oProject.Pages[key], new PointD(x, y), Insert.MoveKind.Absolute);

            Function f = oInsertedObjects[0] as Function;

            f.Name = String.Concat(f.Name.Split('-')[0], "-", IME);

        }

        #endregion
    }
}
