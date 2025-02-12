using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.EObjects;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.HEServices;
using EPLAN_API.SAP;
using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media.Media3D;
using static EPLAN_API.SAP.Electric;

namespace EPLAN_API.User
{
    public class DrawArmStandardIntEN : DrawTools
    {
        private Project oProject;
        private Electric oElectric;
        private string log;
        private uint completedPercentaje = 0;
        public delegate void ProgressChangedDelegate(int value);
        public event ProgressChangedDelegate ProgressChangedToDraw;

        public DrawArmStandardIntEN(Project project, Electric electric) 
        {
            oProject = project;
            oElectric = electric;
            
        }

        public void DrawMacro()
        {
            Caracteristic c, c2, c3, c4;
            String refVal, refVal2, refVal3, refVal4;

            int progress = 0;
            int step = 100/36;

            //Contactores
            Draw_Default_Param();
            progress += step;
            ProgressChanged(progress);

            //Sensores de Freno
            Draw_Freno();
            progress += step;
            ProgressChanged(progress);

            //Sensores sincronismo
            Draw_Sincronismo();
            progress += step;
            ProgressChanged(progress);

            //Display
            Draw_Display();
            progress += step;
           ProgressChanged(progress);;


            //Termico Motor
            c = (Caracteristic)oElectric.CaractComercial["FANTREHT"];
            Draw_Termico(c.CurrentReference);
            progress += step;
           ProgressChanged(progress);;

            //PLC
            c = (Caracteristic)oElectric.CaractIng["TNCR_DO_CONTROL"];
            if (c.CurrentReference.Equals("GEC+PLC"))
            {
                Draw_PLC();
                log = String.Concat(log, "\nIncluido PLC");
            }
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

            //Micros de Zócalo
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            if (c.NumVal >= 4)
            {
                Draw_Micros_Zocalo(Convert.ToInt16(c.NumVal));
                log = String.Concat(log, "\nIncluidos ", Convert.ToInt16(c.NumVal).ToString(), " micros de zócalo");
            }
            progress += step;
           ProgressChanged(progress);;

            //Contacto de fuego
            c = (Caracteristic)oElectric.CaractComercial["TNCR_CONTACTO_FUEGO"];
            if (c.CurrentReference.Equals("S"))
            {
                Draw_Contacto_Fuego();
                log = String.Concat(log, "\nIncluido contacto de fuego");
            }
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;


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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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

            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress += step;
           ProgressChanged(progress);;

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
            progress = 100;
            ProgressChanged(progress); ;

            Reports report = new Reports();
            report.GenerateProject(oProject);

            //Redraw
            Edit edit = new Edit();
            edit.RedrawGed();

            paramGEC(oProject, oElectric);
            writeGECtoEPLAN(oProject, oElectric);

            MessageBox.Show(new Form() { TopMost = true, TopLevel = true }, log, "Resultado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ProgressChanged(0);
            String path = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "Log\\Log_Draw.txt");
            File.WriteAllText(path, log);


        }

        private void ProgressChanged(int progress)
        {
           ProgressChangedToDraw(progress);
        }


        #region Metodos de dibujo

        private void Draw_Default_Param()
        {
            //UI23    UDL1 Standard input 23 X23
            SetGECParameter(oProject, oElectric, "UI23", (uint)GEC.Param.Top_up_key_order);

            //UI24    UDL1 Standard input 24
            SetGECParameter(oProject, oElectric, "UI24", (uint)GEC.Param.Top_down_key_order);

            //LI23    LDL1 Standard input 23 X23
            SetGECParameter(oProject, oElectric, "LI23", (uint)GEC.Param.Bottom_up_key_order);

            //LI24    LDL1 Standard input 24
            SetGECParameter(oProject, oElectric, "LI24", (uint)GEC.Param.Bottom_down_key_order);

            //UO4 UDL1 Relay output 1 NO Q4/ 4L
            SetGECParameter(oProject, oElectric, "UO4", (uint)GEC.Param.Up_indication);

            //UO5 UDL1 Relay output 2NO Q5/ 5L
            SetGECParameter(oProject, oElectric, "UO5", (uint)GEC.Param.Down_indication);

            //UO6 UDL1 Relay output 3 NO Q6/ 6L
            SetGECParameter(oProject, oElectric, "UO6", (uint)GEC.Param.Oil_pump_activation);

            //UO7 UDL1 Relay output 4 NO Q7/ 7L
            SetGECParameter(oProject, oElectric, "UO7", (uint)GEC.Param.Fault_bit_60_63);

            //UO8 UDL1 Relay output 5 NO Q8/ 8L
            SetGECParameter(oProject, oElectric, "UO8", (uint)GEC.Param.Maintenance_indication);

            //SI23 SF Safety Input 15 X23
            SetGECParameter(oProject, oElectric, "SI23", (uint)GEC.Param.Top_open_floor_plate_1);

            //SI24    SF Safety Input 16
            SetGECParameter(oProject, oElectric, "SI24", (uint)GEC.Param.Top_open_floor_plate_2);

            //SI25 SF Safety Input 17
            SetGECParameter(oProject, oElectric, "SI25", (uint)GEC.Param.Bottom_open_floor_plate_1);

            //SI26 SF Safety Input 18
            SetGECParameter(oProject, oElectric, "SI26", (uint)GEC.Param.Bottom_open_floor_plate_2);

            //Conexion Motor "VVF_YD"
            SetGECParameter(oProject, oElectric, "SI12", (uint)GEC.Param.Contactor_FB_2, true);
            SetGECParameter(oProject, oElectric, "SI21", (uint)GEC.Param.Contactor_FB_3, true);
            SetGECParameter(oProject, oElectric, "SI22", (uint)GEC.Param.Contactor_FB_4, true);
            SetGECParameter(oProject, oElectric, "O3", (uint)GEC.Param.Main_2, true);
            SetGECParameter(oProject, oElectric, "O4", (uint)GEC.Param.Main_1, true);
        }
       
        private void Draw_Freno()
        {
            string unidadAccionamiento = ((Caracteristic)oElectric.CaractComercial["FANTREHT"]).CurrentReference;

            insertNewPage(oProject, "Brake Sensors", "Motor Sensors");

            log = "Incluidos sensores de freno: ";

            if (unidadAccionamiento.Equals("QC") ||
                unidadAccionamiento.Equals("FJ"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200002262200_00_BRAKE_ELECTRICAL_ASSEMBLY_UE.ema", 'P', "Brake Sensors", 44, 244);
                log = String.Concat(log, "Finales de carrera");
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\200002262200_00_BRAKE_ELECTRICAL_ASSEMBLY_UE.ema", 'O', "Brake Sensors", 44, 244);
                log = String.Concat(log, "Inductivos");
            }

        }

        private void Draw_Sincronismo()
        {
            string Modelo = (oElectric.CaractComercial["FMODELL"] as Caracteristic).CurrentReference;

            //Sensores superiores
            if (Modelo.Contains("CLASSIC"))
            {
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_I_VC3_0.emp", "Upper Diagnostic Outputs I", "Upper Sensors I");
            }
            else
            {
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_I_VE.emp", "Upper Diagnostic Outputs I", "Upper Sensors I");
            }

            //Sensores superiores
            if (Modelo.Contains("CLASSIC"))
            {
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_I_VC3_0.emp", "Lower Diagnostic Outputs I", "Lower Sensors I");
            }
            else
            {
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_I_VE.emp", "Lower Diagnostic Outputs I", "Lower Sensors I");
            }

        }

        private void Draw_Display()
        {
            string tipoDisplay = (oElectric.CaractIng["TNCR_DO_DISPLAY_TYPE"] as Caracteristic).CurrentReference;

            if (tipoDisplay.Equals("DDU"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'B', "Display", 272.0, 252.0);
            else if (tipoDisplay.Equals("ESCATRONIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'A', "Display", 272.0, 252.0); ;
        }

        private void Draw_Contacto_Fuego()
        {
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Fire_Contact.ema", 'C', "Control Inputs I", 224, 268);
            SetGECParameter(oProject, oElectric, "I6", (uint)GEC.Param.Fire_alarm_smoke_detector_1, true);
        }

        private void Draw_Llavin_Auto_Cont()
        {
            //Upper Keys"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'A', "Upper Keys", 24.0, 256.0);

            //Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'B', "Upper Diagnostic Inputs IV", 220.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI25", (uint)GEC.Param.Continuous_key_top, true);
            SetGECParameter(oProject, oElectric, "UI26", (uint)GEC.Param.Automatic_key_top, true);
        }

        private void Draw_Llavin_Local_Remoto()
        {

            //Upper Keys"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'C', "Upper Keys", 300.0, 256.0);


            //Upper Diagnostic Inputs IV
            //Insertamos página
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Diagnostic_Inputs_V.emp", "Upper Diagnostic Inputs IV", "Upper Diagnostic Inputs V");


            //Upper Diagnostic Inputs V"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'D', "Upper Diagnostic Inputs V", 248.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI31", (uint)GEC.Param.Local_key_top, true);
            SetGECParameter(oProject, oElectric, "UI32", (uint)GEC.Param.Remote_key_top, true);
        }

        private void Draw_LuzEstroboscopica(string type)
        {
            SetGECParameter(oProject, oElectric, "O1", (uint)GEC.Param.Lighting_1, true);
            switch (type)
            {
                //Una tira LED
                case "LED":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'F', "Upper Lighting I", 236.0, 140.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'A', "Lower Lighting I", 168.0, 108.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 1 Tira LED");
                    break;

                //Dos tiras LED
                case "2LED":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'G', "Upper Lighting I", 236.0, 140.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'B', "Lower Lighting I", 168.0, 108.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 2 Tira LED");
                    break;

                //Tres tiras LED
                case "3LED":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'H', "Upper Lighting I", 236.0, 140.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'C', "Lower Lighting I", 168.0, 108.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 3 Tira LED");
                    break;

                //Una lampara
                case "STFSPALTBEL":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'I', "Upper Lighting I", 32.0, 176.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'D', "Lower Lighting I", 168.0, 136.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 1 Lampara");
                    break;

                //Dos lamparas
                case "STFSPALTBEL2":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'J', "Upper Lighting I", 32.0, 176.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'E', "Lower Lighting I", 168.0, 136.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 2 Lamparas");
                    break;
            }
        }

        private void Draw_LuzPeines(string type, string luzEstro)
        {
            SetGECParameter(oProject, oElectric, "O1", (uint)GEC.Param.Lighting_1, true);
            if (type.Equals("DI"))
            {
                //en página de "Upper Lighting I"
                if (luzEstro.Equals("STFSPALTBEL") || luzEstro.Equals("STFSPALTBEL2"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'A', "Upper Lighting I", 80.0, 176.0);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'B', "Upper Lighting I", 32.0, 176.0);

                //en página de "Lower Lighting I"
                if (luzEstro.Equals("STFSPALTBEL") || luzEstro.Equals("STFSPALTBEL2"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'D', "Lower Lighting I", 176.0, 180.0);
                else if (luzEstro.Contains("LED"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'E', "Lower Lighting I", 176.0, 180.0);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'C', "Lower Lighting I", 176.0, 180.0);

                log = String.Concat(log, "\nIncluida luz de peines LED 230V");
            }
        }

        private void Draw_LuzBajopasamanos(string type, string luzEstro, string luzPeines)
        {
            SetGECParameter(oProject, oElectric, "O1", (uint)GEC.Param.Lighting_1, true);
            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'B', "Lower Lighting I", 156.0, 112.0);
                }
                else
                {                  
                    insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                   

                    //en página de "Lower Lighting II"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'A', "Lower Lighting II", 40.0, 188.0);

                    if (!luzEstro.Equals("KEINE") || !luzPeines.Equals("NO"))
                    {
                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'C', "Lower Lighting I", 176, 200.0);
                    }

                }

                log = String.Concat(log, "\nIncluida luz bajpasamanos LED 230V");
            }

        }

        private void Draw_LuzZocalos(string type, string luzBajopasamanos, string luzEstro, string luzPeines)
        {
            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            SetGECParameter(oProject, oElectric, "O1", (uint)GEC.Param.Lighting_1, true);

            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (modelo.CurrentReference.Contains("CLASSIC"))
                {
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'E', "Lower Lighting I", 176, 196.0);
                    insertPageMacro(oProject,"$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'I', "Lower Lighting II", 36.0, 272.0);

                }
                else
                {
                    if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && luzBajopasamanos.Equals("BELOHNE"))
                    {
                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'A', "Lower Lighting I", 156.0, 112.0);
                    }
                    else if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && !luzBajopasamanos.Equals("BELOHNE"))
                    {
                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'B', "Lower Lighting I", 228.0, 172.0);
                    }
                    else
                    {
                        if (!luzBajopasamanos.Equals("BELOHNE"))
                        {
                            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                            //en página de "Lower Lighting II"
                            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'C', "Lower Lighting II", 132.0, 188.0);
                        }
                        else
                        {
                            //Despues de la página de "Lower Lighting I"
                            //Lower Lighting
                            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                            //en página de "Lower Lighting II"
                            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'D', "Lower Lighting II", 40.0, 188.0);

                            //en página de "Lower Lighting I"
                            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'E', "Lower Lighting I", 176, 200.0);
                        }
                    }
                }

                log = String.Concat(log, "\nIncluida luz zócalo LED 230V");
            }

        }

        private void Draw_LuzFoso(string type)
        {
            if (type.Equals("HANDL") ||
                type.Equals("OVAL"))
            {
                //en página de "Main Power Supply"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 'A', "Main Power Supply", 308.0, 60.0);

                //en página de "LLower Lighting I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 'B', "Lower Lighting I", 24.0, 172.0);

                log = String.Concat(log, "\nIncluida luz de foso manual");
            }
        }

        private void Draw_Semaforo(string type)
        {

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Semaforos superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Traffic_Lights.emp", "Upper Keys", "Upper Traffic Lights");

            //Semaforos inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Traffic_Lights.emp", "Lower Keys", "Lower Traffic Lights");

            SetGECParameter(oProject, oElectric, "UO1", (uint)GEC.Param.Top_traffic_light_red, true);
            SetGECParameter(oProject, oElectric, "UO2", (uint)GEC.Param.Top_traffic_light_green, true);
            SetGECParameter(oProject, oElectric, "LO1", (uint)GEC.Param.Bottom_traffic_light_red, true);
            SetGECParameter(oProject, oElectric, "LO2", (uint)GEC.Param.Bottom_traffic_light_green, true);


            if (modelo.CurrentReference.Contains("CLASSIC"))
            {
                //en página de "Upper Traffic Lights"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'M', "Upper Traffic Lights", 128.0, 148.0);
                //en página de "Lower Traffic Lights"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'N', "Lower Traffic Lights", 128.0, 148.0);
                log = String.Concat(log, "\nIncluidos Semáforos Chinos de VC3.0");
            }
            else
            {
                if (type.Equals("BICOLOR"))
                {
                    //en página de "Upper Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'A', "Upper Traffic Lights", 128.0, 148.0);

                    //en página de "Lower Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'C', "Lower Traffic Lights", 128.0, 148.0);

                    log = String.Concat(log, "\nIncluidos Semáforos Bicolor");
                }
                else if (type.Equals("ROTGRUEN") ||
                        type.Equals("F6NEINB") ||
                        type.Equals("PROHI_VERDE"))
                {
                    //en página de "Upper Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'B', "Upper Traffic Lights", 128.0, 148.0);

                    //en página de "Lower Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'D', "Lower Traffic Lights", 128.0, 148.0);

                    log = String.Concat(log, "\nIncluidos Semáforos Flecha/Prohibido");
                }
                else
                {
                    log = String.Concat(log, "\nNo existe macro para el tipo de semaforo seleccionado");
                }
            }


            //en página de "Upper Diagnostic Outputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'P', "Upper Diagnostic Outputs I", 36.0, 124.0);

            //en página de "Lower Diagnostic Outputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'O', "Lower Diagnostic Outputs I", 36.0, 124.0);



        }

        private void Draw_Radar()
        {
            Caracteristic producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Radares Superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", "Upper Diagnostic Outputs I", "Upper People Detection");

            //en página de "Upper People Detection"
            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'I', "Upper People Detection", 200.0, 256.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'A', "Upper People Detection", 200.0, 256.0);


            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'P', "Upper Diagnostic Inputs III", 280.0, 168.0);

            SetGECParameter(oProject, oElectric, "UI19", (uint)GEC.Param.Top_radar_right_NC, true);
            SetGECParameter(oProject, oElectric, "UI20", (uint)GEC.Param.Top_radar_left_NC, true);

            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Radares Inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");

            //en página de "Lower People Detection"
            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'J', "Lower People Detection", 200.0, 256.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'B', "Lower People Detection", 200.0, 256.0);


            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'O', "Lower Diagnostic Inputs III", 280.0, 168.0);
            SetGECParameter(oProject, oElectric, "LI19", (uint)GEC.Param.Bottom_radar_right_NC, true);
            SetGECParameter(oProject, oElectric, "LI20", (uint)GEC.Param.Bottom_radar_left_NC, true);
        }

        private void Draw_Fotocelulas()
        {
            Caracteristic producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Fotocélulas superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", "Upper Diagnostic Outputs I", "Upper People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'H', "Upper People Detection", 12.0, 256.0);
            else
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'B', "Upper People Detection", 12.0, 256.0);


            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'P', "Upper Diagnostic Inputs III", 400.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI21", (uint)GEC.Param.Top_light_barrier_comb_plate_NC, true);

            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Fotocélulas inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'G', "Lower People Detection", 12.0, 256.0);
            else
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'A', "Lower People Detection", 12.0, 256.0);


            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'O', "Lower Diagnostic Inputs III", 400.0, 168.0);
            SetGECParameter(oProject, oElectric, "LI21", (uint)GEC.Param.Bottom_light_barrier_comb_plate_NC, true);
        }

        private void Draw_Trinquete_Magnetico(string type)
        {
            //en página de "Motor Sensors"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'D', "Motor Sensors", 320.0, 260.0);

            //en página de "Control I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'H', "Control I", 332.0, 212.0);


            //en página de "Safety Pulse Inputs"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'E', "Safety Pulse Inputs",288.0, 124.0);
            SetGECParameter(oProject, oElectric, "SI7", (uint)GEC.Param.Main_shaft_speed_monitor_1, true);
            SetGECParameter(oProject, oElectric, "SI8", (uint)GEC.Param.Main_shaft_speed_monitor_2, true);

            if (type.Equals("HWSPERRKMAGN"))
            {
                //en página de "Brake Sensors"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'A', "Brake Sensors", 336.0, 244.0);
            }
            else if (type.Equals("NAB"))
            {
                //en página de "Brake Sensors"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'B', "Brake Sensors", 336.0, 244.0);

            }

            //en página de "Safety Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'C', "Safety Inputs II", 256.0, 164.0);
            SetGECParameter(oProject, oElectric, "SI27", (uint)GEC.Param.Aux_brake_status_1, true);

            //en página de "Control Outputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'D', "Control Outputs III", 28.0, 104.0);

            //en página de "Control I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'E', "Control I", 20.0, 192.0);

            //quitamos puntos de conexión
            deleteArea(oProject, "Control I", 72, 204, 88, 212);

            //SAI
            //Compruebo si ya esta insertada la página "Control II"
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Control_II_Arm_Int.emp", "Control I", "Control II");

        }

        private void Draw_Cerrojo()
        {
            Caracteristic sStopAdicional = (Caracteristic)oElectric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            Caracteristic sStopCarritos = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic sMicrosZocalo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            Caracteristic sPeines = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];

            if (sStopCarritos.CurrentReference.Equals("KEINE") || sStopAdicional.CurrentReference.Equals("N"))
            {
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI17", (uint)GEC.Param.Chain_locking_device_SS, true);
            }
            else if (sMicrosZocalo.NumVal <= 6)
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs III", 76.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI16", (uint)GEC.Param.Chain_locking_device_SS, true);

            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {

                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs II", 372.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI14", (uint)GEC.Param.Chain_locking_device_SS, true);

            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs II", 256.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI12", (uint)GEC.Param.Chain_locking_device_SS, true);
            }
        }

        private void Draw_Trinquete_Mecanico()
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
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI17", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);
            }
            else if (sMicrosZocalo.NumVal <= 6 &&
                cerrojo.CurrentReference.Equals("N"))
            {
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 76.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI16", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);

            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {

                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 16.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI15", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);

            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs II", 312.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI13", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);
            }
        }

        private void Draw_RoturaCadenaPrincipal()
        {

            //en página de "Motor Sensors"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 'A', "Motor Sensors", 208.0, 260.0);


            //en página de "Safety Inputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 'P', "Safety Inputs I", 292.0, 184.0);
            SetGECParameter(oProject, oElectric, "SI17", (uint)GEC.Param.Drive_chain_DuTriplex, true);
        }

        private void Draw_RoturaPasamanos()
        {

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_II.emp", "Lower Sensors I", "Lower Sensors II");

            //en página de "Lower Sensors II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'A', "Lower Sensors II", 36.0, 260.0);


            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'P', "Lower Diagnostic Inputs IV", 216.0, 164.0);
            SetGECParameter(oProject, oElectric, "LI25", (uint)GEC.Param.Broken_handrail_L, true);
            SetGECParameter(oProject, oElectric, "LI26", (uint)GEC.Param.Broken_handrail_R, true);
        }

        private void Draw_ControlDesgasteFrenos()
        {

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_II.emp", "Lower Sensors I", "Lower Sensors II");

            //en página de "Upper Sensors II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'A', "Upper Sensors II", 28.0, 264.0);


            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'P', "Upper Diagnostic Inputs IV", 340.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI27", (uint)GEC.Param.Brake_wear_brake_1_M1, true);
            SetGECParameter(oProject, oElectric, "UI28", (uint)GEC.Param.Brake_wear_brake_2_M1, true);
        }

        private void Draw_BuggyInferior()
        {

            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'B', "Lower Diagnostic Inputs III", 16.0, 156.0);
            SetGECParameter(oProject, oElectric, "LI15", (uint)GEC.Param.Bottom_buggy_right_SS, true);
            SetGECParameter(oProject, oElectric, "LI16", (uint)GEC.Param.Bottom_buggy_left_SS, true);
        }

        private void Draw_BuggySuperior()
        {
            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'M', "Upper Diagnostic Inputs II", 368.0, 156.0);

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'N', "Upper Diagnostic Inputs III", 16.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI14", (uint)GEC.Param.Top_buggy_left_SS, true);
            SetGECParameter(oProject, oElectric, "UI15", (uint)GEC.Param.Top_buggy_right_SS, true);
        }

        private void Draw_SeguridadVerticalPeines()
        {

            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 'A', "Upper Diagnostic Inputs II", 252.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI12", (uint)GEC.Param.Top_vertical_comb_plate_right_SS, true);
            SetGECParameter(oProject, oElectric, "UI13", (uint)GEC.Param.Top_vertical_comb_plate_left_SS, true);

            //en página de "Lower Diagnostic Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 'B', "Lower Diagnostic Inputs II", 312.0, 156.0);
            SetGECParameter(oProject, oElectric, "LI13", (uint)GEC.Param.Bottom_vertical_comb_plate_right_SS, true);
            SetGECParameter(oProject, oElectric, "LI14", (uint)GEC.Param.Bottom_vertical_comb_plate_left_SS, true);
        }

        private void Draw_StopAdicionalSuperior()
        {

            Caracteristic sStopCarritos = (Caracteristic)oElectric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic sMicrosZocalo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            Caracteristic sPeines = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];

            if (sStopCarritos.CurrentReference.Equals("KEINE"))
            {
                //en página de "Upper Diagnostic Inputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'A', "Upper Diagnostic Inputs II", 192.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI11", (uint)GEC.Param.Top_emergency_stop_external_SS, true);
            }
            else if (sMicrosZocalo.NumVal <= 6)
            {
                //en página de "Upper Diagnostic Inputs III"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI17", (uint)GEC.Param.Top_emergency_stop_external_SS, true);
            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {
                //en página de "Upper Diagnostic Inputs III"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'A', "Upper Diagnostic Inputs III", 16.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI15", (uint)GEC.Param.Top_emergency_stop_external_SS, true);
            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                //en página de "Upper Diagnostic Inputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'A', "Upper Diagnostic Inputs II", 312.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI13", (uint)GEC.Param.Top_emergency_stop_external_SS, true);
            }

        }

        private void Draw_StopCarritoSuperior()
        {
            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_Trolley.ema", 'A', "Upper Diagnostic Inputs II", 192.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI11", (uint)GEC.Param.Top_emergency_stop_trolley_SS, true);

        }

        private void Draw_StopAdicionalInferior()
        {

            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'B', "Lower Diagnostic Inputs III", 132.0, 156.0);
            SetGECParameter(oProject, oElectric, "LI17", (uint)GEC.Param.Bottom_emergency_stop_external_SS, true);
        }

        private void Draw_StopCarritoInferior()
        {

            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_Trolley.ema", 'B', "Lower Diagnostic Inputs III", 192.0, 156.0);
            SetGECParameter(oProject, oElectric, "LI18", (uint)GEC.Param.Bottom_emergency_stop_trolley_SS, true);
        }

        private void Draw_Micros_Zocalo(int nMicros)
        {
            if (nMicros >= 4)
            {
                //en página de "Upper Diagnostic Inputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Microswitch_1.ema", 'A', "Upper Diagnostic Inputs II", 72.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI9", (uint)GEC.Param.Top_skirt_right_SS, true);
                SetGECParameter(oProject, oElectric, "UI10", (uint)GEC.Param.Top_skirt_left_SS, true);

                //en página de "Lower Diagnostic Inputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Microswitch_1.ema", 'B', "Lower Diagnostic Inputs II", 192.0, 156.0);
                SetGECParameter(oProject, oElectric, "LI11", (uint)GEC.Param.Bottom_skirt_right_SS, true);
                SetGECParameter(oProject, oElectric, "LI12", (uint)GEC.Param.Bottom_skirt_left_SS, true);
            }

        }

        private void Draw_Termico(string motorType)
        {
            if (motorType.Equals("QC") ||
               motorType.Equals("FJ"))
            {
                //Termistor
                //en página de "Motor
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'D', "Motor", 296.0, 144.0);

                //en página de "Control Inputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'C', "Control Inputs I", 348.0, 268.0);

                //en página de "Control Inputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'E', "Control Inputs I", 96.0, 188.0);

                log = String.Concat(log, "\nIncluido relé termico");
            }
            else
            {
                //Bimetal
                //en página de "Motor"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'A', "Motor", 296.0, 112.0);

                //en página de "Control Inputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'B', "Control Inputs I", 96.0, 176.0);

                log = String.Concat(log, "\nIncluido bimetal");
            }

            SetGECParameter(oProject, oElectric, "I2", (uint)GEC.Param.Overtemperature_M1, true);
        }

        private void Draw_Freno_adicional(string motorType)
        {
            int key;
            Insert oInsert = new Insert();

            //en página de "Control I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'O', "Control I", 208.0, 160.0);

            //en página de "Oil pump & Brake"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'G', "Oil pump & Brake", 316.0, 84.0);

            if (motorType.Equals("QC") ||
                motorType.Equals("FJ"))
            {
                //Finales de carrera
                //en página de "Brake Sensors"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'B', "Brake Sensors", 188.0, 244.0);
            }
            else
            {
                //Inductivos
                //en página de "Brake Sensors"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'C', "Brake Sensors", 176.0, 248.0);
            }



            //en página de "Safety Inputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'L', "Safety Inputs I", 228.0, 184.0);
            SetGECParameter(oProject, oElectric, "SI15", (uint)GEC.Param.Brake_function_brake_3_mot_1, true);
            SetGECParameter(oProject, oElectric, "SI16", (uint)GEC.Param.Brake_function_brake_4_mot_1, true);

        }

        private void Draw_Lubricacion_auto()
        {

            //en página de "Oil pump & Brake"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'A', "Oil pump & Brake", 44.0, 176.0);

            //en página de "Control Outputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'P', "Control Outputs II", 276.0, 168.0);

            SetGECParameter(oProject, oElectric, "O9", (uint)GEC.Param.Oil_pump_activation, true);
            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            if (modelo.CurrentReference.Contains("CLASSIC"))
                SetGECParameter(oProject, oElectric, "O9", (uint)GEC.Param.Oil_pump_control_1, true);

            //en página de "Control I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'B', "Control I", 384.0, 212.0);

            //en página de "Control Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'O', "Control Inputs II", 128.0, 240.0);
            SetGECParameter(oProject, oElectric, "I11", (uint)GEC.Param.Oil_level_in_pump_1, true);


        }

        private void Draw_VVF(Double power)
        {

            //en página de "VVF Power"
            StorableObject[] oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'I', "VVF Power", 44.0, 208.0);

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
            deleteArea(oProject, "VVF Power", 92, 208, 116, 208);


            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\VVF_Control.emp", "VVF Power", "VVF Control");

            //en página de "Control Outputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'O', "Control Outputs I", 236.0, 128.0);
            SetGECParameter(oProject, oElectric, "O2", (uint)GEC.Param.Speed_selection_output_2, true);
            SetGECParameter(oProject, oElectric, "O3", (uint)GEC.Param.Main_2, true);
            SetGECParameter(oProject, oElectric, "O4", (uint)GEC.Param.Main_1, true);

            //en página de "Control Inputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'P', "Control Inputs I", 192.0, 160.0);
            SetGECParameter(oProject, oElectric, "I5", (uint)GEC.Param.VFD_EEC, true);
            SetGECParameter(oProject, oElectric, "I9", (uint)GEC.Param.Bypass_VFD, true);

        }

        private void Draw_NoVVF()
        {
            deleteArea(oProject, "Control Inputs II", 64, 184, 64, 184);
            //Se activa bypass aunque se elimine el selector
            SetGECParameter(oProject, oElectric, "I9", (uint)GEC.Param.Bypass_VFD, true);

        }

        private void Draw_DeleteBypass()
        {
            deleteArea(oProject, "Control Inputs II", 64, 184, 64, 268);
            
        }

        private void Draw_Safety_Curtain()
        {
            insertPageMacro(oProject,"$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Curtain.emp", "Display", "Safety Curtain");

            //en página de "Display"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'P', "Display", 124.0, 272.0);

            //en página de "Safety Curtain"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'A', "Safety Curtain", 52.0, 224.0);

            //en página de "Safety Curtain"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'B', "Safety Curtain", 52.0, 120.0);

            //en página de "Safety Curtain"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'O', "Safety Curtain", 16.0, 288.0);

            insertPageMacro(oProject,"$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Extension_Inputs.emp", "Safety Inputs II", "Safety Extension Inputs");
            SetGECParameter(oProject, oElectric, "SI44", (uint)GEC.Param.Safety_curtain, true);
            SetGECParameter(oProject, oElectric, "SI45", (uint)GEC.Param.Switch_disable_safety_curtain, true);
        }

        private void Draw_PLC()
        {
            //Despues de la página de "Control Outputs III"
            //PLC Input
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\PLC_Input_I.emp", "Control Outputs III", "PLC Input I");

            //PLC Output
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\PLC_Output_I.emp", "PLC Input I", "PLC Output I");


            //en página de "Display"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\PLC.ema", 'A', "Display", 216.0, 252.0);

            //en página de "Communication"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\PLC.ema", 'B', "Communication", 68.0, 180.0);


        }

        public void drawfrenchpakage()
        {
            Caracteristic Pais = (Caracteristic)oElectric.CaractComercial["FLAND"];
            Caracteristic Producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            if (Pais.CurrentReference.Equals("1") ||
                Pais.CurrentReference.Equals("2"))
            {
                //en página de "Main Power Supply"
                deleteArea(oProject, "Main Power Supply", 56, 228, 120, 240);
                deleteArea(oProject, "Main Power Supply", 244, 208, 280, 220);
                deleteArea(oProject, "Main Power Supply", 84, 176, 276, 192);

                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'A', "Main Power Supply", 60.0, 264.0);

                if (Producto.CurrentReference.Equals("ORINOCO"))
                {
                    insertPageMacro(oProject,"$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper People Detection II.emp", "Upper People Detection", "Upper People Detection II");
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'C', "Upper People Detection II", 156.0, 240.0);
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'B', "Upper Diagnostic Outputs I", 52.0, 124.0);
                    insertPageMacro(oProject,"$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower People Detection II.emp", "Lower People Detection", "Lower People Detection II");
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'E', "Lower People Detection II", 156.0, 240.0);
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\French_Package.ema", 'D', "Lower Diagnostic Outputs I", 52.0, 124.0);
                }

                log = String.Concat(log, "\nIncluido Paquete Francia");

            }
        }

        public void drawMercadonapakage()
        {
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Emergency open door.ema", 'B', "Upper Diagnostic Inputs III", 148, 200);
            SetGECParameter(oProject, oElectric, "UI17", (uint)GEC.Param.Shutter_rolling_door_SS, true);

            //en página de "Main Power Supply"
            moveSymbol(oProject, "Control I", 32, 0, 168, 32, 304, 244);

            deleteArea(oProject, "Control I", 72, 212, 72, 212);

            deleteArea(oProject, "Control I", 88, 204, 88, 204);

            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Mercadona_Package.ema", 'A', "Control I", 72, 176);

        }

        #endregion
    }

}
