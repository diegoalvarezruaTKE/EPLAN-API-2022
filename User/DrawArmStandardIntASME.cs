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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;

namespace EPLAN_API.User
{
    public class DrawArmStandardIntASME : DrawTools
    {
        private Project oProject;
        private Electric oElectric;
        private string log;
        public delegate void ProgressChangedDelegate(int value);
        public event ProgressChangedDelegate ProgressChangedToDraw;


        public DrawArmStandardIntASME(Project project, Electric electric)
        {
            oProject = project;
            oElectric = electric;
        }

        public void DrawMacro()
        {
            Caracteristic c, c2;
            String refVal, refVal2;

            int progress = 0;
            int step = 100 / 20;

            Draw_Default_Param();
            progress += step;
            ProgressChanged(progress);

            //Sensores de Freno
            c = (Caracteristic)oElectric.CaractComercial["FANTREHT"];
            Draw_Freno(c.CurrentReference);
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);


            //Freno Auxiliar
            c = (Caracteristic)oElectric.CaractComercial["FZUSBREMSE"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("NAB"))
                {
                    Draw_Freno_Aux();
                    log = String.Concat(log, "\nTrinquete NAB");
                }

                if (refVal.Equals("HWSPERRKMECH"))
                {
                    Draw_Trinquete_Mecanico();
                    log = String.Concat(log, "\nTrinquete Mecánico");
                }

            }
            progress += step;
            ProgressChanged(progress);


            //Sensores de sincronismo pasamanos
            c = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                Draw_Sincro_Pasamanos(refVal);
                log = String.Concat(log, "\nIncluido sincronismo de pasamanos");
            }
            progress += step;
            ProgressChanged(progress);


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
            ProgressChanged(progress);

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
                }
                else if (refVal.Equals("VA_REG"))
                {
                    c2 = (Caracteristic)oElectric.CaractComercial["TNPOTENCIAMOTOR"];
                    Draw_VVF_Regen(c2.NumVal);
                    log = String.Concat(log, "\nIncluido VVF Regenerativo");
                }
            }
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

            //Detección de personas por fotocélulas
            c = (Caracteristic)oElectric.CaractComercial["FLICHTINT"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("LICHTINT"))
                {
                    Draw_Fotocelulas();
                    log = String.Concat(log, "\nIncluidas fotocelulas de peines");
                }
            }
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);


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
            ProgressChanged(100);


            ////Luz bajopasamanos
            //c = (Caracteristic)oElectric.CaractComercial["FBALUBEL"];
            //c2 = (Caracteristic)oElectric.CaractComercial["FSHEITSBEL"];
            //c3 = (Caracteristic)oElectric.CaractComercial["ILUPEI"];
            //refVal = c.CurrentReference;
            //refVal2 = c2.CurrentReference;
            //refVal3 = c3.CurrentReference;
            //if (refVal != null && refVal2 != null && refVal3 != null)
            //{
            //    if (!refVal.Equals("BELOHNE"))
            //    {
            //        Draw_LuzBajopasamanos(refVal, refVal2, refVal3);
            //    }
            //}

            //

            ////Luz Zocalos
            //c = (Caracteristic)oElectric.CaractComercial["FSOCKELBEL"];
            //c2 = (Caracteristic)oElectric.CaractComercial["FBALUBEL"];
            //c3 = (Caracteristic)oElectric.CaractComercial["FSHEITSBEL"];
            //c4 = (Caracteristic)oElectric.CaractComercial["ILUPEI"];
            //refVal = c.CurrentReference;
            //refVal2 = c2.CurrentReference;
            //refVal3 = c3.CurrentReference;
            //refVal4 = c4.CurrentReference;
            //if (refVal != null && refVal2 != null && refVal3 != null && refVal4 != null)
            //{
            //    if (!refVal.Equals("BELOHNE"))
            //    {
            //        Draw_LuzZocalos(refVal, refVal2, refVal3, refVal4);
            //    }
            //}

            //


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
            //I9 C Standard input 9
            SetGECParameter(oProject, oElectric, "I9", (uint)GEC.Param.Fire_alarm_smoke_detector_1);

            //I10 C Standard input 10
            SetGECParameter(oProject, oElectric, "I10", (uint)GEC.Param.Bypass_VFD);

            //I13 C Standard input 13
            SetGECParameter(oProject, oElectric, "I13", (uint)GEC.Param.Earthquake);

            //I16 C Standard input 16
            SetGECParameter(oProject, oElectric, "I16", (uint)GEC.Param.FPGA_fault);

            //UI9 UDL1 Standard input 9
            SetGECParameter(oProject, oElectric, "UI9", (uint)GEC.Param.Top_skirt_right_SS);

            //UI10 UDL1 Standard input 10
            SetGECParameter(oProject, oElectric, "UI10", (uint)GEC.Param.Top_skirt_left_SS);

            //UI13	UDL1 Standard input 13
            SetGECParameter(oProject, oElectric, "UI13", (uint)GEC.Param.Top_vertical_comb_plate_right_SS);

            //UI14	UDL1 Standard input 14
            SetGECParameter(oProject, oElectric, "UI14", (uint)GEC.Param.Top_vertical_comb_plate_left_SS);

            //UI23    UDL1 Standard input 23 X23
            SetGECParameter(oProject, oElectric, "UI23", (uint)GEC.Param.Top_up_key_order);

            //UI24    UDL1 Standard input 24
            SetGECParameter(oProject, oElectric, "UI24", (uint)GEC.Param.Top_down_key_order);

            //LI11 LDL1 Standard input 11
            SetGECParameter(oProject, oElectric, "LI11", (uint)GEC.Param.Bottom_vertical_comb_plate_right_SS);

            //LI12 LDL1 Standard input 12
            SetGECParameter(oProject, oElectric, "LI12", (uint)GEC.Param.Bottom_vertical_comb_plate_left_SS);

            //LI13 LDL1 Standard input 13
            SetGECParameter(oProject, oElectric, "LI13", (uint)GEC.Param.Bottom_buggy_right_SS);

            //LI14 LDL1 Standard input 14
            SetGECParameter(oProject, oElectric, "LI14", (uint)GEC.Param.Bottom_buggy_left_SS);

            //LI15 LDL1 Standard input 15
            SetGECParameter(oProject, oElectric, "LI15", (uint)GEC.Param.Bottom_skirt_right_SS);

            //LI16 LDL1 Standard input 16 X16
            SetGECParameter(oProject, oElectric, "LI16", (uint)GEC.Param.Bottom_skirt_left_SS);

            //O5 C Relay 5 NO 5L/Q5
            SetGECParameter(oProject, oElectric, "O5", (uint)GEC.Param.Up_indication);

            //O6 C Relay 6 NO 6L/Q6
            SetGECParameter(oProject, oElectric, "O6", (uint)GEC.Param.Down_indication);

            //O7 C Relay 7 NO 7L/Q7
            SetGECParameter(oProject, oElectric, "O7", (uint)GEC.Param.Oil_pump_activation);

            //O8 C Relay 8 NO 8L/Q8
            SetGECParameter(oProject, oElectric, "O8", (uint)GEC.Param.Fault_indication);

            //O9 C Relay 9 NO 9L/Q9
            SetGECParameter(oProject, oElectric, "O9", (uint)GEC.Param.Maintenance_indication);

            //UO4 UDL1 Relay output 1 NO Q4/4L
            SetGECParameter(oProject, oElectric, "UO4", (uint)GEC.Param.Buzzer_top);

            //UO7	UDL1 Relay output 4 NO Q7/7L
            SetGECParameter(oProject, oElectric, "UO7", (uint)GEC.Param.Buzzer_during_start);

            //LO4	LDL1 Relay output 1 NO Q4/4L
            SetGECParameter(oProject, oElectric, "LO4", (uint)GEC.Param.Buzzer_bottom);

            //LO7	LDL1 Relay output 4 NO Q7/7L
            SetGECParameter(oProject, oElectric, "LO7", (uint)GEC.Param.Buzzer_during_start);

            //SI17	SF Safety Input 9
            SetGECParameter(oProject, oElectric, "SI17", (uint)GEC.Param.Drive_chain_DuTriplex, true);
            //C39	LOWER_DIAG_SS_LENGTH
            SetGECParameter(oProject, oElectric, "C39", 16);
            //Conexion Motor "VVF_YD"
            SetGECParameter(oProject, oElectric, "SI12", (uint)GEC.Param.Contactor_FB_2, true);
            SetGECParameter(oProject, oElectric, "SI21", (uint)GEC.Param.Contactor_FB_3, true);
            SetGECParameter(oProject, oElectric, "SI22", (uint)GEC.Param.Contactor_FB_4, true);
            SetGECParameter(oProject, oElectric, "O3", (uint)GEC.Param.Main_2, true);
            SetGECParameter(oProject, oElectric, "O4", (uint)GEC.Param.Main_1, true);
        }

        private void Draw_LuzEstroboscopica(string type)
        {

            switch (type)
            {
                //Una tira LED
                case "LED":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'L', "Upper Lighting I", 200.0, 156.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'K', "Lower Lighting I", 192.0, 100.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 1 Tira LED");
                    break;

                //Dos tiras LED
                case "2LED":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'N', "Upper Lighting I", 200.0, 156.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'M', "Lower Lighting I", 192.0, 100.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 2 Tira LED");
                    break;

                //Tres tiras LED
                case "3LED":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'P', "Upper Lighting I", 200.0, 156.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'O', "Lower Lighting I", 192.0, 100.0);

                    log = String.Concat(log, "\nIncluida luz estroboscópica: 3 Tira LED");
                    break;
            }
        }

        private void Draw_LuzPeines(string type, string luzEstro)
        {

            if (type.Contains("PREMIUM"))
            {
                //en página de "Upper Lighting I"
                if (luzEstro.Contains("LED"))
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'K', "Upper Lighting I", 252.0, 156.0);
                else
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'L', "Upper Lighting I", 200.0, 156.0);

                //en página de "Lower Lighting I"
                if (luzEstro.Contains("LED"))
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'N', "Lower Lighting I", 200.0, 144.0);
                else
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'M', "Lower Lighting I", 200.0, 144.0);

                log = String.Concat(log, "\nIncluida luz de peines LED 230V");
            }
        }

        private void Draw_LuzBajopasamanos(string type, string luzEstro, string luzPeines)
        {

            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'B', "Lower Lighting I", 156.0, 112.0);
                }
                else
                {
                    //Despues de la página de "Lower Lighting I"
                    insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                    //en página de "Lower Lighting II"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'A', "Lower Lighting II", 40.0, 188.0);

                    if (!luzEstro.Equals("KEINE") || !luzPeines.Equals("NO"))
                    {
                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'C', "Lower Lighting I", 176, 200.0);
                    }

                }


                log = String.Concat(log, "\nIncluida luz bajpasamanos LED 230V");
            }

        }

        private void Draw_LuzZocalos(string type, string luzBajopasamanos, string luzEstro, string luzPeines)
        {

            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && luzBajopasamanos.Equals("BELOHNE"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'A', "Lower Lighting I", 156.0, 112.0);
                }
                else if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && !luzBajopasamanos.Equals("BELOHNE"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'B', "Lower Lighting I",228.0, 172.0);
                }
                else
                {
                    if (!luzBajopasamanos.Equals("BELOHNE"))
                    {
                        //en página de "Lower Lighting II"
                        insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'C', "Lower Lighting II", 132.0, 188.0);
                    }
                    else
                    {
                        //Despues de la página de "Lower Lighting I"
                        //Lower Lighting
                        insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                        //en página de "Lower Lighting II"
                        insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'D', "Lower Lighting II", 40.0, 188.0);

                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'E', "Lower Lighting I", 176, 200.0);
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
                //en página de "Upper Lighting I"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 'G', "Upper Lighting I", 92.0, 156.0);

                //en página de "LLower Lighting I"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 'H', "Lower Lighting I", 40.0, 160.0);

                log = String.Concat(log, "\nIncluida luz de foso manual");
            }
        }

        private void Draw_Semaforo(string type)
        {

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            SetGECParameter(oProject, oElectric, "UO6", (uint)GEC.Param.Top_traffic_light_red, true);
            SetGECParameter(oProject, oElectric, "UO5", (uint)GEC.Param.Top_traffic_light_green, true);
            SetGECParameter(oProject, oElectric, "LO6", (uint)GEC.Param.Bottom_traffic_light_red, true);
            SetGECParameter(oProject, oElectric, "LO5", (uint)GEC.Param.Bottom_traffic_light_green, true);

            //Semaforos superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Traffic_Lights.emp", "Upper Keys", "Upper Traffic Lights");


            //Semaforos inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Traffic_Lights.emp", "Lower Keys", "Lower Traffic Lights");


            if (modelo.CurrentReference.Contains("CLASSIC"))
            {
                //en página de "Upper Traffic Lights"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'G', "Upper Traffic Lights", 128.0, 148.0);

                //en página de "Lower Traffic Lights"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'H', "Lower Traffic Lights", 128.0, 148.0);
                log = String.Concat(log, "\nIncluidos Semáforos Chinos de VC3.0");
            }
            else
            {

                if (type.Equals("BICOLOR"))
                {
                    //en página de "Upper Traffic Lights"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'A', "Upper Traffic Lights", 128.0, 148.0);

                    //en página de "Lower Traffic Lights"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'C', "Lower Traffic Lights", 128.0, 148.0);

                    log = String.Concat(log, "\nIncluidos Semáforos Bicolor");
                }
                else if (type.Equals("ROTGRUEN") ||
                        type.Equals("F6NEINB") ||
                        type.Equals("PROHI_VERDE"))
                {
                    //en página de "Upper Traffic Lights"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'B', "Upper Traffic Lights", 128.0, 148.0);

                    //en página de "Lower Traffic Lights"
                    insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'D', "Lower Traffic Lights", 128.0, 148.0);

                    log = String.Concat(log, "\nIncluidos Semáforos Flecha/Prohibido");
                }
                else
                {
                    log = String.Concat(log, "\nNo existe macro para el tipo de semaforo seleccionado");
                }

            }

            //en página de "Upper Diagnostic Outputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'P', "Upper Diagnostic Outputs I", 252.0, 116.0);

            //en página de "Lower Diagnostic Outputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light_ASME.ema", 'O', "Lower Diagnostic Outputs I", 252.0, 132.0);

        }

        private void Draw_Radar()
        {
            Caracteristic producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Radares Superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", "Upper Diagnostic Outputs I", "Upper People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'K', "Upper People Detection", 212.0, 256.0);
            else
                //en página de "Upper People Detection"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'E', "Upper People Detection", 212.0, 256.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'N', "Upper Diagnostic Inputs IV", 220.0, 168.0);

            SetGECParameter(oProject, oElectric, "UI25", (uint)GEC.Param.Top_radar_right_NC, true);
            SetGECParameter(oProject, oElectric, "UI26", (uint)GEC.Param.Top_radar_left_NC, true);

            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Radares Inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'L', "Lower People Detection", 212.0, 256.0);
            else
                //en página de "Lower People Detection"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'F', "Lower People Detection", 212.0, 256.0);

            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'M', "Lower Diagnostic Inputs IV", 216.0, 168.0);

            SetGECParameter(oProject, oElectric, "LI26", (uint)GEC.Param.Bottom_radar_right_NC, true);
            SetGECParameter(oProject, oElectric, "LI25", (uint)GEC.Param.Bottom_radar_left_NC, true);
        }

        private void Draw_Fotocelulas()
        {
            Caracteristic producto = (Caracteristic)oElectric.CaractComercial["FMODELL"];

            //Fotocélulas superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", "Upper Diagnostic Outputs I","Upper People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'L', "Upper People Detection", 12.0, 256.0);
            else
                //en página de "Upper People Detection"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'F', "Upper People Detection", 12.0, 256.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'N', "Upper Diagnostic Inputs IV", 220.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI25", (uint)GEC.Param.Top_light_barrier_comb_plate_NC, true);


            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Fotocélulas inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'K', "Lower People Detection", 12.0, 256.0);
            else
                //en página de "Lower People Detection"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'E', "Lower People Detection", 12.0, 256.0);


            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'M', "Lower Diagnostic Inputs IV", 216.0, 168.0);
            SetGECParameter(oProject, oElectric, "LI25", (uint)GEC.Param.Bottom_light_barrier_comb_plate_NC, true);



        }

        private void Draw_RoturaPasamanos()
        {
            //Compruebo si ya esta insertada la página "Lower Sensors II"
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_II.emp", "Lower Sensors I", "Lower Sensors II");

            //en página de "Lower Sensors II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'C', "Lower Sensors II", 36.0, 260.0);

            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'N', "Lower Diagnostic Inputs III", 160.0, 168.0);
            SetGECParameter(oProject, oElectric, "LI17", (uint)GEC.Param.Broken_handrail_R, true);
            SetGECParameter(oProject, oElectric, "LI18", (uint)GEC.Param.Broken_handrail_L, true);
        }

        private void Draw_ControlDesgasteFrenos()
        {

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_II.emp", "Upper Sensors I", "Upper Sensors II");

            //en página de "Upper Sensors II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'C', "Upper Sensors II", 28.0, 264.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'O', "Upper Diagnostic Inputs IV", 340.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI27", (uint)GEC.Param.Brake_wear_brake_1_M1, true);
            SetGECParameter(oProject, oElectric, "UI28", (uint)GEC.Param.Brake_wear_brake_2_M1, true);

        }

        private void Draw_BuggySuperior()
        {

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'E', "Upper Diagnostic Inputs III", 72.0, 156.0);

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'F', "Upper Diagnostic Inputs III", 12.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI15", (uint)GEC.Param.Top_buggy_right_SS, true);
            SetGECParameter(oProject, oElectric, "UI16", (uint)GEC.Param.Top_buggy_left_SS, true);
        }

        private void Draw_Cerrojo()
        {
            Caracteristic sMicrosZocalo = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)oElectric.CaractComercial["F04ZUB"];

            if (sBuggy.CurrentReference.Equals("KEINE") ||
                sBuggy.CurrentReference.Equals("BUGGYUT"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'B', "Upper Diagnostic Inputs III", 76.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI16", (uint)GEC.Param.Chain_locking_device_SS, true);

            }
            else if (sMicrosZocalo.NumVal <= 6)
            {

                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'B', "Upper Diagnostic Inputs II", 256.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI12", (uint)GEC.Param.Chain_locking_device_SS, true);

            }
        }

        private void Draw_StopAdicionalSuperior()
        {

            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'C', "Upper Diagnostic Inputs II", 192.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI11", (uint)GEC.Param.Top_emergency_stop_external_SS, true);
        }

        private void Draw_StopAdicionalInferior()
        {
            //No se puede usar de momento porque no esta habilitado

            //en página de "Lower Diagnostic Inputs III"
            //insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'D', "Lower Diagnostic Inputs III", 132.0, 156.0);
            //SetGECParameter(oProject, oElectric, "UI11", (uint)GEC.Param.Top_emergency_stop_external_SS, true);

        }

        private void Draw_Freno(string motorType)
        {

            if (motorType.Equals("QC") ||
                motorType.Equals("FJ"))
            {
                //Finales de carrera
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_I_ASME_FC.emp", "Control Outputs III", "Brake Sensors");
                log = String.Concat(log, "Finales de carrera");
            }
            else
            {
                //Inductivos
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_I_ASME_Sensor.emp", "Control Outputs III", "Brake Sensors");
                log = String.Concat(log, "Inductivos");
            }
        }

        private void Draw_Freno_Aux()
        {
            //Incluir sensor de encoder
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'J', "Upper Sensors I", 120, 264);

            int key;
            Insert oInsert = new Insert();

            //Incluir conexión en placa safety desplazando bornas para incluir relé
            Dictionary<int, string> dictPages = GetPageTable(oProject);
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Safety Pulse Inputs");
            Page page = oProject.Pages[key];
            page.Filter.FunctionCategory=Eplan.EplApi.Base.Enums.FunctionCategory.Terminal;
            Function[] functions = page.Functions;
            FunctionsFilter ff = new FunctionsFilter();
            ff.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.Terminal;
            ff.Page = page;
            ff.IsPlaced = true;
            DMObjectsFinder objFinder = new DMObjectsFinder(oProject);
            functions = objFinder.GetFunctions(ff);

            foreach (Function f in functions)
            {
                if (f.Name.Contains("X10:7A") ||
                    f.Name.Contains("X10:8A"))
                {
                    f.Location = new PointD(f.Location.X + 8, f.Location.Y);
                }
            }


            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'K', "Safety Pulse Inputs", 232, 140);
            SetGECParameter(oProject, oElectric, "SI7", (uint)GEC.Param.Main_shaft_speed_monitor_1, true);
            SetGECParameter(oProject, oElectric, "SI8", (uint)GEC.Param.Main_shaft_speed_monitor_2, true);

            //Incluir SAI
            moveSymbol(oProject, "Control I", -12, 12, 124, 212, 124, 212);

            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'L', "Control I", 16, 148);
            insertInterruptionPoint(oProject, @"BP", "SPECIAL", 'A', "Control I", "=ESC+MAIN-L_SAI", "Right, 0", 396, 200);
            insertInterruptionPoint(oProject, @"BP", "SPECIAL", 'A', "Control I", "=ESC+MAIN-N_SAI", "Right, 0", 396, 196);


            //Incluir rectificador
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'M', "Control II", 352, 220);
            insertInterruptionPoint(oProject, @"BP", "SPECIAL", 'A', "Control II", "=ESC+MAIN-L_SAI", "Left, 0", 24, 220);
            //insertSymbol(@"CF", "SPECIAL", 'D', "Control II", 336, 216);
            insertInterruptionPoint(oProject, @"BP", "SPECIAL", 'A', "Control II", "=ESC+MAIN-N_SAI", "Left, 0", 24, 216);

            //Incluir bobina y sensores de trinquete
            insertNewPage(oProject, "Auxiliary Brake", "Brake I");
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'N', "Auxiliary Brake", 56, 80);

            //Incluir feedback sensor trinquete
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'O', "Safety Inputs I", 324, 184);
            SetGECParameter(oProject, oElectric, "SI18", (uint)GEC.Param.Aux_brake_status_1, true);

            //Incluir alimentación rectificador
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'D', "Control Outputs I", 28, 80);


        }

        private void Draw_Trinquete_Mecanico()
        {
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI17", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);
        }

        private void Draw_Sincro_Pasamanos(string model)
        {
            if (!model.Contains("CLASSIC"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Handrail_Speed_Sensors.ema", 'A', "Upper Sensors I", 284, 260);
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Handrail_Speed_Sensors.ema", 'B', "Lower Sensors I", 124, 264);
            }
        }

        private void Draw_Freno_adicional(string motorType)
        {

            //en página de "Control I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'M', "Control I", 296.0, 152.0);

            if (motorType.Equals("QC") ||
                motorType.Equals("FJ"))
            {
                //Finales de carrera
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_II_ASME_FC.emp", "Brake I", "Brake II");
                log = String.Concat(log, "Finales de carrera");
            }
            else
            {
                //Inductivos
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_II_ASME_Sensor.emp", "Brake I", "Brake II");
                log = String.Concat(log, "Inductivos");
            }

            //en página de "Safety Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'L', "Safety Inputs I", 228.0, 184.0);
            SetGECParameter(oProject, oElectric, "SI15", (uint)GEC.Param.Brake_function_brake_3_mot_1, true);
            SetGECParameter(oProject, oElectric, "SI16", (uint)GEC.Param.Brake_function_brake_4_mot_1, true);

        }

        private void Draw_Lubricacion_auto()
        {

            //en página de "Sensors & Oil Pump"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'E', "Sensors & Oil Pump", 272.0, 144.0);

            //en página de "Control I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'F', "Control I", 216.0, 52.0);

            //en página de "Control Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'G', "Control Inputs I", 192.0, 204.0);
            SetGECParameter(oProject, oElectric, "I5", (uint)GEC.Param.Oil_level_in_pump_1, true);
        }

        private void Draw_VVF(Double power)
        {
            //en página de "VVF Power"
            StorableObject[] oInsertedObjects = insertWindowMacro_ObjCont(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'J', "VVF Power", 44.0, 208.0);

            foreach (StorableObject oSOTemp in oInsertedObjects)
            {
                //we are searching for PlaceHolder 'Three-Phase' in the results
                PlaceHolder oPlaceHoldeThreePhase = oSOTemp as PlaceHolder;
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

            deleteArea(oProject, "VVF Power", 92, 208, 116, 208);

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\VVF_Control_ASME.emp", "VVF Power", "VVF Control");

            //en página de "Control Outputs II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'L', "Control Outputs II", 188.0, 132.0);
            SetGECParameter(oProject, oElectric, "O2", (uint)GEC.Param.Speed_selection_output_2, true);

            //en página de "Control Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'K', "Control Inputs I", 160.0, 204.0);
            SetGECParameter(oProject, oElectric, "I4", (uint)GEC.Param.VFD_EEC, true);

        }

        private void Draw_VVF_Regen(Double power)
        {

            //en página de "VVF Power"
            StorableObject[] oInsertedObjects = insertWindowMacro_ObjCont(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'N', "VVF Power", 76.0, 184.0);

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
            deleteArea(oProject, "VVF Power", 92, 208, 116, 208);

            //Despues de la página de "VVF"
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\VVF_Control_ASME_Regenerative.emp", "VVF Power", "VVF Control");


            //en página de "Control Outputs II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'L', "Control Outputs II", 188.0, 132.0);
            SetGECParameter(oProject, oElectric, "O2", (uint)GEC.Param.Speed_selection_output_2, true);

            //en página de "Control Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'K', "Control Inputs I", 160.0, 204.0);
            SetGECParameter(oProject, oElectric, "I4", (uint)GEC.Param.VFD_EEC, true);




        }
        #endregion

    }

}
