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
    public class DrawArmStandardIntASME : DrawTools
    {
        private Project oProject;
        private Electric oElectric;
        private string log;


        public DrawArmStandardIntASME(Project project, Electric electric)
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
            c = (Caracteristic)oElectric.CaractComercial["FANTREHT"];
            Draw_Freno(c.CurrentReference);

            //oProgressBar.Maximum = oElectric.Caracteristics.Count;

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
            

            //Freno Auxiliar
            c = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Contains("CLASSIC"))
                {
                    Draw_Freno_Aux();
                    log = String.Concat(log, "\nFreno Auxiliar");
                }

            }
            

            //Sensores de sincronismo pasamanos
            c = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                Draw_Sincro_Pasamanos(refVal);
                log = String.Concat(log, "\nIncluido sincronismo de pasamanos");
            }
            

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

            MessageBox.Show(log);
            String path = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "Log\\Log_Draw.txt");
            File.WriteAllText(path, log);



        }

       
        #region Metodos de dibujo
        public void Draw_LuzEstroboscopica(string type)
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

        public void Draw_LuzPeines(string type, string luzEstro)
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

        public void Draw_LuzBajopasamanos(string type, string luzEstro, string luzPeines)
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

        public void Draw_LuzZocalos(string type, string luzBajopasamanos, string luzEstro, string luzPeines)
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

        public void Draw_LuzFoso(string type)
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

        public void Draw_Semaforo(string type)
        {

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];

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

        public void Draw_Radar()
        {

            //Radares Superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", "Upper Diagnostic Outputs I", "Upper People Detection");

            //en página de "Upper People Detection"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'E', "Upper People Detection", 212.0, 256.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'N', "Upper Diagnostic Inputs IV", 220.0, 168.0);

            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Radares Inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");
            
            //en página de "Lower People Detection"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'F', "Lower People Detection", 212.0, 256.0);

            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'M', "Lower Diagnostic Inputs IV", 216.0, 168.0);

        }

        public void Draw_Fotocelulas()
        {

            //Fotocélulas superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_People_Detection.emp", "Upper Diagnostic Outputs I","Upper People Detection");
           
            //en página de "Upper People Detection"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'F', "Upper People Detection", 12.0, 256.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'N', "Upper Diagnostic Inputs IV", 220.0, 168.0);


            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Fotocélulas inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");

            //en página de "Lower People Detection"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'E', "Lower People Detection", 12.0, 256.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'M', "Lower Diagnostic Inputs IV", 216.0, 168.0);



        }

        public void Draw_RoturaCadenaPrincipal()
        {

            //en página de "Motor Sensors"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 'A', "Motor Sensors", 296.0, 260.0);


            //en página de "Safety Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 'P', "Safety Inputs I", 292.0, 184.0);
        }

        public void Draw_RoturaPasamanos()
        {
            //Compruebo si ya esta insertada la página "Lower Sensors II"
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Sensors_II.emp", "Lower Sensors I", "Lower Sensors II");

            //en página de "Lower Sensors II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'C', "Lower Sensors II", 36.0, 260.0);

            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'N', "Lower Diagnostic Inputs III", 160.0, 168.0);
        }

        public void Draw_ControlDesgasteFrenos()
        {

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Sensors_II.emp", "Upper Sensors I", "Upper Sensors II");

            //en página de "Upper Sensors II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'C', "Upper Sensors II", 28.0, 264.0);

            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'O', "Upper Diagnostic Inputs IV", 340.0, 168.0);
        }

        public void Draw_BuggyInferior()
        {

            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'B', "Lower Diagnostic Inputs III", 16.0, 156.0);
        }

        public void Draw_BuggySuperior()
        {

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'E', "Upper Diagnostic Inputs III", 72.0, 156.0);

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'F', "Upper Diagnostic Inputs III", 12.0, 156.0);
        }

        public void Draw_Cerrojo()
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

        public void Draw_SeguridadVerticalPeines()
        {

            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 'A', "Upper Diagnostic Inputs II", 252.0, 156.0);

            //en página de "Lower Diagnostic Inputs II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 'B', "Lower Diagnostic Inputs II", 312.0, 156.0);
        }

        public void Draw_StopAdicionalSuperior()
        {

            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'C', "Upper Diagnostic Inputs II", 192.0, 156.0);
        }

        public void Draw_StopAdicionalInferior()
        {
            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Stop_External.ema", 'D', "Lower Diagnostic Inputs III", 132.0, 156.0);
        }

        public void Draw_Micros_Zocalo(int nMicros)
        {
            int key;
            Insert oInsert = new Insert();

            if (nMicros >= 4)
            {
                //en página de "Upper Diagnostic Inputs II"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Microswitch_1.ema", 'A', "Upper Diagnostic Inputs II", 72.0, 156.0);

                //en página de "Lower Diagnostic Inputs II"
                insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Microswitch_1.ema", 'B', "Lower Diagnostic Inputs II", 192.0, 156.0);
            }

        }

        public void Draw_Freno(string motorType)
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

        public void Draw_Freno_Aux()
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

        public void Draw_Sincro_Pasamanos(string model)
        {
            if (!model.Equals("VELINO_CLASSIC"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Handrail_Speed_Sensors.ema", 'A', "Upper Sensors I", 284, 260);
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Handrail_Speed_Sensors.ema", 'B', "Lower Sensors I", 124, 264);
            }
        }

        public void Draw_Freno_adicional(string motorType)
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

        }

        public void Draw_Lubricacion_auto()
        {

            ////Oil Level In Pump 1
            //GEC C_Standard_input_5 = new GEC("I5", 64);
            //GEC_Parameter.Add(C_Standard_input_5);

            //en página de "Sensors & Oil Pump"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'E', "Sensors & Oil Pump", 272.0, 144.0);

            //en página de "Control I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'F', "Control I", 216.0, 52.0);

            //en página de "Control Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'G', "Control Inputs I", 192.0, 204.0);
        }

        public void Draw_VVF(Double power)
        {

            ////Entrada Fallo variador
            //GEC C_Standard_input_4 = new GEC("I4", 69);
            //GEC_Parameter.Add(C_Standard_input_4);

            ////Salida segunda velocidad
            //GEC C_Relay_2_NO = new GEC("O2", 201);
            //GEC_Parameter.Add(C_Relay_2_NO);

            ////Entrada Bypass
            //GEC C_Standard_input_10 = new GEC("I10", 65);
            //GEC_Parameter.Add(C_Standard_input_10);

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

            //en página de "Control Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'K', "Control Inputs I", 160.0, 204.0);

        }

        public void Draw_VVF_Regen(Double power)
        {

            ////Entrada Fallo variador
            //GEC C_Standard_input_4 = new GEC("I4", 69);
            //GEC_Parameter.Add(C_Standard_input_4);

            ////Salida segunda velocidad
            //GEC C_Relay_2_NO = new GEC("O2", 201);
            //GEC_Parameter.Add(C_Relay_2_NO);

            ////Entrada Bypass
            //GEC C_Standard_input_10 = new GEC("I10", 65);
            //GEC_Parameter.Add(C_Standard_input_10);

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

            //en página de "Control Inputs I"
            insertWindowMacro(oProject,"$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'K', "Control Inputs I", 160.0, 204.0);




        }
        #endregion

    }

}
