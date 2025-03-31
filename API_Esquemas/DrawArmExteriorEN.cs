using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel.Graphics;
using Eplan.EplApi.HEServices;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.E3D;
using EPLAN_API.SAP;
using System.Threading;
using Eplan.EplApi.Base.Enums;

namespace EPLAN_API.User
{
    class DrawArmExteriorEN : DrawTools
    {
        private Project oProject;
        private Page[] oPages;
        private Hashtable oHPages;
        //Dictionary<int, string> dictPages;
        private Electric oElectric;
        private string log;
        public delegate void ProgressChangedDelegate(int value);
        public event ProgressChangedDelegate ProgressChangedToDraw;

        public DrawArmExteriorEN(Project project, Electric electric)
        {
            oProject = project;
            oElectric = electric;
        }


        public void DrawMacro()
        {
            Caracteristic c, c2, c3, c4;
            String refVal, refVal2, refVal3, refVal4;

            int progress = 0;
            int step = 100 / 44;

            Draw_Default_Param();
            progress += step;
            ProgressChanged(progress);

            Draw_Implantacion();
            Draw_Aux_Implantacion();
            progress += step;
            ProgressChanged(progress);

            //sensores de freno
            c = (Caracteristic)oElectric.CaractComercial["FANTREHT"];
            Draw_Freno(c.CurrentReference);
            progress += step;
            ProgressChanged(progress);

            Draw_Motor();
            progress += step;
            ProgressChanged(progress);

            Draw_Display();
            progress += step;
            ProgressChanged(progress);

            CalculateCableSec();
            progress += step;
            ProgressChanged(progress);

            Draw_VVF();
            progress += step;
            ProgressChanged(progress);

            Draw_Contactors();
            progress += step;
            ProgressChanged(progress);

            Draw_Temperatura_Motor();
            progress += step;
            ProgressChanged(progress);

            Draw_Sincronismo();
            progress += step;
            ProgressChanged(progress);

            //PLC
            c = (Caracteristic)oElectric.CaractIng["TNCR_DO_CONTROL"];
            if (c.CurrentReference.Equals("GEC+PLC"))
            {
                Draw_PLC();
                log = String.Concat(log, "\nIncluido PLC");
            }
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
                    log = String.Concat(log, "\r\nIncluido segundo freno");
                }
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
                    log = String.Concat(log, "\r\nIncluida Bomba de lubricacion");
                }
            }
            progress += step;
            ProgressChanged(progress);

            #region linea de seguridad

            //Micros de Zócalo
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            if (c.NumVal >= 4)
            {
                Draw_Micros_Zocalo(Convert.ToInt16(c.NumVal));
                log = String.Concat(log, "\r\nIncluidos ", Convert.ToInt16(c.NumVal).ToString(), " micros de zócalo");
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
                    log = String.Concat(log, "\r\nIncluido Stop adicional superior");
                }
            }
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);

            //Stop adicional inferior
            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("2"))
                {
                    Draw_StopAdicionalInferior();
                    log = String.Concat(log, "\r\nIncluido Stop adicional inferior");
                }
            }
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);

            //Seguridad vertical de peines
            c = (Caracteristic)oElectric.CaractComercial["FKAMMPLHK"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("INDEPENDIENTE"))
                {
                    Draw_SeguridadVerticalPeines();
                    log = String.Concat(log, "\r\nIncluida seguridad vertical de peines");
                }
            }
            progress += step;
            ProgressChanged(progress);

            //Seguridad de buggy inferior
            c = (Caracteristic)oElectric.CaractComercial["F04ZUB"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("BUGGY") ||
                refVal.Equals("BUGGYUT"))
                {
                    Draw_BuggyInferior();
                    log = String.Concat(log, "\r\nIncluida seguridad buggy inferior");
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
                    log = String.Concat(log, "\r\nIncluida seguridad buggy superior");
                }
            }
            progress += step;
            ProgressChanged(progress);

            #endregion

            #region otras seguridades
            //Control de desgaste de frenos
            c = (Caracteristic)oElectric.CaractComercial["F01ZUB"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("INDUCTIVO"))
                {
                    Draw_ControlDesgasteFrenos();
                    log = String.Concat(log, "\r\nIncluido control de desgaste de frenos");
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
                    log = String.Concat(log, "\r\nIncluida seguridad de rotura de pasamanos");
                }
            }
            progress += step;
            ProgressChanged(progress);

            //Seguridad de cadena principal
            c = (Caracteristic)oElectric.CaractComercial["TNCR_S_DRIVE_CHAIN"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("SI"))
                {
                    Draw_RoturaCadenaPrincipal();
                    log = String.Concat(log, "\r\nIncluida seguridad de cadena principal");
                }
            }
            progress += step;
            ProgressChanged(progress);

            //Trinquete Magnetico
            c = (Caracteristic)oElectric.CaractComercial["FZUSBREMSE"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("HWSPERRKMAGN") || refVal.Equals("NAB"))
                {
                    Draw_Trinquete(refVal);
                    log = String.Concat(log, "\nIncluido Trinquete Magnetico");
                }
            }
            progress += step;
            ProgressChanged(progress);

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
            #endregion

            #region Deteccion de personas
            //Detección de personas por radar
            c = (Caracteristic)oElectric.CaractComercial["FLICHTINT"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("RADAR"))
                {
                    Draw_Radar();
                    log = String.Concat(log, "\r\nIncluidos Radares");
                }
            }
            progress += step;
            ProgressChanged(progress);

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
                    log = String.Concat(log, "\r\nIncluidas fotocelulas de peines");
                }
            }
            progress += step;
            ProgressChanged(progress);

            c = (Caracteristic)oElectric.CaractComercial["TNCR_OT_NIVEL_AGUA"];
            refVal = c.CurrentReference;
            if (refVal != null)
            {
                if (refVal.Equals("S"))
                {
                    Draw_Nivel_Agua();
                }
            }
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);

            #endregion

            #region Llavines
            //Llavin Local/Remoto
            c = (Caracteristic)oElectric.CaractComercial["LLAVES_LOCAL_REM"];
            c2 = (Caracteristic)oElectric.CaractIng["TNCR_DO_CONTROL"];
            refVal = c.CurrentReference;
            refVal2 = c2.CurrentReference;
            if (refVal != null && refVal2 != null)
            {
                if (!refVal.Equals("N"))
                {
                    Draw_Llavin_Local_Remoto(refVal, refVal2);
                    log = String.Concat(log, "\nIncluido Llavin de Local/Remoto");
                }
            }
            progress += step;
            ProgressChanged(progress);

            //Llavin Automatico/Continuo
            c = (Caracteristic)oElectric.CaractComercial["LLAVES_AUT_CONT"];
            c2 = (Caracteristic)oElectric.CaractIng["TNCR_DO_CONTROL"];
            refVal = c.CurrentReference;
            refVal2 = c2.CurrentReference;
            if (refVal != null && refVal2 != null)
            {
                if (!refVal.Equals("N"))
                {
                    Draw_Llavin_Auto_Cont(refVal, refVal2);
                    log = String.Concat(log, "\nIncluido Llavin de Automatico/Continuo");
                }
            }
            progress += step;
            ProgressChanged(progress);

            //Llavín paro por impulso
            c = (Caracteristic)oElectric.CaractComercial["LLAVES_PARO"];
            c2 = (Caracteristic)oElectric.CaractIng["TNCR_DO_CONTROL"];
            refVal = c.CurrentReference;
            refVal2 = c2.CurrentReference;
            if (refVal != null && refVal2 != null)
            {
                if (!refVal.Equals("N"))
                {
                    Draw_Llavin_Paro(refVal);
                    log = String.Concat(log, "\nIncluido Llavín paro por impulso");
                }
            }
            progress += step;
            ProgressChanged(progress);
            #endregion

            #region Iluminacion
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
            progress += step;
            ProgressChanged(progress);

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
            ProgressChanged(progress);

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
            ProgressChanged(progress);

            #endregion

            #region Armario
            Draw_Envolvente_Armario();
            progress += step;
            ProgressChanged(progress);

            Draw_Protecciones();
            progress += step;
            ProgressChanged(progress);

            paramGEC(oProject, oElectric);
            ProgressChanged(100);

            #endregion

            Reports report = new Reports();
            report.GenerateProject(oProject);

            //Redraw
            Edit edit = new Edit();
            edit.RedrawGed();

            MessageBox.Show(new Form() { TopMost = true, TopLevel = true }, log, "Resultado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ProgressChanged(0);
            String path = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "Log\\Log_Draw.txt");
            File.WriteAllText(path, log);


        }

        private void ProgressChanged(int progress)
        {
            ProgressChangedToDraw(progress);
        }


        private void SetRecordContactor(StorableObject[] oInsertedObjects, Caracteristic iMotor)
        {
            foreach (StorableObject oSOTemp in oInsertedObjects)
            {
                PlaceHolder oPlaceHoldeThreePhase = oSOTemp as PlaceHolder;
                if (oPlaceHoldeThreePhase != null)
                {
                    if (iMotor.NumVal <= 18.0)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("I<=18");
                    }
                    else if (iMotor.NumVal > 18.0 && iMotor.NumVal <= 25.0)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("18<I<=25");
                    }
                    else if (iMotor.NumVal > 25.0 && iMotor.NumVal <= 32)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("25<I<=32");
                    }
                    else if (iMotor.NumVal > 32.0 && iMotor.NumVal <= 40)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("32<I<=40");
                    }
                    else if (iMotor.NumVal > 40.0 && iMotor.NumVal <= 50)
                    {
                        oPlaceHoldeThreePhase.ApplyRecord("40<I<=50");
                    }
                }
            }
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

        }

        private void Draw_Implantacion()
        {
            Caracteristic envolvente = (Caracteristic)oElectric.CaractIng["ENVOLV_ARM_EXT"];

            InstallationSpace installationSpace = new InstallationSpace();
            foreach (InstallationSpace iSpace in oProject.InstallationSpaces)
            {
                if (iSpace.ToString().Equals("MAIN"))
                    installationSpace = iSpace;
            }

            //2D
            if (envolvente.CurrentReference.Contains("1800x800x400"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Implantacion_Arm_Ext.ema", 'A', "Layout", 615, 2180);

            if (envolvente.CurrentReference.Contains("1800x1000x400"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Implantacion_Arm_Ext.ema", 'B', "Layout", 615, 2180);

            if (envolvente.CurrentReference.Contains("1800x1200x400"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Implantacion_Arm_Ext.ema", 'C', "Layout", 615, 2180);

            //3D
            if (envolvente.CurrentReference.Contains("1800x800x400")) 
                insert3DMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Implantacion_Arm_Ext.ema", 'A', installationSpace, 0, 0, 0);

            if (envolvente.CurrentReference.Contains("1800x1000x400"))
                insert3DMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Implantacion_Arm_Ext.ema", 'B', installationSpace, 0, 0, 0);

            if (envolvente.CurrentReference.Contains("1800x1200x400"))
                insert3DMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Implantacion_Arm_Ext.ema", 'C', installationSpace, 0, 0, 0);


        }

        private void Draw_Aux_Implantacion()
        {
            //Mover placas GEC hacia arriba
            Functions3DFilter oFunctions3DFilter = new Functions3DFilter();
            Function3DPropertyList oFunction3DPropertyList = new Function3DPropertyList();
            oFunction3DPropertyList.FUNC_VISIBLENAME = "-A28";
            oFunctions3DFilter.SetFilteredPropertyList(oFunction3DPropertyList);
            Function3D[] oFunctions3D = new DMObjectsFinder(oProject).GetFunctions3D(oFunctions3DFilter);
            Component A28 = oFunctions3D[0] as Component;
            A28.MoveRelative(0, 0, 20);

            oFunction3DPropertyList.FUNC_VISIBLENAME = "-A29";
            oFunctions3DFilter.SetFilteredPropertyList(oFunction3DPropertyList);
            oFunctions3D = new DMObjectsFinder(oProject).GetFunctions3D(oFunctions3DFilter);
            Component A29 = oFunctions3D[0] as Component;
            A29.MoveRelative(0, 0, 20);

        }

        private void Draw_Display()
        {
            Caracteristic tipoDisplay = (Caracteristic)oElectric.CaractIng["TNCR_DO_DISPLAY_TYPE"];
            Caracteristic ubicacionDisplay = (Caracteristic)oElectric.CaractComercial["TNDIAGNOSTICO"];
            Caracteristic armario = (Caracteristic)oElectric.CaractIng["ENVOLV_ARM_EXT"];

            if (!ubicacionDisplay.CurrentReference.Equals("NO"))
            {
                if (tipoDisplay.CurrentReference.Equals("DDU"))
                {
                    if (ubicacionDisplay.CurrentReference.Equals("AR") ||
                        ubicacionDisplay.CurrentReference.Equals("AR_BA") ||
                        ubicacionDisplay.CurrentReference.Equals("AR_ZO"))
                    {
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'G', "Display", 20.0, 252.0);

                        if (armario.CurrentReference.Contains("1800x800"))
                            insertDeviceLayout(oProject, "U1", "Display", "M1", 0, 'A', "A", "Layout", 1220, 1410);

                        if (armario.CurrentReference.Contains("1800x1000") ||
                            armario.CurrentReference.Contains("1800x1200"))
                            insertDeviceLayout(oProject, "U1", "Display", "M1", 0, 'A', "A", "Layout", 1204, 1243);

                        Insert3DDeviceIntoMountingPlate(oProject, "-U1", 370, 1010, planeoffset: 60);
                    }
                    else
                    {
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'F', "Display", 20.0, 252.0);
                    }
                }
                else if (tipoDisplay.CurrentReference.Equals("ESCATRONIC"))
                {
                    if (ubicacionDisplay.CurrentReference.Equals("AR") ||
                        ubicacionDisplay.CurrentReference.Equals("AR_BA") ||
                        ubicacionDisplay.CurrentReference.Equals("AR_ZO"))
                    {
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'E', "Display", 20.0, 252.0);

                        if (armario.CurrentReference.Contains("1800x800"))
                        {
                            insertDeviceLayout(oProject,"U1", "Display", "M1", 0, 'A', "A", "Layout", 1220, 1410);
                            insertDeviceLayout(oProject,"U2", "Display", "M1", 0, 'A', "A", "Layout", 1316, 986);
                        }

                        if (armario.CurrentReference.Contains("1800x1000") ||
                            armario.CurrentReference.Contains("1800x1200"))
                        {
                            insertDeviceLayout(oProject,"U1", "Display", "M1", 0, 'A', "A", "Layout", 1204, 1243);
                            insertDeviceLayout(oProject,"U2", "Display", "M1", 0, 'A', "A", "Layout", 1503, 1240);
                        }
                    }
                    else
                    {
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Display.ema", 'D', "Display", 20.0, 252.0);
                    }
                }
            }



        }

        private void CalculateCableSec()
        {
            Caracteristic cSeccMotor = (Caracteristic)oElectric.CaractIng["SECCABLEMOT"];

            switch (cSeccMotor.CurrentReference)
            {
                case "2,5":
                case "4":
                case "4A":
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x4_BN_SH", 40);
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x6_Y-BN_SH", 10);
                    break;

                case "6":
                case "6A":
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x6_BN_SH", 40);
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x6_Y-BN_SH", 10);
                    break;

                case "10":
                case "10A":
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x10_BN_SH", 40);
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x10_Y-BN_SH", 10);
                    break;

                case "16":
                case "16A":
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x16_BN_SH", 40);
                    insertArticle(oProject,"Main Supply", "WIRE", "W.1x16_Y-BN_SH", 10);
                    break;
            }

        }

        private void Draw_Motor()
        {
            Caracteristic motorConnection = (Caracteristic)oElectric.CaractIng["CONEXMOTOR"];
            Caracteristic iTermico = (Caracteristic)oElectric.CaractIng["ITERMICO"];
            Caracteristic seccionCableMotor = (Caracteristic)oElectric.CaractIng["SECCABLEMOT"];

            int key;
            Insert oInsert = new Insert();

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            //en página de "Motor"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Motor");

            //Seleccionar variante en funcion de conexion de motor
            char variante = 'F';
            switch (motorConnection.CurrentReference)
            {
                case "VVF_YD":
                    variante = 'F';
                    SetGECParameter(oProject, oElectric, "SI12", (uint)GEC.Param.Contactor_FB_2, true);
                    SetGECParameter(oProject, oElectric, "SI21", (uint)GEC.Param.Contactor_FB_3, true);
                    SetGECParameter(oProject, oElectric, "SI22", (uint)GEC.Param.Contactor_FB_4, true);
                    SetGECParameter(oProject, oElectric, "O3", (uint)GEC.Param.Main_2, true);
                    SetGECParameter(oProject, oElectric, "O4", (uint)GEC.Param.Main_1, true);
                    log = String.Concat(log, "\r\nIncluido Motor con variador y bypass YD");
                    break;
                case "VVF_D":
                    variante = 'G';
                    SetGECParameter(oProject, oElectric, "SI21", (uint)GEC.Param.Contactor_FB_3, true);
                    SetGECParameter(oProject, oElectric, "SI22", (uint)GEC.Param.Contactor_FB_4, true);
                    SetGECParameter(oProject, oElectric, "O3", (uint)GEC.Param.Main_2, true);
                    SetGECParameter(oProject, oElectric, "O4", (uint)GEC.Param.Main_1, true);
                    log = String.Concat(log, "\r\nIncluido Motor con variador y bypass D");
                    break;
                case "YD":
                    variante = 'H';
                    SetGECParameter(oProject, oElectric, "SI12", (uint)GEC.Param.Contactor_FB_2, true);
                    log = String.Concat(log, "\r\nIncluido Motor sin variador con arranque YD");
                    break;
                case "D":
                    variante = 'I';
                    log = String.Concat(log, "\r\nIncluido Motor sin variador con arranque D");
                    break;
                case "VVF":
                    variante = 'J';
                    SetGECParameter(oProject, oElectric, "SI21", (uint)GEC.Param.Contactor_FB_3, true);
                    SetGECParameter(oProject, oElectric, "O3", (uint)GEC.Param.Main_2, true);
                    log = String.Concat(log, "\r\nIncluido Motor con variador sin bypass");
                    break;
            }

            //Insertar macro con objetos configurables
            StorableObject[] oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Motor.ema", variante, "Motor", 12.0, 140.0);

            //Configurar objetos
            foreach (StorableObject oSOTemp in oInsertedObjects)
            {
                PlaceHolder oPlaceHoldeThreePhase = oSOTemp as PlaceHolder;
                if (oPlaceHoldeThreePhase != null)
                {
                    switch (oPlaceHoldeThreePhase.Name)
                    {
                        case "Termico":
                            if (iTermico.NumVal <= 13)
                            {
                                oPlaceHoldeThreePhase.ApplyRecord("0<I<13");
                            }
                            else if (iTermico.NumVal > 13 && iTermico.NumVal <= 18.0)
                            {
                                oPlaceHoldeThreePhase.ApplyRecord("13<=I<18");
                            }
                            else if (iTermico.NumVal > 18.0 && iTermico.NumVal <= 25)
                            {
                                oPlaceHoldeThreePhase.ApplyRecord("18<=I<25");
                            }
                            else if (iTermico.NumVal > 25.0 && iTermico.NumVal <= 32)
                            {
                                oPlaceHoldeThreePhase.ApplyRecord("25<=I<32");
                            }
                            else if (iTermico.NumVal > 32.0 && iTermico.NumVal <= 40)
                            {
                                oPlaceHoldeThreePhase.ApplyRecord("32<=I<40");
                            }
                            else if (iTermico.NumVal > 40.0 && iTermico.NumVal <= 50)
                            {
                                oPlaceHoldeThreePhase.ApplyRecord("40<=I<50");
                            }
                            break;

                        case "Cable Motor":
                            oPlaceHoldeThreePhase.ApplyRecord(seccionCableMotor.CurrentReference);
                            break;
                    }
                }
            }

            //Insertar implantación del termico
            insertDeviceLayout(oProject, "FR1", "Motor", "M1", 0, 'A', "A", "Layout", 1200, 765);
            Insert3DDeviceIntoDINRail(oProject, "U17", "-FR1",0, firstInRail:true);

        }

        private void Draw_Contactors()
        {
            int widthOffset = 0;

            Caracteristic motorConnection = (Caracteristic)oElectric.CaractIng["CONEXMOTOR"];
            Caracteristic iTermico = (Caracteristic)oElectric.CaractIng["ITERMICO"];
            Caracteristic iMotor = (Caracteristic)oElectric.CaractIng["IMOTOR"];

            List<StorableObject> oInsertedObjectsList = new List<StorableObject>();
            StorableObject[] oInsertedObjects;

            //K1.1 y K1.2
            if (!motorConnection.CurrentReference.Equals("VVF"))
            {
                //Coil
                //en página de "Control Outputs I"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'D', "Control Outputs I", 284.0, 128.0);
                SetRecordContactor(oInsertedObjects, iMotor);
                //Feedback
                //en página de "Safety Inputs I"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'H', "Safety Inputs I", 100.0, 268.0);
                SetRecordContactor(oInsertedObjects, iMotor);
                //Layout
                Insert3DDeviceIntoDINRail(oProject, "U17", "-K1.1", 0, rightTo: "-FR1");
                Insert3DDeviceIntoDINRail(oProject, "U17", "-K1.2", 0, rightTo: "-K1.1");
                Function placement = insertDeviceLayout(oProject, "K1.1", "Control Outputs I", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
                placement = insertDeviceLayout(oProject, "K1.2", "Control Outputs I", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
            }
            else
            {
                //Coil
                //en página de "Control Outputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Reles.ema", 'A', "Control Outputs I", 284.0, 128.0);
                //Feedback
                //en página de "Safety Inputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Reles.ema", 'B', "Safety Inputs I", 100.0, 268.0);
                //Layout
                Insert3DDeviceIntoDINRail(oProject, "U17", "-KA1.1", 0, rightTo: "-FR1");
                Insert3DDeviceIntoDINRail(oProject, "U17", "-KA1.2", 0, rightTo: "-KA1.1");
                Function placement = insertDeviceLayout(oProject,"KA1.1", "Control Outputs I", "M1", 1, 'A', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
                placement = insertDeviceLayout(oProject,"KA1.2", "Control Outputs I", "M1", 1, 'A', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
            }

            //K2.1 y K2.2
            if (motorConnection.CurrentReference.Equals("YD") ||
                motorConnection.CurrentReference.Equals("VVF_YD"))
            {
                //Coil
                //en página de "Control Outputs I"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'E', "Control Outputs I", 340.0, 128.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Feedback
                //en página de "Safety Inputs I"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'I', "Safety Inputs I", 132.0, 268.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Layout
                Insert3DDeviceIntoDINRail(oProject, "U17", "-K2.1", 0, rightTo: "-K1.2");
                Insert3DDeviceIntoDINRail(oProject, "U17", "-K2.2", 0, rightTo: "-K2.1");
                Function placement = insertDeviceLayout(oProject,"K2.1", "Control Outputs I", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
                placement = insertDeviceLayout(oProject,"K2.2", "Control Outputs I", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
            }

            //K10.1
            if (motorConnection.CurrentReference.Equals("VVF_D") ||
                motorConnection.CurrentReference.Equals("VVF_YD") ||
                motorConnection.CurrentReference.Equals("VVF"))
            {
                //Coil
                //en página de "Control Outputs II"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'K', "Control Outputs II", 284.0, 76.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Feedback
                //en página de "Safety Inputs II"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'F', "Safety Inputs II", 64.0, 268.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Layout
                if (motorConnection.CurrentReference.Equals("YD") ||
                    motorConnection.CurrentReference.Equals("VVF_YD"))
                    Insert3DDeviceIntoDINRail(oProject, "U17", "-K10.1", 0, rightTo: "-K2.2");
                else if (motorConnection.CurrentReference.Equals("VVF_YD"))
                    Insert3DDeviceIntoDINRail(oProject, "U17", "-K10.1", 0, rightTo: "-KA1.2");
                else
                    Insert3DDeviceIntoDINRail(oProject, "U17", "-K10.1", 0, rightTo: "-K1.2");
                Function placement = insertDeviceLayout(oProject,"K10.1", "Control Outputs II", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
            }

            //K10.2
            if (motorConnection.CurrentReference.Equals("VVF_D") ||
                motorConnection.CurrentReference.Equals("VVF_YD"))
            {
                //Coil
                //en página de "Control Outputs II"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'B', "Control Outputs II", 284.0, 76.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Feedback
                //en página de "Safety Inputs II"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'G', "Safety Inputs II", 96.0, 268.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Layout
                Insert3DDeviceIntoDINRail(oProject, "U17", "-K10.2", 0, rightTo: "-K10.1");
                Function placement = insertDeviceLayout(oProject,"K10.2", "Control Outputs II", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
            }

            //K10.3
            if (motorConnection.CurrentReference.Equals("VVF"))
            {
                //Coil
                //en página de "Control Outputs II"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'C', "Control Outputs II", 284.0, 76.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Feedback
                //en página de "Safety Inputs II"
                oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Contactores.ema", 'J', "Safety Inputs II", 64.0, 268.0);
                SetRecordContactor(oInsertedObjects, iMotor);

                //Layout
                Insert3DDeviceIntoDINRail(oProject, "U17", "-K10.3", 0, rightTo: "-K10.1");
                Function placement = insertDeviceLayout(oProject,"K10.3", "Control Outputs II", "M1", 0, 'B', "A", "Layout", 1259 + widthOffset, 763);
                widthOffset = widthOffset + placement.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
            }

            log = String.Concat(log, "\r\nIncluidos contactores");
        }

        private void Draw_Sincronismo()
        {
            string Modelo = (oElectric.CaractComercial["FMODELL"] as Caracteristic).CurrentReference;

            //Sensores superiores
            if (Modelo.Contains("CLASSIC"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Handrail_Speed_Sensors.ema", 'D', "Upper Sensors II", 48.0, 116.0);
                log = String.Concat(log, "\r\nIncluidos sensores Sincronismo en CI");
            }
            else
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Handrail_Speed_Sensors.ema", 'C', "Upper Sensors II", 48.0, 116.0);
                log = String.Concat(log, "\r\nIncluidos sensores Sincronismo en CI");
            }
        }

        private void Draw_VVF()
        {

            Caracteristic motorConnection = (Caracteristic)oElectric.CaractIng["CONEXMOTOR"];
            Caracteristic power = (Caracteristic)oElectric.CaractComercial["TNPOTENCIAMOTOR"];
            Caracteristic conexionMotor = (Caracteristic)oElectric.CaractIng["CONEXMOTOR"];
            Caracteristic bypass = (Caracteristic)oElectric.CaractComercial["TNCR_OT_BYPASS_VARIADOR"];

            if (conexionMotor.CurrentReference.Contains("VVF"))
            {
                insertPageMacro(oProject,"$(MD_MACROS)\\_Esquema\\1_Pagina\\VVF_Page.emp", "Control II", "VVF Power");

                //en página de "VVF Power
                char variante;
                if (bypass.CurrentReference.Equals("N"))
                    variante = 'E';
                else
                    variante = 'F';

                StorableObject[] oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", variante, "VVF Power", 24.0, 280.0);

                foreach (StorableObject oSOTemp in oInsertedObjects)
                {
                    //we are searching for PlaceHolder 'Three-Phase' in the results
                    PlaceHolder oPlaceHoldeThreePhase = oSOTemp as Eplan.EplApi.DataModel.Graphics.PlaceHolder;
                    if ((oPlaceHoldeThreePhase != null)
                        && (oPlaceHoldeThreePhase.Name == "VVF"))
                    {
                        if (power.NumVal <= 9.0)
                        {
                            oPlaceHoldeThreePhase.ApplyRecord("9kW");
                        }
                        else if (power.NumVal > 9.0 && power.NumVal <= 15.0)
                        {
                            oPlaceHoldeThreePhase.ApplyRecord("15kW");
                        }
                        else if (power.NumVal > 15.0 && power.NumVal <= 23.5)
                        {
                            oPlaceHoldeThreePhase.ApplyRecord("23,5kW");
                        }

                    }
                }

                //Insert VDF
                insertDeviceLayout(oProject,"VDF", "VVF Power", "M1", 0, 'A', "A", "Layout", 1160, 2080);
                Insert3DDeviceIntoMountingPlate(oProject, "-VDF", 30, 1600);
                //Insert Ferrita
                insertDeviceLayout(oProject,"VDF", "VVF Power", "M1", 1, 'A', "A", "Layout", 1197, 1562);
                Insert3DDeviceIntoMountingPlate(oProject, "-VDF", 70, 975, articleRef:1);

                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\VVF_Page_Control.emp", "Control II", "VVF Control");

                //en página de "VVF Control"
                if (!motorConnection.CurrentReference.Equals("VVF"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'G', "VVF Control", 48.0, 212.0);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'H', "VVF Control", 48.0, 212.0);

                //en página de "Control Outputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'O', "Control Outputs II", 236.0, 128.0);
                SetGECParameter(oProject, oElectric, "O2", (uint)GEC.Param.Speed_selection_output_2, true);

                //en página de "Control Inputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\VVF.ema", 'P', "Control Inputs I", 192.0, 160.0);
                SetGECParameter(oProject, oElectric, "I5", (uint)GEC.Param.VFD_EEC, true);

                if (!bypass.CurrentReference.Equals("N"))
                {
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Bypass_VVF.ema", 'A', "Control Inputs II", 64.0, 268.0);
                    SetGECParameter(oProject, oElectric, "I9", (uint)GEC.Param.Bypass_VFD, true);
                    insertDeviceLayout(oProject,"S4.4.1", "Control Inputs II", "M1", 1, 'A', "A", "Layout", "SB28");

                    Insert3DDeviceIntoDINRail(oProject, "U14", "-S4.4.1", 0, articleRef: 1);
                    Insert3DDeviceIntoComponent(oProject, "-S4.4.1", "-S4.4.1", "C", "MONTAJE BOTONERÍA", articleRef: 0);

                }

                log = String.Concat(log, "\r\nIncluido VVF");
            }

        }

        private void Draw_Temperatura_Motor()
        {
            int key;
            Insert oInsert = new Insert();

            Caracteristic motorType = (Caracteristic)oElectric.CaractComercial["FANTREHT"];

            //en página de "Motor"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'H', "Motor", 300.0, 44.0);

            if (motorType.CurrentReference.Equals("QC") ||
               motorType.CurrentReference.Equals("FJ"))
            {

                log = String.Concat(log, "\r\nIncluido relé termico");
            }
            else
            {
                //en página de "Motor"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Termal_Protection.ema", 'K', "Motor", 300.0, 196.0);

                log = String.Concat(log, "\r\nIncluido bimetal");
            }

            log = String.Concat(log, "\r\nIncluido relé termico");
        }

        private void Draw_LuzEstroboscopica(string type)
        {
            SetGECParameter(oProject, oElectric, "O5", (uint)GEC.Param.Lighting_1, true);
            switch (type)
            {
                //Una lampara
                case "STFSPALTBEL":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'I', "Upper Lighting I", 204.0, 176.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'D', "Lower Lighting I", 168.0, 136.0);

                    log = String.Concat(log, "\r\nIncluida luz estroboscópica: 1 Lampara");
                    break;

                //Dos lamparas
                case "STFSPALTBEL2":
                    //en página de "Upper Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'J', "Upper Lighting I", 204, 176.0);

                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Green_Light.ema", 'E', "Lower Lighting I", 168.0, 136.0);

                    log = String.Concat(log, "\r\nIncluida luz estroboscópica: 2 Lamparas");
                    break;

                default:
                    log = String.Concat(log, "\r\n****luz estroboscópica selecionada no disponible");
                    break;
            }
        }

        private void Draw_LuzPeines(string type, string luzEstro)
        {
            SetGECParameter(oProject, oElectric, "O5", (uint)GEC.Param.Lighting_1, true);
            if (type.Equals("DI"))
            {
                //en página de "Upper Lighting I"
                if (luzEstro.Equals("STFSPALTBEL") || luzEstro.Equals("STFSPALTBEL2"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'F', "Upper Lighting I", 252, 176.0);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'G', "Upper Lighting I", 204, 176.0);

                //en página de "Lower Lighting I"
                if (luzEstro.Equals("STFSPALTBEL") || luzEstro.Equals("STFSPALTBEL2"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'I', "Lower Lighting I", 176.0, 180.0);
                else if (luzEstro.Contains("LED"))
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'J', "Upper Lighting I", 176.0, 180.0);
                else
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Light.ema", 'K', "Upper Lighting I", 176.0, 180.0);

                log = String.Concat(log, "\r\nIncluida luz de peines LED 230V");
            }
            else
            {
                log = String.Concat(log, "\r\n****luz de peines selecionada no disponible");
            }
        }

        private void Draw_LuzBajopasamanos(string type, string luzEstro, string luzPeines)
        {
            SetGECParameter(oProject, oElectric, "O5", (uint)GEC.Param.Lighting_1, true);
            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'B', "Lower Lighting I", 156.0, 112.0);
                }
                else
                {
                    //Despues de la página de "Lower Lighting I"
                    insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                    //en página de "Lower Lighting II"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'D', "Lower Lighting II", 40.0, 188.0);

                    if (!luzEstro.Equals("KEINE") || !luzPeines.Equals("NO"))
                    {
                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Underhandrail_Light.ema", 'C', "Lower Lighting I", 176, 200.0);

                    }

                }


                log = String.Concat(log, "\r\nIncluida luz bajpasamanos LED 230V");
            }

            else
            {
                log = String.Concat(log, "\r\n****luz de bajo pasamanos selecionada no disponible");
            }

        }

        private void Draw_LuzZocalos(string type, string luzBajopasamanos, string luzEstro, string luzPeines)
        {
            SetGECParameter(oProject, oElectric, "O5", (uint)GEC.Param.Lighting_1, true);
            if (type.Equals("DIRECTA") || type.Equals("LED"))
            {
                if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && luzBajopasamanos.Equals("BELOHNE"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'A', "Lower Lighting I", 156.0, 112.0);
                }
                else if (luzEstro.Equals("KEINE") && luzPeines.Equals("NO") && !luzBajopasamanos.Equals("BELOHNE"))
                {
                    //en página de "Lower Lighting I"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'F', "Lower Lighting I", 228.0, 172.0);
                }
                else
                {
                    if (!luzBajopasamanos.Equals("BELOHNE"))
                    {
                        //en página de "Lower Lighting II"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'G', "Lower Lighting II", 132.0, 188.0);
                    }
                    else
                    {
                        //Despues de la página de "Lower Lighting I"
                        insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Lighting_II.emp", "Lower Lighting I", "Lower Lighting II");

                        //en página de "Lower Lighting II"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'H', "Lower Lighting II", 40.0, 188.0);

                        //en página de "Lower Lighting I"
                        insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Skirt_Light.ema", 'E', "Lower Lighting II", 176, 200.0);
                    }
                }


                log = String.Concat(log, "\r\nIncluida luz zócalo LED 230V");
            }

            else
            {
                log = String.Concat(log, "\r\n****luz de zocalo selecionada no disponible");
            }

        }

        private void Draw_LuzFoso(string type)
        {

            if (type.Equals("HANDL") ||
                type.Equals("OVAL"))
            {
                //en página de "Upper Lighting I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 'E', "Upper Lighting I", 68.0, 164.0);

                //en página de "Lower Lighting I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pit_Light.ema", 'F', "Lower Lighting I", 20.0, 176.0);

                log = String.Concat(log, "\r\nIncluida luz de foso manual");
            }
            else
            {
                log = String.Concat(log, "\r\n****luz de foso selecionada no disponible");
            }
        }

        private void Draw_Semaforo(string type)
        {

            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            //Semaforos superiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Oil pump & Traffic Lights.emp", "Upper Keys", "Oil pump & Traffic Lights");


            //Semaforos inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_Traffic_Lights.emp", "Lower Keys", "Lower Traffic Lights");

            SetGECParameter(oProject, oElectric, "UO1", (uint)GEC.Param.Top_traffic_light_red, true);
            SetGECParameter(oProject, oElectric, "UO2", (uint)GEC.Param.Top_traffic_light_green, true);
            SetGECParameter(oProject, oElectric, "LO1", (uint)GEC.Param.Bottom_traffic_light_red, true);
            SetGECParameter(oProject, oElectric, "LO2", (uint)GEC.Param.Bottom_traffic_light_green, true);

            if (modelo.CurrentReference.Contains("CLASSIC"))
            {
                //en página de "Upper Traffic Lights"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'R', "Oil pump & Traffic Lights", 192.0, 120.0);
                //en página de "Lower Traffic Lights"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'Q', "Lower Traffic Lights", 128.0, 148.0);
                log = String.Concat(log, "\nIncluidos Semáforos Chinos de VC3.0");
            }
            else
            {

                if (type.Equals("BICOLOR"))
                {
                    //en página de "Upper Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'G', "Oil pump & Traffic Lights", 228.0, 120.0);

                    //en página de "Lower Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'I', "Lower Traffic Lights", 128.0, 148.0);

                    log = String.Concat(log, "\r\nIncluidos Semáforos Bicolor");
                }
                else if (type.Equals("ROTGRUEN") ||
                        type.Equals("F6NEINB") ||
                        type.Equals("PROHI_VERDE"))
                {
                    //en página de "Upper Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'H', "Oil pump & Traffic Lights", 228.0, 120.0);

                    //en página de "Lower Traffic Lights"
                    insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Traffic_Light.ema", 'J', "Lower Traffic Lights", 128.0, 148.0);

                    log = String.Concat(log, "\r\nIncluidos Semáforos Flecha/Prohibido");
                }
                else
                {
                    log = String.Concat(log, "\r\n*****No existe macro para el tipo de semaforo seleccionado");
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
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'Q', "Upper People Detection", 200.0, 256.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'C', "Upper People Detection", 200.0, 256.0);

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
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'R', "Lower People Detection", 200.0, 256.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Radar.ema", 'D', "Lower People Detection", 200.0, 256.0);

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
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'J', "Upper People Detection", 12.0, 256.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'D', "Upper People Detection", 12.0, 256.0);

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'P', "Upper Diagnostic Inputs III", 400.0, 168.0);
            SetGECParameter(oProject, oElectric, "UI21", (uint)GEC.Param.Top_light_barrier_comb_plate_NC, true);

            //******************************************************************************************************************************************************
            //******************************************************************************************************************************************************

            //Fotocélulas inferiores
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Lower_People_Detection.emp", "Lower Diagnostic Outputs I", "Lower People Detection");

            if (producto.CurrentReference.Contains("CLASSIC"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'I', "Lower People Detection", 12.0, 256.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'C', "Lower People Detection", 12.0, 256.0);

            //en página de "Upper Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Comb_Photocell.ema", 'O', "Lower Diagnostic Inputs III", 400.0, 168.0);
            SetGECParameter(oProject, oElectric, "LI21", (uint)GEC.Param.Bottom_light_barrier_comb_plate_NC, true);

        }

        private void Draw_RoturaCadenaPrincipal()
        {

            //en página de "Upper Sensors I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 'B', "Upper Sensors I", 140.0, 264.0);


            //en página de "Safety Inputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Chain.ema", 'O', "Safety Inputs I", 324.0, 184.0);
            SetGECParameter(oProject, oElectric, "SI18", (uint)GEC.Param.Drive_chain_DuTriplex, true);
        }

        private void Draw_RoturaPasamanos()
        {
            //en página de "Lower Sensors I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'B', "Lower Sensors I", 160.0, 252.0);

            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Broken_Handrail.ema", 'O', "Lower Diagnostic Inputs IV", 216.0, 164.0);
            SetGECParameter(oProject, oElectric, "LI25", (uint)GEC.Param.Broken_handrail_L, true);
            SetGECParameter(oProject, oElectric, "LI26", (uint)GEC.Param.Broken_handrail_R, true);
        }

        private void Draw_Nivel_Agua()
        {

            //en página de "Lower Sensors I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Water_detection.ema", 'C', "Lower Sensors I", 352.0, 252.0);

            //en página de "Lower Diagnostic Inputs IV"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Water_detection.ema", 'D', "Lower Diagnostic Inputs IV", 336.0, 164.0);
            SetGECParameter(oProject, oElectric, "LI27", (uint)GEC.Param.Water_detection_bottom, true);
        }

        private void Draw_ControlDesgasteFrenos()
        {

            //en página de "Upper Sensors I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'B', "Upper Sensors I", 240.0, 264.0);


            //en página de "Upper Diagnostic Inputs IV"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_Wear.ema", 'P', "Upper Diagnostic Inputs IV", 340.0, 156.0);
            SetGECParameter(oProject, oElectric, "UI27", (uint)GEC.Param.Brake_wear_brake_1_M1, true);
            SetGECParameter(oProject, oElectric, "UI28", (uint)GEC.Param.Brake_wear_brake_2_M1, true);
        }

        private void Draw_BuggyInferior()
        {
            //en página de "Lower Diagnostic Inputs III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'O', "Lower Diagnostic Inputs III", 20.0, 156.0);
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'P', "Lower Diagnostic Inputs III", 72.0, 156.0);
            SetGECParameter(oProject, oElectric, "LI15", (uint)GEC.Param.Bottom_buggy_right_SS, true);
            SetGECParameter(oProject, oElectric, "LI16", (uint)GEC.Param.Bottom_buggy_left_SS, true);
        }

        private void Draw_BuggySuperior()
        {

            //en página de "Upper Diagnostic Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Buggy.ema", 'M', "Upper Diagnostic Inputs II", 372.0, 156.0);

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
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Vertical_Comb.ema", 'B', "Lower Diagnostic Inputs II", 308.0, 156.0);
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

        private void Draw_Freno(string motorType)
        {
            //en página de "Brake I"
            if (motorType.Equals("QC") ||
                motorType.Equals("FJ"))
            {
                //Finales de carrera
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_1.ema", 'B', "Brake I", 64.0, 96.0);
                log = String.Concat(log, "Finales de carrera");
            }
            else
            {
                //Inductivos
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_1.ema", 'A', "Brake I", 64.0, 96.0);
                log = String.Concat(log, "Inductivos");
            }

        }

        private void Draw_Freno_adicional(string motorType)
        {

            //en página de "Control II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'N', "Control II", 176.0, 160.0);
            

            //Segundo freno
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_II.emp", "Brake I", "Brake II");

            //en página de "Brake II"
            if (motorType.Equals("QC") ||
                motorType.Equals("FJ"))
            {
                //Finales de carrera
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'E', "Brake II", 64.0, 96.0);
                log = String.Concat(log, "Finales de carrera");
            }
            else
            {
                //Inductivos
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'D', "Brake II", 64.0, 96.0);
                log = String.Concat(log, "Inductivos");
            }

            //Layout
            insertDeviceLayout(oProject, "UN2", "Control II", "M1", 0, 'A', "A", "Layout", "UN1");
            Insert3DDeviceIntoDINRail(oProject, "U16", "-UN2", 0);
            //Conexiones
            //en página de "Safety Inputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Brake_2.ema", 'F', "Safety Inputs I", 228.0, 184.0);
            SetGECParameter(oProject, oElectric, "SI15", (uint)GEC.Param.Brake_function_brake_3_mot_1, true);
            SetGECParameter(oProject, oElectric, "SI16", (uint)GEC.Param.Brake_function_brake_4_mot_1, true);
        }

        private void Draw_Trinquete(string type)
        {
            int key;
            Caracteristic OilPump = (Caracteristic)oElectric.CaractComercial["TNCR_ENGRASE_AUTOMATICO"];
            Caracteristic DoubleBrake = (Caracteristic)oElectric.CaractComercial["FBREMSE2"];
            Insert oInsert = new Insert();

            //en página de "Control I"
            StorableObject[] oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'F', "Control I", 324.0, 196.0);
            foreach (StorableObject oSOTemp in oInsertedObjects)
            {
                //we are searching for PlaceHolder 'Three-Phase' in the results
                PlaceHolder oPlaceHoldeThreePhase = oSOTemp as PlaceHolder;
                if ((oPlaceHoldeThreePhase != null)
                    && (oPlaceHoldeThreePhase.Name == "UPS"))
                {
                    if (type.Equals("NAB"))
                        oPlaceHoldeThreePhase.ApplyRecord("NAB");
                    else
                        oPlaceHoldeThreePhase.ApplyRecord("AB");
                }
            }

            //Layout QF5.1, QF5.2 (SAI en la puerta)
            if (OilPump.CurrentReference.Equals("S") ||
                OilPump.CurrentReference.Equals("C"))
                insertDeviceLayout(oProject,"QF5.1", "Control I", "M1", 0, 'A', "A", "Layout", "Q4.2.7");
            else
                insertDeviceLayout(oProject,"QF5.1", "Control I", "M1", 0, 'A', "A", "Layout", "QD2");

            insertDeviceLayout(oProject,"QF5.2", "Control I", "M1", 0, 'A', "A", "Layout", "QF5.1");


            Insert3DDeviceIntoDINRail(oProject, "U13", "-QF5.1", 0);
            Insert3DDeviceIntoDINRail(oProject, "U13", "-QF5.2", 0);

            Dictionary<int, string> dictPages = GetPageTable(oProject);
            //en página de "Control II"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Control II");

            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.Y >= 256 &&
                    placement.Location.Y <= 264 &&
                    placement.Location.X >= 32 &&
                    placement.Location.X <= 48)
                {
                    placement.Remove();
                }
            }

            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'G', "Control II", 24.0, 248.0);

            //en página de "Control II"
            oInsertedObjects = insertWindowMacro_ObjCont(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'H', "Control II", 376.0, 240.0);

            foreach (StorableObject oSOTemp in oInsertedObjects)
            {
                //we are searching for PlaceHolder 'Three-Phase' in the results
                PlaceHolder oPlaceHoldeThreePhase = oSOTemp as PlaceHolder;
                if ((oPlaceHoldeThreePhase != null)
                    && (oPlaceHoldeThreePhase.Name == "AB"))
                {
                    if (type.Equals("NAB"))
                        oPlaceHoldeThreePhase.ApplyRecord("NAB");
                    else
                        oPlaceHoldeThreePhase.ApplyRecord("AB");
                }
            }

            //Layout UN3 KA2
            if (DoubleBrake.CurrentReference.Equals("4/4"))
                insertDeviceLayout(oProject, "UN3", "Control II", "M1", 0, 'A', "A", "Layout", "UN2");
            else
                insertDeviceLayout(oProject,"UN3", "Control II", "M1", 0, 'A', "A", "Layout", "UN1");

            insertDeviceLayout(oProject,"KA2", "Control II", "M1", 1, 'A', "A", "Layout", "UN3");

            Insert3DDeviceIntoDINRail(oProject, "U16", "-UN3", 0);
            Insert3DDeviceIntoDINRail(oProject, "U17", "-KA2", 0, articleRef: 1, leftTo:"-K25");
            Insert3DDeviceIntoComponent(oProject, "-KA2", "-KA2", "G1", "MOUNTING POINT", articleRef: 0);

            //en página de "Control Outputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'J', "Control Outputs I", 28.0, 88.0);


            //Tercer freno
            //Compruebo si ya esta insertada la página "Brake III"
            if (DoubleBrake.CurrentReference.Equals("4/4"))
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_III.emp", "Brake II", "Brake III");
            else
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Brake_III.emp", "Brake I", "Brake III");


            //en página de "Brake III"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'C', "Brake III", 324.0, 260.0);

            //en página de "Pulse Inputs"
            deleteArea(oProject, "Pulse Inputs", 256, 140, 288, 140);

            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Encoder.ema", 'A', "Pulse Inputs", 264.0, 180.0);
            SetGECParameter(oProject, oElectric, "SI7", (uint)GEC.Param.Main_shaft_speed_monitor_1, true);
            SetGECParameter(oProject, oElectric, "SI8", (uint)GEC.Param.Main_shaft_speed_monitor_2, true);

            //en página de "Safety Inputs I"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Pawl_Brake.ema", 'I', "Safety Inputs I", 292.0, 184.0);
            SetGECParameter(oProject, oElectric, "SI17", (uint)GEC.Param.Aux_brake_status_1, true);
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
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Main Shaft Maintenance Lock.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
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
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 132.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI17", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);
            }
            else if (sMicrosZocalo.NumVal <= 6 &&
                cerrojo.CurrentReference.Equals("N"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 76.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI16", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);

            }
            else if (sBuggy.CurrentReference.Equals("KEINE") ||
                    sBuggy.CurrentReference.Equals("BUGGYUT"))
            {

                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs III", 16.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI15", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);

            }
            else if (!sPeines.CurrentReference.Equals("INDEPENDIENTE"))
            {
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Mechanical_Pawl_Brake.ema", 'A', "Upper Diagnostic Inputs II", 312.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI13", (uint)GEC.Param.Mechanical_pawl_brake_SS, true);
            }
        }

        private void Draw_Lubricacion_auto()
        {
            Caracteristic AuxBrake = (Caracteristic)oElectric.CaractComercial["FZUSBREMSE"];

            //Compruebo si ya esta insertada la página "Oil pump & Traffic Lights"
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Oil pump & Traffic Lights.emp", "Upper Keys", "Oil pump & Traffic Lights");

            //en página de "Oil pump & Brake"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'C', "Oil pump & Traffic Lights", 16.0, 184.0);

            //en página de "Control Outputs II"
            if (AuxBrake.CurrentReference.Equals("HWSPERRKMAGN") || AuxBrake.CurrentReference.Equals("NAB"))
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'I', "Control Outputs III", 276.0, 184.0);
            else
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'N', "Control Outputs III", 276.0, 184.0);

            insertDeviceLayout(oProject,"Q4.2.7", "Control Outputs III", "M1", 0, 'A', "A", "Layout", "QD2");
            Insert3DDeviceIntoDINRail(oProject, "U13", "-Q4.2.7", 0);

            SetGECParameter(oProject, oElectric, "O9", (uint)GEC.Param.Oil_pump_activation, true);
            Caracteristic modelo = (Caracteristic)oElectric.CaractComercial["FMODELL"];
            if (modelo.CurrentReference.Contains("CLASSIC"))
                SetGECParameter(oProject, oElectric, "O9", (uint)GEC.Param.Oil_pump_control_1, true);


            //en página de "Control I
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'M', "Control II", 72.0, 100.0);


            //en página de "Control Inputs II"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Oil_Pump.ema", 'L', "Upper Diagnostic Inputs IV", 40.0, 160.0);
            SetGECParameter(oProject, oElectric, "UI22", (uint)GEC.Param.Oil_level_in_pump_1, true);
        }

        private void Draw_Llavin_Auto_Cont(String ubicacion, String controlador)
        {

            //Armario
            if (ubicacion.Equals("A"))
            {
                //Control Inputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'E', "Control Inputs II", 160.0, 268.0);

                Insert3DDeviceIntoDINRail(oProject, "U14", "-SA4", 0, articleRef: 2);
                Insert3DDeviceIntoComponent(oProject, "-SA4", "-SA4", "C", "MONTAJE BOTONERÍA", articleRef: 0);
                Insert3DDeviceIntoComponent(oProject, "-SA4", "-SA4", "C", "MONTAJE BOTONERÍA", articleRef: 1);

                SetGECParameter(oProject, oElectric, "I12", (uint)GEC.Param.Automatic_key_top, true);
                SetGECParameter(oProject, oElectric, "I13", (uint)GEC.Param.Continuous_key_top, true);
            }
            //En cubrezocalo o balaustrada
            else
            {
                //Upper Keys"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'I', "Upper Keys", 116.0, 264.0);

                //Upper Diagnostic Inputs IV"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'B', "Upper Diagnostic Inputs IV", 220.0, 156.0);
                SetGECParameter(oProject, oElectric, "UI26", (uint)GEC.Param.Automatic_key_top, true);
                SetGECParameter(oProject, oElectric, "UI25", (uint)GEC.Param.Continuous_key_top, true);
            }
        }

        private void Draw_Llavin_Local_Remoto(String ubicacion, String controlador)
        {

            //En Armario
            if (ubicacion.Equals("A"))
            {
                //Control Inputs I"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'F', "Control Inputs I", 256.0, 268.0);

                Insert3DDeviceIntoDINRail(oProject, "U14", "-SA3", 0, articleRef: 2);
                Insert3DDeviceIntoComponent(oProject, "-SA3", "-SA3", "C", "MONTAJE BOTONERÍA", articleRef: 0);
                Insert3DDeviceIntoComponent(oProject, "-SA3", "-SA3", "C", "MONTAJE BOTONERÍA", articleRef: 1);

                SetGECParameter(oProject, oElectric, "I7", (uint)GEC.Param.Local_key_top, true);
                SetGECParameter(oProject, oElectric, "I8", (uint)GEC.Param.Remote_key_top, true);
            }
            //En cubrezocalo o balaustrada
            else
            {
                //Upper Keys"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'J', "Upper Keys", 224.0, 264.0);

                //Upper Diagnostic Inputs IV
                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Diagnostic_Inputs_V_Arm_Ext.emp", "Upper Diagnostic Inputs IV", "Upper Diagnostic Inputs V");

                //Upper Diagnostic Inputs V"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'D', "Upper Diagnostic Inputs V", 248.0, 168.0);
                SetGECParameter(oProject, oElectric, "UI31", (uint)GEC.Param.Local_key_top, true);
                SetGECParameter(oProject, oElectric, "UI32", (uint)GEC.Param.Remote_key_top, true);
            }

        }

        private void Draw_Llavin_Paro(String ubicacion)
        {

            if (ubicacion.Equals("A"))
            {
                //Control Inputs II"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'H', "Control Inputs II", 224.0, 268.0);

                Insert3DDeviceIntoDINRail(oProject, "U14", "-SA5", 0, articleRef: 2);
                Insert3DDeviceIntoComponent(oProject, "-SA5", "-SA5", "C", "MONTAJE BOTONERÍA", articleRef: 0);
                Insert3DDeviceIntoComponent(oProject, "-SA5", "-SA5", "C", "MONTAJE BOTONERÍA", articleRef: 1);

                //SetGECParameter(oProject, oElectric, "I14", (uint)GEC.Param.Stop_switch_for_safety_curtain, true); //Pendiente de correccion de bug. Previsto en R15
                SetGECParameter(oProject, oElectric, "I14", (uint)GEC.Param.Top_operational_stop_local_B16, true);
            }
            else
            {
                //Upper Keys"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'K', "Upper Keys", 332.0, 264.0);

                insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Upper_Diagnostic_Inputs_V_Arm_Ext.emp", "Upper Diagnostic Inputs IV", "Upper Diagnostic Inputs V");

                //Upper Diagnostic Inputs V"
                insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Additional_Keys.ema", 'L', "Upper Diagnostic Inputs V", 188.0, 168.0);
                //SetGECParameter(oProject, oElectric, "UI30", (uint)GEC.Param.Stop_switch_for_safety_curtain, true); //Pendiente de correccion de bug. Previsto en R15
                SetGECParameter(oProject, oElectric, "UI30", (uint)GEC.Param.Top_operational_stop_local_B16, true);

            }
        }

        private void Draw_Envolvente_Armario()
        {
            int key;

            Caracteristic Envolvente = (Caracteristic)oElectric.CaractIng["ENVOLV_ARM_EXT"];

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            //Main Supply"
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Main Supply");

            Page page = oProject.Pages[key];

            foreach (Function function in page.Functions)
            {
                String name = function.Name;
            }

            page.Filter.FunctionCategory = FunctionCategory.ArticleDefinitionPoint;
            Function[] f = page.Functions;
            Function Arm = null;

            foreach (Function function in f)
            {
                if (function.VisibleName.Equals("ARM"))
                {
                    Arm = function;
                    break;
                }
            }

            if (Arm == null)
                return;

            switch (Envolvente.CurrentReference)
            {
                case "M_1800x800x400":
                    Arm.AddArticleReference("IDE.FSC1808040PO/SP");
                    Arm.AddArticleReference("IDE.PL18080");
                    break;

                case "I_1800x800x400":
                    Arm.AddArticleReference("DELV.MVAC188040");
                    break;

                case "M_1800x1000x400":
                    Arm.AddArticleReference("IDE.FSC18010040PO/SP");
                    Arm.AddArticleReference("IDE.PL180100");
                    break;

                case "I_1800x1000x400":
                    Arm.AddArticleReference("DELV.MVAC181040");
                    break;

                case "M_1800x1200x400":
                    Arm.AddArticleReference("IDE.BIG18012040POD");
                    break;
            }
        }

        private void Draw_Protecciones()
        {
            int key;

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == "Main Supply");
            Caracteristic iMotor = (Caracteristic)oElectric.CaractIng["IMOTOR"];

            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.TypeIdentifier == 106)
                {
                    PlaceHolder placeHolder = (PlaceHolder)placement;
                    if (placeHolder.Name.Equals("QF1"))
                    {
                        if (iMotor.NumVal < 38)
                            placeHolder.ApplyRecord("40");
                        else
                            placeHolder.ApplyRecord("63");
                    }
                }
            }
        }

        private void Draw_Safety_Curtain()
        {

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Curtain.emp", "Display", "Safety_Curtain");

            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\Safety_Extension_Inputs.emp", "Safety Inputs II", "Safety Extension Inputs I");
            SetGECParameter(oProject, oElectric, "SI44", (uint)GEC.Param.Safety_curtain, true);
            SetGECParameter(oProject, oElectric, "SI45", (uint)GEC.Param.Switch_disable_safety_curtain, true);

            //en página de "Display"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain_Arm_Ext.ema", 'C', "Display", 124.0, 272.0);

            //en página de "Safety Curtain"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain_Arm_Ext.ema", 'A', "Safety Curtain", 52.0, 224.0);

            //en página de "Safety Curtain"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain.ema", 'B', "Safety Curtain", 52.0, 120.0);

            //en página de "Safety Curtain"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\Safety_Curtain_Arm_Ext.ema", 'B', "Safety Curtain", 16.0, 288.0);

        }

        private void Draw_PLC()
        {
            Caracteristic armario = (Caracteristic)oElectric.CaractIng["ENVOLV_ARM_EXT"];

            //Despues de la página de "Communication
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\PLC_Input_I.emp", "Communication", "PLC Input I");
            if (armario.CurrentReference.Contains("1800x800"))
                insertDeviceLayout(oProject,"A1", "PLC Input I", "M1", 0, 'A', "A", "Layout", 1539, 1415);
            if (armario.CurrentReference.Contains("1800x1000") ||
               armario.CurrentReference.Contains("1800x1200"))
                insertDeviceLayout(oProject,"A1", "PLC Input I", "M1", 0, 'A', "A", "Layout", 1825, 1250);

            Insert3DDeviceIntoDINRail(oProject, "U16", "-A1", 0);

            //Despues de la página de "PLC Input I"
            insertPageMacro(oProject, "$(MD_MACROS)\\_Esquema\\1_Pagina\\PLC_Output_I.emp", "PLC Input I", "PLC Output I");

            //en página de "Display"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\PLC.ema", 'A', "Display", 216.0, 252.0);

            if (armario.CurrentReference.Contains("1800x800"))
                insertDeviceLayout(oProject,"U3", "Display", "M1", 0, 'A', "A", "Layout", 1337, 997);

            if (armario.CurrentReference.Contains("1800x1000") ||
                armario.CurrentReference.Contains("1800x1200"))
                insertDeviceLayout(oProject,"U3", "Display", "M1", 0, 'A', "A", "Layout", 1525, 1253);

            Insert3DDeviceIntoDINRail(oProject, "U16", "-A1", 0);
            Insert3DDeviceIntoDINRail(oProject, "U16", "-U3", 0);



            //en página de "Communication"
            insertWindowMacro(oProject, "$(MD_MACROS)\\_Esquema\\2_Ventana\\PLC.ema", 'B', "Communication", 68.0, 180.0);

        }

        #endregion
    
    }

}
