using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using EPLAN_API.API_Basic;
using EPLAN_API.SAP;
using EPLAN_API.User;
using Microsoft.VisualBasic.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.Services.Protocols;
using System.Windows.Forms;

namespace EPLAN_API_2022.Forms
{
    public partial class Configurador : Form
    {
        public List<Electric> oElectricList;
        public Project oProject;
        public Draw Draw;
        public DrawTools DrawTools;
        public PanelGEC PGEC;

        public Configurador()
        {
            DrawTools = new DrawTools();

            oElectricList = new List<Electric>();
            oElectricList.Add(new Electric());
            oElectricList[0].ComboboxDataToConfigurador += new Electric.ComboboxDelegateToConfigurador(ComboboxChanged);
            oElectricList[0].TextBoxDataToConfigurador += new Electric.TextboxDelegateToConfigurador(TextboxChanged);

            //Try to open a project
            oProject = new ProjectManager().GetCurrentProjectWithDialog();

            if (oProject != null)
            {
                ReadCaractFromProject();
            }

            InitializeComponent();

            AbrirFormEnPanel(new PanelCaracComercial(oElectricList));

            LoadNameFromProject();

        }

        #region EPLAN Functions
        private void ReadCaractFromProject()
        {
            //Comercial
            foreach (Caracteristic c in oElectricList[0].CaractComercial.Values)
            {
                try
                {
                    AnyPropertyId propertyId = new AnyPropertyId(oProject, c.NameReference);
                    PropertyValue propertyValue = oProject.Properties[propertyId];

                    if (!c.IsNumeric)
                    {
                        c.setActualValue(propertyValue.ToMultiLangString().GetString(ISOCode.Language.L_en_US));
                    }
                    else
                    {
                        CultureInfo esES = new CultureInfo("es-ES");
                        CultureInfo enUS = new CultureInfo("en-US");

                        LanguageList languageList = new LanguageList();
                        propertyValue.ToMultiLangString().GetLanguageList(ref languageList);


                        string Svalue = propertyValue.ToMultiLangString().GetString(ISOCode.Language.L___);

                        if (languageList.Contains(ISOCode.Language.L_en_US))
                        {
                            Svalue = propertyValue.ToMultiLangString().GetString(ISOCode.Language.L_en_US);
                        }

                        if (Svalue.Contains(","))
                            c.setActualNumVal(Convert.ToDouble(Svalue, esES));
                        else
                            c.setActualNumVal(Convert.ToDouble(Svalue, enUS));
                    }

                }
                catch (Exception ex)
                {
                    ;
                }
            }

            //Ingenieria
            foreach (Caracteristic c in oElectricList[0].CaractIng.Values)
            {
                try
                {
                    AnyPropertyId propertyId = new AnyPropertyId(oProject, c.NameReference);
                    PropertyValue propertyValue = oProject.Properties[propertyId];

                    if (!c.IsNumeric)
                    {
                        c.setActualValue(propertyValue.ToMultiLangString().GetString(ISOCode.Language.L_en_US));
                    }
                    else
                    {
                        CultureInfo esES = new CultureInfo("es-ES");
                        CultureInfo enUS = new CultureInfo("en-US");

                        LanguageList languageList = new LanguageList();
                        propertyValue.ToMultiLangString().GetLanguageList(ref languageList);


                        string Svalue = propertyValue.ToMultiLangString().GetString(ISOCode.Language.L___);

                        if (languageList.Contains(ISOCode.Language.L_en_US))
                        {
                            Svalue = propertyValue.ToMultiLangString().GetString(ISOCode.Language.L_en_US);
                        }

                        if (Svalue.Contains(","))
                            c.setActualNumVal(Convert.ToDouble(Svalue, esES));
                        else
                            c.setActualNumVal(Convert.ToDouble(Svalue, enUS));
                    }

                }
                catch (Exception ex)
                {
                    ;
                }
            }


        }

        private void LoadNameFromProject()
        {
            try
            {
                this.Text = ((Caracteristic)oElectricList[0].CaractComercial["TNCR_COM_NOMBREOBRA_VBACK"]).TextVal;
                this.tB_OE.Text = ((Caracteristic)oElectricList[0].CaractComercial["TNCR_COM_COD_PEDIDO_CLIENTE"]).TextVal;
            }
            catch (Exception ex)
            {
                ;
            }
        }

        private void LoadCaractToProject()
        {
            //Commercial
            foreach (Caracteristic c in oElectricList[0].CaractComercial.Values)
            {
                try
                {
                    AnyPropertyId propertyId = new AnyPropertyId(oProject, c.NameReference);
                    PropertyValue propertyValue = oProject.Properties[propertyId];

                    if (!c.IsText)
                    {
                        if (!c.IsNumeric)
                        {
                            MultiLangString langString = new MultiLangString();
                            langString.AddString(ISOCode.Language.L_en_US, c.CurrentReference);
                            propertyValue.Set(langString);
                        }
                        else
                        {
                            CultureInfo esES = CultureInfo.CreateSpecificCulture("es-ES");
                            MultiLangString value = new MultiLangString();
                            value.AddString(ISOCode.Language.L_en_US, String.Format(esES, "{0:0.00}", c.NumVal));
                            propertyValue.Set(value);
                        }
                    }
                    else 
                    {
                        CultureInfo esES = CultureInfo.CreateSpecificCulture("es-ES");
                        MultiLangString value = new MultiLangString();
                        value.AddString(ISOCode.Language.L_en_US, c.TextVal);
                        propertyValue.Set(value);
                    }

                }
                catch (Exception ex)
                {
                    ;
                }
            }

            //Ingenieria
            foreach (Caracteristic c in oElectricList[0].CaractIng.Values)
            {
                try
                {
                    AnyPropertyId propertyId = new AnyPropertyId(oProject, c.NameReference);
                    PropertyValue propertyValue = oProject.Properties[propertyId];

                    if (!c.IsText)
                    {
                        if (!c.IsNumeric)
                        {
                            MultiLangString langString = new MultiLangString();
                            langString.AddString(ISOCode.Language.L_en_US, c.CurrentReference);
                            propertyValue.Set(langString);
                        }
                        else
                        {
                            CultureInfo esES = CultureInfo.CreateSpecificCulture("es-ES");
                            MultiLangString value = new MultiLangString();
                            value.AddString(ISOCode.Language.L_en_US, String.Format(esES, "{0:0.00}", c.NumVal));
                            propertyValue.Set(value);
                        }
                    }
                    else
                    {
                        CultureInfo esES = CultureInfo.CreateSpecificCulture("es-ES");
                        MultiLangString value = new MultiLangString();
                        value.AddString(ISOCode.Language.L_en_US, c.TextVal);
                        propertyValue.Set(value);
                    }

                }
                catch (Exception ex)
                {
                    ;
                }
            }

        }

        public void LoadSAPtoEPLAN(Dictionary<string, string> SAPCararct)
        {
            this.Text=SAPCararct["TNCR_COM_NOMBREOBRA_VBACK"].ToString();
            this.tB_OE.Text = SAPCararct["TNCR_COM_COD_PEDIDO_CLIENTE"].ToString();
            foreach (string key in SAPCararct.Keys.Where(n => n != null))
            {
                Caracteristic c = oElectricList[0].CaractComercial[key] as Caracteristic;

                if (c != null)
                {
                    if (!c.IsNumeric && SAPCararct[key] != null)
                    {
                        c.setActualValue(SAPCararct[key]);
                    }
                    else if (c.IsNumeric && SAPCararct[key] != null)
                    {
                        CultureInfo esES = new CultureInfo("es-ES");
                        CultureInfo enUS = new CultureInfo("en-US");

                        string Svalue = SAPCararct[key].TrimStart();
                        if (Svalue.Split(' ').Length > 0)
                            Svalue = Svalue.Split(' ')[0];

                        c.setActualNumVal(Convert.ToDouble(Svalue, esES));
                    }
                    else if (c.IsNumeric && SAPCararct[key] == null)
                    {
                        c.setActualNumVal(0);
                    }

                }

            }
        }

        public void UpdateSpecialCaract()
        {
            Caracteristic Producto = (Caracteristic)oElectricList[0].CaractComercial["FMODELL"];           // Modelo de producto
            Caracteristic PasoCadena = (Caracteristic)oElectricList[0].CaractComercial["FSTFKT"];          // Paso de cadena (mm)
            Caracteristic Inclinacion = (Caracteristic)oElectricList[0].CaractComercial["FNEIGUNG"];       // Angulo de inclinacion (º)
            Caracteristic desarrollo = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_DESARROLLO"];
            Caracteristic lCS = (Caracteristic)oElectricList[0].CaractComercial["FOT"];
            Caracteristic lCI = (Caracteristic)oElectricList[0].CaractComercial["FUT"];
            Caracteristic desnivel = (Caracteristic)oElectricList[0].CaractComercial["FHOEHEV"];
            Caracteristic sbuggy = (Caracteristic)oElectricList[0].CaractComercial["F04ZUB"];
            Caracteristic incClablesInf = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_INCREM_CABLES_INFERIOR"];
            Caracteristic incCablesSup = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_INCREM_CABLES_SUPERIOR"];
            Caracteristic voltaje = (Caracteristic)oElectricList[0].CaractComercial["TNCR_S_TENSION_N"];
            Caracteristic tensionRed = (Caracteristic)oElectricList[0].CaractComercial["FSPANNUNG"];
            Caracteristic stopCarritos = (Caracteristic)oElectricList[0].CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic IluminacionPeines = (Caracteristic)oElectricList[0].CaractComercial["ILUPEI"];
            Caracteristic IluminacionPeinesVC3 = (Caracteristic)oElectricList[0].CaractComercial["F35ZUB1"];
            Caracteristic modoFunc = (Caracteristic)oElectricList[0].CaractComercial["FBETRART"];
            Caracteristic llavAutoCont = (Caracteristic)oElectricList[0].CaractComercial["LLAVES_AUT_CONT"];
            Caracteristic llavLocalRem = (Caracteristic)oElectricList[0].CaractComercial["LLAVES_LOCAL_REM"];
            Caracteristic llavParo = (Caracteristic)oElectricList[0].CaractComercial["LLAVES_PARO"];
            Caracteristic sisteAnden = (Caracteristic)oElectricList[0].CaractComercial["FWIEDERB"];
            Caracteristic stopAdicional = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            Caracteristic pais = (Caracteristic)oElectricList[0].CaractComercial["FLAND"];
            Caracteristic semaforo = (Caracteristic)oElectricList[0].CaractComercial["FAMPELSYM"];
            Caracteristic ubicacionControlador = (Caracteristic)oElectricList[0].CaractComercial["F53ZUB7"];
            Caracteristic distArmario = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_DISTANCIA_ARMARIO"];
            Caracteristic microsZocalo = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_NUM_MICROCONT"];

            //Especiales para Orinoco Classic
            if (Producto.CurrentReference.Equals("ORINOCO") &&
                Inclinacion.NumVal == 12)
            {
                //Buggy: Los pasillos no llevan buggy
                sbuggy.setActualValue("KEINE");

                //Paso cadena
                PasoCadena.setActualValue("101,15");
            }

            //Stop para carritos solo si es pasillo
            switch (Producto.CurrentReference)
            {
                case "TUGELA":
                case "VELINO":
                case "VICTORIA":
                case "TNE35L":
                case "VELINO_CLASSIC":
                    stopCarritos.setActualValue("KEINE");
                    break;

            }

            //Iluminación de peines en VC3.0

            if (IluminacionPeines.CurrentReference == null)
            {
                if (IluminacionPeinesVC3.CurrentReference != null)
                {
                    if (IluminacionPeinesVC3.CurrentReference.Equals("KAMMBELLD"))
                    {
                        IluminacionPeines.setActualValue("DI");
                        IluminacionPeines.combobox.Enabled = true;
                    }
                    else if (IluminacionPeinesVC3.CurrentReference.Equals("KEINE"))
                    {
                        IluminacionPeines.setActualValue("NO");
                        IluminacionPeines.combobox.Enabled = true;
                    }
                }
            }

            //Llavin de automatico/Continuo
            if (!modoFunc.CurrentReference.Equals("SG") &&
                !modoFunc.CurrentReference.Equals("SGBV") &&
                !modoFunc.CurrentReference.Equals("INTERM") &&
                llavAutoCont.CurrentReference == null)
            {
                llavAutoCont.setActualValue("N");
            }

            //Llavin de paro (Orinoco)
            if (Producto.CurrentReference.Equals("ORINOCO") &&
                llavLocalRem.CurrentReference == null)
            {
                llavLocalRem.setActualValue("N");
            }

            //Llavin de local/remoto (orinoco)
            if (llavParo.CurrentReference == null &&
                Producto.CurrentReference.Equals("ORINOCO") &&
                sisteAnden.CurrentReference.Equals("KEINE"))
            {
                llavParo.setActualValue("N");
            }

            //Stop Adicional (orinoco)
            if (Producto.CurrentReference.Equals("ORINOCO") &&
                stopAdicional.CurrentReference == null)
            {
                if (pais.CurrentReference.Equals("1") &&
                    pais.CurrentReference.Equals("2"))
                    stopAdicional.setActualValue("2");
                else
                    stopAdicional.setActualValue("N");
            }

            //Semáforos VC3.0
            if (Producto.CurrentReference.Contains("CLASSIC") &&
                semaforo.CurrentReference == null)
            {
                semaforo.setActualValue("NINGUNO");
            }

            //Distancia en metros al armario
            if (ubicacionControlador.CurrentReference.Equals("INNENOBEN"))
            {
                distArmario.setActualNumVal(0);
            }



        }

        public void CalculateCaractIng()
        {
            Caracteristic potencia = (Caracteristic)oElectricList[0].CaractComercial["TNPOTENCIAMOTOR"];
            Caracteristic tension = (Caracteristic)oElectricList[0].CaractComercial["FSPANNUNG"];
            Caracteristic VVF = (Caracteristic)oElectricList[0].CaractComercial["TNCR_SD_SIST_AHORRO"];
            Caracteristic Bypass = (Caracteristic)oElectricList[0].CaractComercial["TNCR_OT_BYPASS_VARIADOR"];
            Caracteristic tipoMotor = (Caracteristic)oElectricList[0].CaractComercial["FANTREHT"];
            Caracteristic sAnden = (Caracteristic)oElectricList[0].CaractComercial["FWIEDERB"];
            Caracteristic ubicacionControlador = (Caracteristic)oElectricList[0].CaractComercial["F53ZUB7"];
            Caracteristic pais = (Caracteristic)oElectricList[0].CaractComercial["FLAND"];
            Caracteristic DenominacionObra = (Caracteristic)oElectricList[0].CaractComercial["TNCR_COM_NOMBREOBRA_VBACK"];

            Caracteristic controller = (Caracteristic)oElectricList[0].CaractIng["TNCR_DO_CONTROL"];
            Caracteristic cSeccMotor = (Caracteristic)oElectricList[0].CaractIng["SECCABLEMOT"];
            Caracteristic ConexMotor = (Caracteristic)oElectricList[0].CaractIng["CONEXMOTOR"];
            Caracteristic cMotor = (Caracteristic)oElectricList[0].CaractIng["IMOTOR"];
            Caracteristic cTermico = (Caracteristic)oElectricList[0].CaractIng["ITERMICO"];
            Caracteristic armario = (Caracteristic)oElectricList[0].CaractIng["ARMARIO"];
            Caracteristic maniobra = (Caracteristic)oElectricList[0].CaractIng["MANIOBRA"];
            Caracteristic tipoDisplay = (Caracteristic)oElectricList[0].CaractIng["TNCR_DO_DISPLAY_TYPE"];
            Caracteristic envolvente = (Caracteristic)oElectricList[0].CaractIng["ENVOLV_ARM_EXT"];
            Caracteristic paqueteEsp = (Caracteristic)oElectricList[0].CaractIng["PAQUETE_ESP"];
            Caracteristic rearranqueTrasCorteTension = (Caracteristic)oElectricList[0].CaractIng["POWER_OUTAGE_RESTART"];


            if (tension.CurrentReference != null &&
                potencia.NumVal > 0 &&
                Bypass.CurrentReference != null &&
                VVF.CurrentReference != null &&
                tipoMotor.CurrentReference != null &&
                sAnden.CurrentReference != null &&
                ubicacionControlador.CurrentReference != null)
            {



                #region  Calculate Motor conection
                double tensiond;
                Double.TryParse(tension.CurrentReference.Split('/')[1], out tensiond);

                bool YD = false;

                //Si armario interior VVF con bypass YD
                if (ubicacionControlador.CurrentReference.Equals("INNENOBEN") ||
                    ubicacionControlador.CurrentReference.Equals("INNENUNTEN"))
                {
                    ConexMotor.setActualValue("VVF_YD");
                }
                else
                {

                    if ((potencia.NumVal >= 7.5 && tensiond < 480) ||
                        (potencia.NumVal >= 9 && tensiond >= 480))
                        YD = true;

                    if (VVF.CurrentReference.Equals("VA") ||
                        VVF.CurrentReference.Equals("VA_REG"))
                    {
                        if (!Bypass.CurrentReference.Equals("N"))
                        {
                            if (YD)
                            {
                                //VVF con bypass estrella-triangulo
                                ConexMotor.setActualValue("VVF_YD");
                            }
                            else
                            {
                                //VVF con bypass triangulo
                                ConexMotor.setActualValue("VVF_D");
                            }
                        }
                        else
                        {
                            //VVF sin bypass
                            ConexMotor.setActualValue("VVF");
                        }
                    }
                    else
                    {
                        if (YD)
                        {
                            //Estrella-Triangulo sin VVF
                            ConexMotor.setActualValue("YD");
                        }
                        else
                        {
                            //Triangulo sin VVF
                            ConexMotor.setActualValue("D");
                        }
                    }
                }
                #endregion

                #region  Calculate Motor Current
                double corrienteMotor = 0;

                //Motor chino
                if (tipoMotor.CurrentReference.Equals("QC") ||
                    tipoMotor.CurrentReference.Equals("FJ"))
                {
                    if (potencia.NumVal <= 9)
                    {
                        corrienteMotor = 18.5;
                    }
                    else if (potencia.NumVal <= 15)
                    {
                        corrienteMotor = 27.5;
                    }
                    else if (potencia.NumVal <= 23.5)
                    {
                        corrienteMotor = 44;
                    }

                }
                //Motor local
                else
                {
                    #region Data Motor
                    List<MotorCurrentData> motorCurrentData = new List<MotorCurrentData>();
                    motorCurrentData.Add(new MotorCurrentData(4.5, "120/208", 60, 19.30));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "127/220", 50, 17.80));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "127/220", 60, 18.30));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "133/230", 50, 17.00));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "220/380", 50, 10.30));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "220/380", 60, 10.60));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "230/400", 50, 9.80));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "240/415", 50, 9.40));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "278/480", 60, 8.40));
                    motorCurrentData.Add(new MotorCurrentData(4.5, "347/600", 60, 6.70));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "120/208", 60, 26.30));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "127/220", 50, 24.30));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "127/220", 60, 24.80));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "133/230", 50, 23.30));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "220/380", 50, 14.10));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "220/380", 60, 14.40));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "230/400", 50, 13.40));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "240/415", 50, 12.90));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "278/480", 60, 11.40));
                    motorCurrentData.Add(new MotorCurrentData(6.0, "347/600", 60, 9.10));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "120/208", 60, 33.40));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "127/220", 50, 30.40));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "127/220", 60, 31.60));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "133/230", 50, 29.10));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "220/380", 50, 17.60));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "220/380", 60, 18.30));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "230/400", 50, 16.70));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "240/415", 50, 16.10));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "266/460", 60, 15.50));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "278/480", 60, 14.50));
                    motorCurrentData.Add(new MotorCurrentData(7.5, "347/600", 60, 11.60));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "120/208", 60, 40.10));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "127/220", 50, 35.90));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "127/220", 60, 37.90));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "133/230", 50, 34.30));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "220/380", 50, 20.80));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "220/380", 60, 21.90));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "230/400", 50, 19.70));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "230/400", 60, 21.90));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "240/415", 50, 19.00));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "278/480", 60, 17.40));
                    motorCurrentData.Add(new MotorCurrentData(9.0, "347/600", 60, 13.90));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "120/208", 60, 46.00));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "127/220", 50, 43.00));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "127/220", 60, 43.00));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "133/230", 50, 41.00));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "220/380", 50, 25.10));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "220/38V0  3*380V+PE (+N)", 60, 24.90));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "230/400", 50, 23.80));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "240/415", 50, 23.00));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "278/480", 60, 19.70));
                    motorCurrentData.Add(new MotorCurrentData(10.5, "347/600", 60, 15.80));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "120/208", 60, 50.00));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "127/220", 50, 45.00));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "127/220", 60, 47.00));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "133/230", 50, 43.00));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "220/380", 50, 26.20));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "220/380", 60, 27.20));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "230/400", 50, 24.90));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "266/460", 60, 25.00));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "240/415", 50, 24.00));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "278/480", 60, 21.50));
                    motorCurrentData.Add(new MotorCurrentData(12.0, "347/600", 60, 17.20));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "120/208", 60, 55.00));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "127/220", 50, 52.00));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "127/220", 60, 52.00));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "133/230", 50, 49.00));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "220/380", 50, 30.00));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "220/380", 60, 30.30));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "230/400", 50, 28.40));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "240/415", 50, 27.30));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "278/480", 60, 24.00));
                    motorCurrentData.Add(new MotorCurrentData(13.5, "347/600", 60, 19.20));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "120/208", 60, 61.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "127/220", 50, 57.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "127/220", 60, 57.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "133/230", 50, 54.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "220/380", 50, 33.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "220/380", 60, 33.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "230/400", 50, 31.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "240/415", 50, 30.00));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "254/440", 60, 26.80));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "278/480", 60, 26.30));
                    motorCurrentData.Add(new MotorCurrentData(15.0, "347/600", 60, 21.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "120/208", 60, 67.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "127/220", 50, 62.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "127/220", 60, 64.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "133/230", 50, 59.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "220/380", 50, 36.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "220/380", 60, 37.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "230/400", 50, 34.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "240/415", 50, 33.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "254/440", 60, 34.00));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "278/480", 60, 29.20));
                    motorCurrentData.Add(new MotorCurrentData(16.5, "347/600", 60, 23.40));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "120/208", 60, 74.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "127/220", 50, 68.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "127/220", 60, 70.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "133/230", 50, 65.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "220/380", 50, 40.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "220/380", 60, 40.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "230/400", 50, 38.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "240/415", 50, 36.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "254/440", 60, 34.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "266/460", 60, 33.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "278/480", 60, 32.00));
                    motorCurrentData.Add(new MotorCurrentData(19.0, "347/600", 60, 25.60));
                    motorCurrentData.Add(new MotorCurrentData(23.5, "127/220", 50, 85.00));
                    motorCurrentData.Add(new MotorCurrentData(23.5, "133/230", 50, 81.00));
                    motorCurrentData.Add(new MotorCurrentData(23.5, "220/380", 50, 49.00));
                    motorCurrentData.Add(new MotorCurrentData(23.5, "230/400", 50, 47.00));
                    motorCurrentData.Add(new MotorCurrentData(23.5, "240/415", 50, 45.00));
                    motorCurrentData.Add(new MotorCurrentData(23.5, "254/440", 60, 42.00));
                    motorCurrentData.Add(new MotorCurrentData(27.0, "127/220", 50, 94.00));
                    motorCurrentData.Add(new MotorCurrentData(27.0, "133/230", 50, 90.00));
                    motorCurrentData.Add(new MotorCurrentData(27.0, "220/380", 50, 55.00));
                    motorCurrentData.Add(new MotorCurrentData(27.0, "230/400", 50, 52.00));
                    motorCurrentData.Add(new MotorCurrentData(27.0, "240/415", 50, 50.00));
                    #endregion
                    foreach (MotorCurrentData motorCurrent in motorCurrentData)
                    {
                        if (motorCurrent.Power == potencia.NumVal &&
                            motorCurrent.Voltage == tension.CurrentReference)
                        {
                            corrienteMotor = motorCurrent.Current;
                            break;
                        }
                    }
                }

                cMotor.setActualNumVal(corrienteMotor);
                #endregion

                #region  Calculate Motor conection
                if (ConexMotor.CurrentReference.Equals("YD") ||
                    ConexMotor.CurrentReference.Equals("VVF_YD"))
                {
                    corrienteMotor = corrienteMotor / 1.732;
                }

                cTermico.setActualNumVal(corrienteMotor);
                #endregion

                #region Calculate motor cable secction
                //Proteccion principal de 25A (20A Motor y 5A mando)
                if (cMotor.NumVal <= 20)
                {
                    if (VVF.CurrentReference.Equals("VA") ||
                       VVF.CurrentReference.Equals("VA_REG"))
                        cSeccMotor.setActualValue("4A"); // 4mm2 apantallado
                    else
                        cSeccMotor.setActualValue("4"); // 4mm2 sin apantallar
                }

                //Proteccion principal de 32A (27.5A Motor (incluye motor chino 15kW) y 4.5A mando)
                else if (cMotor.NumVal <= 27.5)
                {
                    if (VVF.CurrentReference.Equals("VA") ||
                       VVF.CurrentReference.Equals("VA_REG"))
                        cSeccMotor.setActualValue("6A"); // 4mm2 apantallado
                    else
                        cSeccMotor.setActualValue("6"); // 4mm2 sin apantallar
                }

                //Proteccion principal de 40A (35A Motor y 5A mando)
                else if (cMotor.NumVal <= 35)
                {
                    if (VVF.CurrentReference.Equals("VA") ||
                       VVF.CurrentReference.Equals("VA_REG"))
                        cSeccMotor.setActualValue("10A"); // 4mm2 apantallado
                    else
                        cSeccMotor.setActualValue("10"); // 4mm2 sin apantallar
                }

                //Proteccion principal de 63A (58A Motor y 5A mando)
                else if (cMotor.NumVal <= 58)
                {
                    if (VVF.CurrentReference.Equals("VA") ||
                       VVF.CurrentReference.Equals("VA_REG"))
                        cSeccMotor.setActualValue("16A"); // 4mm2 apantallado
                    else
                        cSeccMotor.setActualValue("16"); // 4mm2 sin apantallar
                }

                #endregion

                #region Calculate Controller
                if (!sAnden.CurrentReference.Equals("KEINE"))
                {
                    controller.setActualValue("GEC+PLC");
                }
                else
                {
                    controller.setActualValue("GEC");
                }
               
                #endregion

                #region Calculate tipo display
                if (controller.CurrentReference.Equals("GEC"))
                    tipoDisplay.setActualValue("DDU");
                #endregion

                #region Paquete especial
                if (pais.CurrentReference.Equals("1") ||
                    pais.CurrentReference.Equals("2"))
                {
                    paqueteEsp.setActualValue("FRANCIA");
                }
                else if ((pais.CurrentReference.Equals("10") ||
                        (pais.CurrentReference.Equals("11")) &&
                    DenominacionObra.TextVal.Contains("MERCADONA")))
                {
                    paqueteEsp.setActualValue("MERCADONA");
                }
                else
                {
                    paqueteEsp.setActualValue("NO");
                }
                #endregion

                #region Envolvente Armario Exterior
                if (ubicacionControlador.CurrentReference.Equals("INNENOBEN"))
                {
                    envolvente.setActualValue("CUSTOM");
                }
                #endregion

                #region Rearranque tras corte de tension
                if (sAnden.CurrentReference.Equals("KEINE"))
                {
                    rearranqueTrasCorteTension.setActualValue("NO");
                }
                #endregion

                #region Default Values
                maniobra.setActualValue("ESTANDAR");
                #endregion
            }
        }

        #endregion

        #region GUI Funtions
        private void AbrirFormEnPanel(object Formhijo)
        {
            if (this.panelApp.Controls.Count > 0)
                this.panelApp.Controls.RemoveAt(0);
            Form fh = Formhijo as Form;
            fh.TopLevel = false;
            fh.Dock = DockStyle.Fill;
            this.panelApp.Controls.Add(fh);
            this.panelApp.Tag = fh;
            fh.Show();

        }

        private void ComboboxChanged(string data, string reference)
        {
            Caracteristic c;
            //Check if Commercial or Engineering
            c = oElectricList[0].CaractComercial[reference] as Caracteristic;
            if (c is null)
                c = oElectricList[0].CaractIng[reference] as Caracteristic;

            if (oProject != null)
            {
                AnyPropertyId propertyId = new AnyPropertyId(oProject, c.NameReference);
                PropertyValue propertyValue = oProject.Properties[propertyId];

                MultiLangString langString = new MultiLangString();
                langString.AddString(ISOCode.Language.L_en_US, c.CurrentReference);
                propertyValue.Set(langString);
            }
        }

        private void TextboxChanged(string data, string reference)
        {

            Caracteristic c;
            //Check if Commercial or Engineering
            c = oElectricList[0].CaractComercial[reference] as Caracteristic;
            if (c is null)
                c = oElectricList[0].CaractIng[reference] as Caracteristic;


            if (oProject != null)
            {
                AnyPropertyId propertyId = new AnyPropertyId(oProject, c.NameReference);
                PropertyValue propertyValue = oProject.Properties[propertyId];

                LanguageList languageList = new LanguageList();
                propertyValue.ToMultiLangString().GetLanguageList(ref languageList);
                propertyValue.ToMultiLangString().Clear();
                MultiLangString value = new MultiLangString();
                CultureInfo esES = new CultureInfo("es-ES");

                if (c.IsNumeric)
                {
                    value.AddString(ISOCode.Language.L_en_US, String.Format(esES, "{0:0.00}", c.NumVal));
                }
                else
                {
                    value.AddString(ISOCode.Language.L_en_US, c.TextVal);
                }
                propertyValue.Set(value);
            }
        }

        private void DrawDataToConfigurador(Project project)
        {
            oProject = project;
            LoadCaractToProject();
        }

        private void UpdateProgressBar(int progress)
        {
            this.TopMost = true;
            progressBar_Draw.Value=progress;
        }

        #endregion

        #region GUI Buttons
        private void BCaracComer_Click(object sender, EventArgs e)
        {
            AbrirFormEnPanel(new PanelCaracComercial(oElectricList));
        }

        private void BCaracIngClick(object sender, EventArgs e)
        {
            AbrirFormEnPanel(new PanelCaracIngenieria(oElectricList));
        }

        private void BGEC_Click(object sender, EventArgs e)
        {
            PGEC = new PanelGEC(oElectricList);
            AbrirFormEnPanel(PGEC);
            DrawTools.readGECfromEPLAN(oProject, oElectricList[0]);
            PGEC.UpdateGECData();
        }

        private void BRead_Click(object sender, EventArgs e)
        {
            if (tB_OE.Text.Length == 10)
            {
                try
                {
                    SAPService sAPService = new SAPService();
                    Dictionary<string, string> values = sAPService.ReadSAPCaract(tB_OE.Text);
                    if (values != null)
                        LoadSAPtoEPLAN(values);
                    else
                    {
                        LoadSAPtoEPLAN(new PortalContReader(tB_OE.Text).readCaracConfigElec());
                    }
                }
                catch
                {
                    LoadSAPtoEPLAN(new PortalContReader(tB_OE.Text).readCaracConfigElec());
                }

                UpdateSpecialCaract();
                CalculateCaractIng();
            }
        }

        private void BDraw_Click(object sender, EventArgs e)
        {
            DrawTools.calcParmGEC_Basic(oProject, oElectricList[0]);

            try
            {
                Draw = new Draw(oElectricList[0]);
                Draw.ProjectOpenedToConfigurador += new Draw.ProejctOpenedDelegate(DrawDataToConfigurador);
                Draw.ProgressChangedToConfigurador += new Draw.ProgressChangedDelegate(UpdateProgressBar);
                Draw.StartDrawing();
            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form() { TopMost = true, TopLevel = true }, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally 
            {
                this.TopMost = false; 
            }
        }

        private void BTest_Click(object sender, EventArgs e)
        {
            SAPService sAPService = new SAPService();
            sAPService.ReadSAPBOM("1150015558");
        }

        private void BCalc_Click(object sender, EventArgs e)
        {
            CalculateCaractIng();
        }

        private void BExport_Click(object sender, EventArgs e)
        {
            var filepath = "";

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    filepath = fbd.SelectedPath;
                }
            }

            filepath = String.Concat(filepath, "\\Export.csv");

            using (StreamWriter writer = new StreamWriter(new FileStream(filepath,
            FileMode.Create, FileAccess.Write)))
            {
                foreach (Caracteristic c in oElectricList[0].CaractComercial.Values)
                {
                    if (c.IsNumeric)
                        writer.WriteLine(String.Concat(c.NameReference, ";", c.NumVal.ToString(), ";", c.IsNumeric));
                    else
                    {
                        if (c.Values != null)
                            writer.WriteLine(String.Concat(c.NameReference, ";", c.CurrentReference, ";", c.IsNumeric));
                        else
                            writer.WriteLine(String.Concat(c.NameReference, ";", c.TextVal, ";", c.IsNumeric));
                    }

                }

                foreach (Caracteristic c in oElectricList[0].CaractIng.Values)
                {
                    if (c.IsNumeric)
                        writer.WriteLine(String.Concat(c.NameReference, ";", c.NumVal.ToString(), ";", c.IsNumeric));
                    else
                    {
                        if (c.Values != null)
                            writer.WriteLine(String.Concat(c.NameReference, ";", c.CurrentReference, ";", c.IsNumeric));
                        else
                            writer.WriteLine(String.Concat(c.NameReference, ";", c.TextVal, ";", c.IsNumeric));
                    }

                }
            }
        }

        private void BImport_Click(object sender, EventArgs e)
        {
            var filepath = "";

            using (var fbd = new OpenFileDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.FileName))
                {
                    filepath = fbd.FileName;
                }
            }



            using (var reader = new StreamReader(filepath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(';');
                    double res;

                    Caracteristic c = (Caracteristic)oElectricList[0].CaractComercial[values[0]];
                    if (c != null)
                    {
                        if (values[2].Equals("False"))
                            c.setActualValue(values[1]);
                        else
                        {
                            Double.TryParse(values[1], out res);
                            c.setActualNumVal(res);
                        }
                    }
                    else
                    {
                        c = (Caracteristic)oElectricList[0].CaractIng[values[0]];
                        if (values[2].Equals("False"))
                            c.setActualValue(values[1]);
                        else
                        {
                            Double.TryParse(values[1], out res);
                            c.setActualNumVal(res);
                        }
                    }

                }
            }

            LoadNameFromProject();
        }

        #endregion


    }
}
