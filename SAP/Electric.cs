using Eplan.EplApi.DataModel;
using EPLAN_API.User;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API.SAP
{
    public class Electric
    {
        public SortedList CaractComercial;
        public SortedList CaractIng;

        public SortedList OrderCaractComercial;
        public SortedList OrderCaractIng;

        public Dictionary<string, GEC> GECParameterList;
        public Dictionary<uint, string> IDFunctions;

        //Delegate to pass info to configurador
        public delegate void ComboboxDelegateToConfigurador(string data, string reference);

        //Event of delegate
        public event ComboboxDelegateToConfigurador ComboboxDataToConfigurador;

        //Delegate to pass info to configurador
        public delegate void TextboxDelegateToConfigurador(string data, string reference);

        //Event of delegate
        public event TextboxDelegateToConfigurador TextBoxDataToConfigurador;

        public Electric()
        {
            CreateCaractComercial();
            CreateCaractIng();
            createDefaultGECParam();

        }

        private void CreateCaractComercial()
        {
            CaractComercial = new SortedList();
            OrderCaractComercial = new SortedList();
            Caracteristic Caract;

            #region CONFIGURACION
            //Producto
            SortedList FMODELL = new SortedList();
            FMODELL.Add("TUGELA", new String[] { "Escalera Tugela", "Tugela" });
            FMODELL.Add("VELINO", new String[] { "Escalera Velino" });
            FMODELL.Add("CASCATA", new String[] { "Escalera Cascata" });
            FMODELL.Add("NIAGARA", new String[] { "Escalera Niagara" });
            FMODELL.Add("VICTORIA", new String[] { "Escalera Victoria" });
            FMODELL.Add("LOIRE", new String[] { "Pasillo Loire" });
            FMODELL.Add("AMAZONAS", new String[] { "Pasillo Amazonas" });
            FMODELL.Add("ORINOCO", new String[] { "Pasillo Orinoco", "Orinoco Inclinado", "Orinoco Horizontal" });
            FMODELL.Add("FT810", new String[] { "Escalera FT810" });
            FMODELL.Add("TNE35L", new String[] { "TNE35 APTA-L Victoria" });
            FMODELL.Add("IWALK", new String[] { "Pasillo Iwalk" });
            FMODELL.Add("IMOD", new String[] { "Imod" });
            FMODELL.Add("VELINO_CLASSIC", new String[] { "Escalera velino classic" });
            FMODELL.Add("TUGELA_CLASSIC", new String[] { "Escalera tugela classic" });
            Caract = new Caracteristic("FMODELL", "Producto", false, FMODELL, true, "-.6");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FMODELL", Caract);
            OrderCaractComercial.Add(10, Caract);

            //Documento de ventas
            Caract = new Caracteristic("TNCR_COM_COD_PEDIDO_CLIENTE", "Documento de ventas", true, "", false);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_COM_COD_PEDIDO_CLIENTE", Caract);
            OrderCaractComercial.Add(11, Caract);

            //Denominación de obra
            Caract = new Caracteristic("TNCR_COM_NOMBREOBRA_VBACK", "Denominación de obra",true,"",false);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_COM_NOMBREOBRA_VBACK", Caract);
            OrderCaractComercial.Add(12, Caract);

            //Desnivel
            Caract = new Caracteristic("FHOEHEV", "Desnivel", true, 0.0, true, "-.9");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FHOEHEV", Caract);
            OrderCaractComercial.Add(20, Caract);

            //Largo cabeza SUP incluso alarg
            Caract = new Caracteristic("FOT", "Largo cabeza SUP incluso alarg", true, 0.0, true, "-.11");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FOT", Caract);
            OrderCaractComercial.Add(30, Caract);

            //Largo cabeza INF incluso alarg
            Caract = new Caracteristic("FUT", "Largo cabeza INF incluso alarg", true, 0.0, true, "-.12");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FUT", Caract);
            OrderCaractComercial.Add(40, Caract);

            //Longitud desarrollo (metros)
            Caract = new Caracteristic("TNCR_OT_DESARROLLO", "Longitud desarrollo (metros)", true, 0.0, true, "-.14");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_OT_DESARROLLO", Caract);
            OrderCaractComercial.Add(50, Caract);
            #endregion

            #region 0 - DESIGN
            //Ancho nominal peldaño/paleta
            SortedList FBREITE = new SortedList();
            FBREITE.Add("3", new String[] { "3EK=0,6m", "3EK" });
            FBREITE.Add("4", new String[] { "4EK=0,8m", "4EK" });
            FBREITE.Add("5", new String[] { "5EK=1m", "5EK" });
            FBREITE.Add("6", new String[] { "6EK=1,2m", "6EK" });
            FBREITE.Add("7", new String[] { "7EK=1,4m", "7EK" });
            FBREITE.Add("8", new String[] { "8EK=1,6m", "8EK" });
            FBREITE.Add("5,5", new String[] { "5,5EK" });
            Caract = new Caracteristic("FBREITE", "Ancho nominal peldaño/paleta", false, FBREITE, true, "0.1");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FBREITE", Caract);
            OrderCaractComercial.Add(100, Caract);

            // Angulo de inclinacion
            Caract = new Caracteristic("FNEIGUNG", "Angulo de inclinacion", true, 0.0, true, "0.5");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FNEIGUNG", Caract);
            OrderCaractComercial.Add(110, Caract);
            #endregion

            #region 1 - GENERAL
            //Velocidad de funcionamiento
            Caract = new Caracteristic("FGESCHW", "Velocidad de funcionamiento", true, 0.0, true, "1.1");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FGESCHW", Caract);
            OrderCaractComercial.Add(200, Caract);

            // Normativa
            SortedList FNORM = new SortedList();
            FNORM.Add("EN", new String[] { "EN 115", "EN115" });
            FNORM.Add("ASME", new String[] { "ASME", "CSA/ASME" });
            FNORM.Add("AS", new String[] { "AS" });
            FNORM.Add("BS", new String[] { "BS" });
            FNORM.Add("CSA", new String[] { "CSA" });
            FNORM.Add("COP", new String[] { "COP" });
            Caract = new Caracteristic("FNORM", "Normativa", false, FNORM, true, "1.3");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FNORM", Caract);
            OrderCaractComercial.Add(210, Caract);

            // Condición Climática
            SortedList FKLIMAKLS = new SortedList();
            FKLIMAKLS.Add("1", new String[] { "Interior climatizado (Clase I)", "Interior climatizado (Clase I)" });
            FKLIMAKLS.Add("2", new String[] { "Exterior cubierto (Clase II)", "Exterior cubierto (Clase II)" });
            FKLIMAKLS.Add("3", new String[] { "Intemperie moderado(Clase III)", "Intemperie moderado(Clase III)" });
            FKLIMAKLS.Add("4", new String[] { "Intemperie (Clase IV)", "Intemperie (Clase IV)" });
            FKLIMAKLS.Add("5", new String[] { "Tropical (Clase V)", "Tropical (Clase V)" });
            Caract = new Caracteristic("FKLIMAKLS", "Condición Climática", false, FKLIMAKLS, true, "1.5");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FKLIMAKLS", Caract);
            OrderCaractComercial.Add(211, Caract);

            //Stop para carritos
            SortedList TNCR_POSTE_STOP_CARRITOS = new SortedList();
            TNCR_POSTE_STOP_CARRITOS.Add("KEINE", new String[] { "No lleva" });
            TNCR_POSTE_STOP_CARRITOS.Add("POSTE", new String[] { "Poste + STOP (MD+MI)" });
            TNCR_POSTE_STOP_CARRITOS.Add("POSMI", new String[] { "Poste + STOP (MI+MI)" });
            TNCR_POSTE_STOP_CARRITOS.Add("STOP", new String[] { "STOP", "En la pared del edificio" });
            TNCR_POSTE_STOP_CARRITOS.Add("ST_SP", new String[] { "En poste perfil exterior", "En poste exterior" });
            TNCR_POSTE_STOP_CARRITOS.Add("CRMD", new String[] { "En cristal (Lateral MD)", "En cristal" });
            TNCR_POSTE_STOP_CARRITOS.Add("CRMI", new String[] { "En cristal (Lateral MI)" });
            TNCR_POSTE_STOP_CARRITOS.Add("ESP", new String[] { "Especial" });
            Caract = new Caracteristic("TNCR_POSTE_STOP_CARRITOS", "Stop para carritos", false, TNCR_POSTE_STOP_CARRITOS, true, "1.6.1");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_POSTE_STOP_CARRITOS", Caract);
            OrderCaractComercial.Add(220, Caract);
            #endregion

            #region 2 - VOLTAGE AND FREQUENCY
            //Tensión de red
            SortedList FSPANNUNG = new SortedList();
            FSPANNUNG.Add("120/208", new String[] { "120/208V  3*208V+PE" });
            FSPANNUNG.Add("127/220", new String[] { "127/220V  3*220V+PE" });
            FSPANNUNG.Add("133/230", new String[] { "133/230V  3*230V+PE" });
            FSPANNUNG.Add("220/380", new String[] { "220/380V  3*380V+PE (+N)" });
            FSPANNUNG.Add("230/400", new String[] { "230/400V  3*400V+PE (+N)", "400V" });
            FSPANNUNG.Add("240/415", new String[] { "240/415V  3*415V+PE (+N)" });
            FSPANNUNG.Add("254/440", new String[] { "254/440V  3*440V+PE(+N)" });
            FSPANNUNG.Add("266/460", new String[] { "266/460V  3*460V+PE" });
            FSPANNUNG.Add("278/480", new String[] { "278/480V  3*480V+PE" });
            FSPANNUNG.Add("347/600", new String[] { "347/600V  3*600V+PE" });
            Caract = new Caracteristic("FSPANNUNG", "Tension de red", false, FSPANNUNG, true, "2.1");
            CaractComercial.Add("FSPANNUNG", Caract);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            OrderCaractComercial.Add(300, Caract);

            //Voltaje
            Caract = new Caracteristic("TNCR_S_TENSION_N", "Voltaje", true, 0.0, false, "2.1");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_S_TENSION_N", Caract);
            OrderCaractComercial.Add(310, Caract);

            //Entrada de Cables
            SortedList TNCR_OT_CABLES_ACOMETIDA = new SortedList();
            TNCR_OT_CABLES_ACOMETIDA.Add("SUP", new String[] { "Por cabeza superior", "Por cabeza superior" });
            TNCR_OT_CABLES_ACOMETIDA.Add("INF", new String[] { "Por cabeza inferior", "Por cabeza inferior" });
            TNCR_OT_CABLES_ACOMETIDA.Add("CEN", new String[] { "Por Zona Central (Especial)" });
            Caract = new Caracteristic("TNCR_OT_CABLES_ACOMETIDA", "Entrada de Cables", false, TNCR_OT_CABLES_ACOMETIDA, true, "2.3");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_OT_CABLES_ACOMETIDA", Caract);
            OrderCaractComercial.Add(320, Caract);
            #endregion

            #region 5 - LIGHTING
            // Iluminacion zócalos
            SortedList FSOCKELBEL = new SortedList();
            FSOCKELBEL.Add("BELOHNE", new String[] { "No", "No" });
            FSOCKELBEL.Add("LED", new String[] { "Estándar LED Blanca", "Estándar LED Blanca" });
            FSOCKELBEL.Add("PREMIUM", new String[] { "Premium LED Blanca", "Premium LED Blanca" });
            FSOCKELBEL.Add("RGB", new String[] { "RGB LED", "RGB LED" });
            FSOCKELBEL.Add("BELWEIS", new String[] { "Fluorescente blanco" });
            FSOCKELBEL.Add("BELKKLWEIS", new String[] { "Neon blanco" });
            FSOCKELBEL.Add("DIRECTA", new String[] { "LED Directa" });
            FSOCKELBEL.Add("4000K", new String[] { "4000K" });
            FSOCKELBEL.Add("SP", new String[] { "Especial" });
            Caract = new Caracteristic("FSOCKELBEL", "Iluminación de zócalos", false, FSOCKELBEL, true, "5.1");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FSOCKELBEL", Caract);
            OrderCaractComercial.Add(400, Caract);

            //Iluminación bajopasamanos
            SortedList FBALUBEL = new SortedList();
            FBALUBEL.Add("BELOHNE", new String[] { "No", "No" });
            FBALUBEL.Add("UHL", new String[] { "UHL LED Blanca", "UHL LED Blanca" });
            FBALUBEL.Add("UHLRGB", new String[] { "UHL LED RGB", "UHL LED RGB" });
            FBALUBEL.Add("LED", new String[] { "Estándar LED Blanca", "Estándar LED Blanca" });
            FBALUBEL.Add("PREMIUM", new String[] { "Premium LED Blanca", "Premium LED Blanca" });
            FBALUBEL.Add("RGB", new String[] { "RGB LED", "LED RGB" });
            FBALUBEL.Add("BELWEIS", new String[] { "Fluorescente blanco" });
            FBALUBEL.Add("BELKKLWEIS", new String[] { "Neon blanco" });
            FBALUBEL.Add("DIRECTA", new String[] { "LED Directa" });
            FBALUBEL.Add("4000K", new String[] { "4000K" });
            FBALUBEL.Add("SP", new String[] { "Especial" });
            Caract = new Caracteristic("FBALUBEL", "Iluminación bajopasamanos", false, FBALUBEL, true, "5.2");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FBALUBEL", Caract);
            OrderCaractComercial.Add(410, Caract);

            //Iluminación de peines
            SortedList ILUPEI = new SortedList();
            ILUPEI.Add("NO", new String[] { "No", "No" });
            ILUPEI.Add("DI", new String[] { "Led", "Fija" });
            ILUPEI.Add("PREMIUM_Y", new String[] { "Premium_amarilla" });
            ILUPEI.Add("PREMIUM_W", new String[] { "Premium_blanca" });
            ILUPEI.Add("PREMIUM_G", new String[] { "Premium_verde" });
            ILUPEI.Add("4000K", new String[] { "Led_4000K" });
            Caract = new Caracteristic("ILUPEI", "Iluminación de peines", false, ILUPEI, true, "5.3");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("ILUPEI", Caract);
            OrderCaractComercial.Add(420, Caract);

            //Iluminación de peines
            SortedList F35ZUB1 = new SortedList();
            F35ZUB1.Add("KEINE", new String[] { "No" });
            F35ZUB1.Add("KAMMBELLD", new String[] { "Fija" });
            F35ZUB1.Add("KAMMBELFLASH", new String[] { "Intermitente" });
            F35ZUB1.Add("KAMMBELGL", new String[] { "Incandescente" });
            F35ZUB1.Add("KAMMBELLR", new String[] { "Fluorescente" });
            F35ZUB1.Add("KAMMBELNORKA", new String[] { "Norka" });
            F35ZUB1.Add("LAUFLICHT", new String[] { "LEDs dinámico" });
            Caract = new Caracteristic("F35ZUB1", "Iluminacion de peines", false, F35ZUB1, false);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("F35ZUB1", Caract);
            OrderCaractComercial.Add(430, Caract);

            //Iluminación estroboscópica
            SortedList FSHEITSBEL = new SortedList();
            FSHEITSBEL.Add("KEINE", new String[] { "No", "No" });
            FSHEITSBEL.Add("STFSPALTBEL", new String[] { "1 lampara", "1 lámpara" });
            FSHEITSBEL.Add("STFSPALTBEL2", new String[] { "2 lamparas", "2 lámparas" });
            FSHEITSBEL.Add("LED", new String[] { "LED stripe", "Tira LED" });
            FSHEITSBEL.Add("2LED", new String[] { "2 Tiras LED" });
            FSHEITSBEL.Add("3LED", new String[] { "3 Tiras LED" });
            Caract = new Caracteristic("FSHEITSBEL", "Iluminación estroboscópica", false, FSHEITSBEL, true, "5.4");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FSHEITSBEL", Caract);
            OrderCaractComercial.Add(440, Caract);

            //Iluminacion en fosos
            SortedList FBELANTRST = new SortedList();
            FBELANTRST.Add("KEINE", new String[] { "Ninguna", "No", "Lámpara portátil" });
            FBELANTRST.Add("NORKA", new String[] { "Norka" });
            FBELANTRST.Add("OVAL", new String[] { "Oval", "Fija (Oval)" });
            FBELANTRST.Add("HANDL", new String[] { "Manual" });
            FBELANTRST.Add("AUTO", new String[] { "Automática" });
            Caract = new Caracteristic("FBELANTRST", "Iluminacion en fosos", false, FBELANTRST, true, "5.6");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FBELANTRST", Caract);
            OrderCaractComercial.Add(450, Caract);
            #endregion

            #region 8 - CONTROL
            //MAX
            SortedList TNCR_S_MAX = new SortedList();
            TNCR_S_MAX.Add("N", new String[] { "No disponible", "No disponible" });
            TNCR_S_MAX.Add("R", new String[] { "Preparado", "Preparado" });
            TNCR_S_MAX.Add("A", new String[] { "Instalado", "Instalado" });
            Caract = new Caracteristic("TNCR_S_MAX", "MAX", false, TNCR_S_MAX, true, "8.2");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_S_MAX", Caract);
            OrderCaractComercial.Add(499, Caract);

            //Ubicación del Controlador
            SortedList F53ZUB7 = new SortedList();
            F53ZUB7.Add("INNENOBEN", new String[] { "Armario interior en cabeza sup", "Armario interior en cabeza superior" });
            F53ZUB7.Add("INNENUNTEN", new String[] { "Dentro del foso inferior" });
            F53ZUB7.Add("AUSSEN", new String[] { "Armario exterior_Metalico", "Armario exterior - Metálico" });
            F53ZUB7.Add("ARM_ESP", new String[] { "Armario exterior_Acero Inox", "Armario exterior - Acero inoxidable 304" });
            Caract = new Caracteristic("F53ZUB7", "Ubicación del Controlador", false, F53ZUB7, true, "8.3");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("F53ZUB7", Caract);
            OrderCaractComercial.Add(500, Caract);

            //Distancia en metros al armario
            Caract = new Caracteristic("TNCR_OT_DISTANCIA_ARMARIO", "Distancia en metros al armario", true, 0.0, true, "8.3.1");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_OT_DISTANCIA_ARMARIO", Caract);
            OrderCaractComercial.Add(510, Caract);

            //Llavín local remoto
            SortedList LLAVES_LOCAL_REM = new SortedList();
            LLAVES_LOCAL_REM.Add("N", new String[] { "No lleva llaves L/R", "No" });
            LLAVES_LOCAL_REM.Add("B", new String[] { "Lleva 1 en balaustrada" });
            LLAVES_LOCAL_REM.Add("P", new String[] { "Lleva 1 en poste" });
            LLAVES_LOCAL_REM.Add("E", new String[] { "Lleva 1 en entrada pasamanos" });
            LLAVES_LOCAL_REM.Add("S", new String[] { "Existe" });
            LLAVES_LOCAL_REM.Add("A", new String[] { "Lleva 1 en armario" });
            Caract = new Caracteristic("LLAVES_LOCAL_REM", "Llavín local remoto", false, LLAVES_LOCAL_REM, true, "8.4");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("LLAVES_LOCAL_REM", Caract);
            OrderCaractComercial.Add(520, Caract);
            //Llavín paro por impulso
            SortedList LLAVES_PARO = new SortedList();
            LLAVES_PARO.Add("N", new String[] { "No lleva llaves de paro", "No" });
            LLAVES_PARO.Add("B", new String[] { "Lleva 1 en balaustrada" });
            LLAVES_PARO.Add("P", new String[] { "Lleva 1 en poste" });
            LLAVES_PARO.Add("A", new String[] { "Lleva 1 en armario" });
            LLAVES_PARO.Add("C", new String[] { "En cubrezocalo" });
            Caract = new Caracteristic("LLAVES_PARO", "Llavín paro por impulso", false, LLAVES_PARO, true, "8.5");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("LLAVES_PARO", Caract);
            OrderCaractComercial.Add(530, Caract);

            //Llavín automatico cotinuo
            SortedList LLAVES_AUT_CONT = new SortedList();
            LLAVES_AUT_CONT.Add("N", new String[] { "No lleva llaves A/U" });
            LLAVES_AUT_CONT.Add("B", new String[] { "Lleva 1 en balaustrada" });
            LLAVES_AUT_CONT.Add("P", new String[] { "Lleva 1 en poste" });
            LLAVES_AUT_CONT.Add("E", new String[] { "Lleva 1 en entrada pasamanos" });
            LLAVES_AUT_CONT.Add("A", new String[] { "Lleva 1 en armario" });
            Caract = new Caracteristic("LLAVES_AUT_CONT", "Llavín automatico cotinuo", false, LLAVES_AUT_CONT, true, "CUSTOM");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("LLAVES_AUT_CONT", Caract);
            OrderCaractComercial.Add(540, Caract);

            #endregion

            #region 9 - SAFETY DEVICES
            //Seguridad rotura de pasamanos
            SortedList F09ZUB1 = new SortedList();
            F09ZUB1.Add("BRUCHSCHALT", new String[] { "Si", "Si" });
            F09ZUB1.Add("KEINE", new String[] { "No", "No" });
            Caract = new Caracteristic("F09ZUB1", "Seguridad rotura de pasamanos", false, F09ZUB1, true, "9.3");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("F09ZUB1", Caract);
            OrderCaractComercial.Add(600, Caract);

            //Control desgaste de frenos
            SortedList F01ZUB = new SortedList();
            F01ZUB.Add("BVERSCHLF", new String[] { "Antena" });
            F01ZUB.Add("BVERSCHLINI", new String[] { "Iniciador" });
            F01ZUB.Add("KEINE", new String[] { "No", "No" });
            F01ZUB.Add("INDUCTIVO", new String[] { "Si", "Si" });
            Caract = new Caracteristic("F01ZUB", "Control desgaste de frenos", false, F01ZUB, true, "9.5");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("F01ZUB", Caract);
            OrderCaractComercial.Add(610, Caract);

            //Doble freno
            SortedList FBREMSE2 = new SortedList();
            FBREMSE2.Add("2/2", new String[] { "No", "No (H ≤ 6m)" });
            FBREMSE2.Add("4/4", new String[] { "Si", "Si (H > 6m)", "Si (H ≤ 6 m)" });
            FBREMSE2.Add("N", new String[] { "No calculado" });
            Caract = new Caracteristic("FBREMSE2", "Doble freno", false, FBREMSE2, true, "9.6");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FBREMSE2", Caract);
            OrderCaractComercial.Add(620, Caract);

            //Seguridad de peines
            SortedList FKAMMPLHK = new SortedList();
            FKAMMPLHK.Add("INDEPENDIENTE", new String[] { "Vertical independient de horiz", "Vertical independiente de la horizontal" });
            FKAMMPLHK.Add("HUBKONTAKT", new String[] { "Vertical combinada con horiz", "Vertical combinada con horizontal" });
            FKAMMPLHK.Add("KEINE", new String[] { "Horizontal", "Horizontal" });
            Caract = new Caracteristic("FKAMMPLHK", "Seguridad de peines", false, FKAMMPLHK, true, "9.14");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FKAMMPLHK", Caract);
            OrderCaractComercial.Add(630, Caract);

            //Seguridad Buggy
            SortedList F04ZUB = new SortedList();
            F04ZUB.Add("BUGGY", new String[] { "Ambas cabezas", "Ambas cabezas" });
            F04ZUB.Add("BUGGYOT", new String[] { "Cabeza superior", "Cabeza superior" });
            F04ZUB.Add("BUGGYUT", new String[] { "Cabeza inferior", "Cabeza inferior" });
            F04ZUB.Add("KEINE", new String[] { "No", "No" });
            Caract = new Caracteristic("F04ZUB", "Seguridad Buggy", false, F04ZUB, true, "9.20");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("F04ZUB", Caract);
            OrderCaractComercial.Add(640, Caract);

            //Seguridad cadena principal
            SortedList TNCR_S_DRIVE_CHAIN = new SortedList();
            TNCR_S_DRIVE_CHAIN.Add("SI", new String[] { "Si", "Si" });
            TNCR_S_DRIVE_CHAIN.Add("NO", new String[] { "No", "No" });
            Caract = new Caracteristic("TNCR_S_DRIVE_CHAIN", "Seguridad cadena principal", false, TNCR_S_DRIVE_CHAIN, true, "9.21");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_S_DRIVE_CHAIN", Caract);
            OrderCaractComercial.Add(650, Caract);

            //Número de microcontactos
            Caract = new Caracteristic("TNCR_OT_NUM_MICROCONT", "Número de microcontactos", true, 0.0, true, "9.22.1");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_OT_NUM_MICROCONT", Caract);
            OrderCaractComercial.Add(660, Caract);

            //Stop adicional
            SortedList TNCR_OT_E_STOP_ADICIONAL = new SortedList();
            TNCR_OT_E_STOP_ADICIONAL.Add("N", new String[] { "No", "No" });
            TNCR_OT_E_STOP_ADICIONAL.Add("S", new String[] { "Uno en cabeza superior" });
            TNCR_OT_E_STOP_ADICIONAL.Add("2", new String[] { "Uno en cada cabeza", "Sólo stop", "Stop en balaustrada" });
            Caract = new Caracteristic("TNCR_OT_E_STOP_ADICIONAL", "Stop adicional", false, TNCR_OT_E_STOP_ADICIONAL, true, "9.25");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_OT_E_STOP_ADICIONAL", Caract);
            OrderCaractComercial.Add(670, Caract);

            //Semáforos
            SortedList FAMPELSYM = new SortedList();
            FAMPELSYM.Add("F6N", new String[] { "Azul con flecha blanca" });
            FAMPELSYM.Add("W1N", new String[] { "Direccion de trafico(reversib)" });
            FAMPELSYM.Add("ROTGRUEN", new String[] { "Rojo / Verde (2)" });
            FAMPELSYM.Add("ROTGRUENW1N", new String[] { "Rojo/Verde/Direccion trafic(3)" });
            FAMPELSYM.Add("BICOLOR", new String[] { "Bicolor", "Bicolor" });
            FAMPELSYM.Add("F6NEINB", new String[] { "Flecha/Direccion unica (2)" });
            FAMPELSYM.Add("F6NEINBW1N", new String[] { "Flecha/Direcc unic/Direcc traf" });
            FAMPELSYM.Add("DINAMICO", new String[] { "Dinámico", "Dinámico" });
            FAMPELSYM.Add("NINGUNO", new String[] { "No" });
            FAMPELSYM.Add("PROHI_VERDE", new String[] { "Prohibido / Verde (2)" });
            Caract = new Caracteristic("FAMPELSYM", "Semáforos", false, FAMPELSYM, true, "9.26.1");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FAMPELSYM", Caract);
            OrderCaractComercial.Add(680, Caract);

            //Display
            SortedList TNDIAGNOSTICO = new SortedList();
            TNDIAGNOSTICO.Add("SI", new String[] { "Cubrezócalo" });
            TNDIAGNOSTICO.Add("NO", new String[] { "No" });
            TNDIAGNOSTICO.Add("AR", new String[] { "En armario exterior" });
            TNDIAGNOSTICO.Add("MI", new String[] { "Mini-display en cubrezocalos" });
            TNDIAGNOSTICO.Add("CI", new String[] { "Cassette interior" });
            TNDIAGNOSTICO.Add("BA", new String[] { "En balaustrada" });
            TNDIAGNOSTICO.Add("ES", new String[] { "Especial" });
            TNDIAGNOSTICO.Add("AR_BA", new String[] { "En armario y balustrada" });
            TNDIAGNOSTICO.Add("AR_ZO", new String[] { "En armario y cubrezócalo" });
            TNDIAGNOSTICO.Add("A2", new String[] { "En armario interior" });
            Caract = new Caracteristic("TNDIAGNOSTICO", "Diagnostico", false, TNDIAGNOSTICO, true, "9.27");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNDIAGNOSTICO", Caract);
            OrderCaractComercial.Add(681, Caract);

            // Sistema anden
            SortedList FWIEDERB = new SortedList();
            FWIEDERB.Add("KEINE", new String[] { "No", "No" });
            FWIEDERB.Add("WBEREITSCH", new String[] { "Si", "Si" });
            Caract = new Caracteristic("FWIEDERB", "Sistema andén", false, FWIEDERB, true, "9.28");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FWIEDERB", Caract);
            OrderCaractComercial.Add(690, Caract);

            // Detector nivel de agua en foso
            SortedList TNCR_OT_NIVEL_AGUA = new SortedList();
            TNCR_OT_NIVEL_AGUA.Add("N", new String[] { "No", "No" });
            TNCR_OT_NIVEL_AGUA.Add("S", new String[] { "Si", "Si" });
            Caract = new Caracteristic("TNCR_OT_NIVEL_AGUA", "Detector nivel de agua en foso", false, TNCR_OT_NIVEL_AGUA, true, "9.30");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_OT_NIVEL_AGUA", Caract);
            OrderCaractComercial.Add(700, Caract);

            //Cerrojo mantenimiento eje ppal
            SortedList TNCR_OT_CERROJO_MANTENIMIENTO = new SortedList();
            TNCR_OT_CERROJO_MANTENIMIENTO.Add("S", new String[] { "Si", "Si" });
            TNCR_OT_CERROJO_MANTENIMIENTO.Add("N", new String[] { "No", "No" });
            TNCR_OT_CERROJO_MANTENIMIENTO.Add("P", new String[] { "En banda de peldaños" });
            Caract = new Caracteristic("TNCR_OT_CERROJO_MANTENIMIENTO", "Cerrojo mantenimiento eje ppal", false, TNCR_OT_CERROJO_MANTENIMIENTO, true, "9.31");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_OT_CERROJO_MANTENIMIENTO", Caract);
            OrderCaractComercial.Add(710, Caract);

            // Sensor aceite en reductor
            SortedList TNCR_SENSOR_ACEITE_REDUC = new SortedList();
            TNCR_SENSOR_ACEITE_REDUC.Add("N", new String[] { "No", "No" });
            TNCR_SENSOR_ACEITE_REDUC.Add("S", new String[] { "Si", "Si" });
            Caract = new Caracteristic("TNCR_SENSOR_ACEITE_REDUC", "Sensor aceite en reductor", false, TNCR_SENSOR_ACEITE_REDUC, true, "9.34");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_SENSOR_ACEITE_REDUC", Caract);
            OrderCaractComercial.Add(720, Caract);

            //Contacto de fuego
            SortedList TNCR_CONTACTO_FUEGO = new SortedList();
            TNCR_CONTACTO_FUEGO.Add("S", new String[] { "Si", "Si" });
            TNCR_CONTACTO_FUEGO.Add("N", new String[] { "No", "No" });
            Caract = new Caracteristic("TNCR_CONTACTO_FUEGO", "Contacto de fuego", false, TNCR_CONTACTO_FUEGO, true, "9.35");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_CONTACTO_FUEGO", Caract);
            OrderCaractComercial.Add(730, Caract);
            #endregion

            #region 10 - DRIVE SYSTEM
            //Modo de funcionamiento
            SortedList FBETRART = new SortedList();
            FBETRART.Add("DAUER", new String[] { "Contínuo", "Operación continua " });
            FBETRART.Add("INTERM", new String[] { "Automático" });
            FBETRART.Add("BV", new String[] { "Baja velocidad", "Operación a baja velocidad" });
            FBETRART.Add("SG", new String[] { "Stop&Go", "Operación Stop&Go" });
            FBETRART.Add("SGBV", new String[] { "Stop&Go y baja velocidad", "Operación Stop&Go y baja velocidad" });
            Caract = new Caracteristic("FBETRART", "Modo de funcionamiento", false, FBETRART, true, "10.2");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FBETRART", Caract);
            OrderCaractComercial.Add(800, Caract);


            //VVVF
            SortedList TNCR_SD_SIST_AHORRO = new SortedList();
            TNCR_SD_SIST_AHORRO.Add("NO", new String[] { "No" });
            TNCR_SD_SIST_AHORRO.Add("VA", new String[] { "Variador Full Load(Estándar)", "Estándar" });
            TNCR_SD_SIST_AHORRO.Add("EC", new String[] { "Con economizador estrella-tr" });
            TNCR_SD_SIST_AHORRO.Add("SINUMEC", new String[] { "SINUMEC (EEC)" });
            TNCR_SD_SIST_AHORRO.Add("VA_PART", new String[] { "Variador Partial Load" });
            TNCR_SD_SIST_AHORRO.Add("VA_REG", new String[] { "Regenerativo", "Regenerativo" });
            Caract = new Caracteristic("TNCR_SD_SIST_AHORRO", "VVVF", false, TNCR_SD_SIST_AHORRO, true, "10.2.1");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_SD_SIST_AHORRO", Caract);
            OrderCaractComercial.Add(810, Caract);

            //Tipo detección
            SortedList FLICHTINT = new SortedList();
            FLICHTINT.Add("KEINE", new String[] { "Ninguno" });
            FLICHTINT.Add("LICHTINT", new String[] { "Fotocelula", "Fotocélula" });
            FLICHTINT.Add("RADAR", new String[] { "Radar", "Radar " });
            Caract = new Caracteristic("FLICHTINT", "Tipo detección", false, FLICHTINT, true, "10.2.2");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FLICHTINT", Caract);
            OrderCaractComercial.Add(820, Caract);

            //By Pass
            SortedList TNCR_OT_BYPASS_VARIADOR = new SortedList();
            TNCR_OT_BYPASS_VARIADOR.Add("N", new String[] { "No", "No" });
            TNCR_OT_BYPASS_VARIADOR.Add("S", new String[] { "Emergencia" });
            TNCR_OT_BYPASS_VARIADOR.Add("E", new String[] { "Estrella-triangulo" });
            TNCR_OT_BYPASS_VARIADOR.Add("SI", new String[] { "Si (valorar a mano directo/ET)", "Sí" });
            Caract = new Caracteristic("TNCR_OT_BYPASS_VARIADOR", "By-Pass", false, TNCR_OT_BYPASS_VARIADOR, true, "10.2.3");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_OT_BYPASS_VARIADOR", Caract);
            OrderCaractComercial.Add(830, Caract);

            //Revoluciones del motor
            Caract = new Caracteristic("FACH", "Revoluciones del motor", true, 0.0, true, "10.3");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FACH", Caract);
            OrderCaractComercial.Add(840, Caract);

            //Lubricacion automatica
            SortedList TNCR_ENGRASE_AUTOMATICO = new SortedList();
            TNCR_ENGRASE_AUTOMATICO.Add("N", new String[] { "No", "No" });
            TNCR_ENGRASE_AUTOMATICO.Add("S", new String[] { "Si", "Si (EN115): Sin prenivel}", "Si (ASME)" });
            TNCR_ENGRASE_AUTOMATICO.Add("C", new String[] { "Automático, capacidad 12L" });
            TNCR_ENGRASE_AUTOMATICO.Add("P", new String[] { "Bomba con prenivel", "Si (EN115): Con prenivel" });
            Caract = new Caracteristic("TNCR_ENGRASE_AUTOMATICO", "Lubricacion automatica", false, TNCR_ENGRASE_AUTOMATICO, true, "10.11");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("TNCR_ENGRASE_AUTOMATICO", Caract);
            OrderCaractComercial.Add(850, Caract);

            // Freno auxiliar en eje principal
            SortedList FZUSBREMSE = new SortedList();
            FZUSBREMSE.Add("HWBUBENCER", new String[] { "Bubencer en eje principal" });
            FZUSBREMSE.Add("HWSPERRKMAGN", new String[] { "Trinquete magnét. en eje ppal" });
            FZUSBREMSE.Add("HWSPERRKMECH", new String[] { "Trinquete mecánico en eje ppal", "Trinquete mecánico en eje principal" });
            FZUSBREMSE.Add("KEINE", new String[] { "No", "No" });
            FZUSBREMSE.Add("MOTORW", new String[] { "Eje motor" });
            FZUSBREMSE.Add("FRENO_VA", new String[] { "Freno variador" });
            FZUSBREMSE.Add("NAB", new String[] { "NAB", "NAB" });
            Caract = new Caracteristic("FZUSBREMSE", "Freno auxiliar en eje principal", false, FZUSBREMSE, true, "10.13");
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FZUSBREMSE", Caract);
            OrderCaractComercial.Add(860, Caract);

            #endregion

            #region 11 - TRANSPORTATION
            //Número particiones centrales
            Caract = new Caracteristic("FEFL0", "Número particiones centrales", true, 0.0, true, "11.1");
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("FEFL0", Caract);
            OrderCaractComercial.Add(870, Caract);
            #endregion

            #region OTHER
            //País de destino
            SortedList FLAND = new SortedList();
            FLAND.Add("1", new String[] { "Francia" });
            FLAND.Add("2", new String[] { "Mónaco" });
            FLAND.Add("3", new String[] { "Países Bajos" });
            FLAND.Add("4", new String[] { "Alemania" });
            FLAND.Add("5", new String[] { "Italia" });
            FLAND.Add("6", new String[] { "Reino Unido" });
            FLAND.Add("7", new String[] { "Irlanda" });
            FLAND.Add("8", new String[] { "Dinamarca" });
            FLAND.Add("9", new String[] { "Grecia" });
            FLAND.Add("10", new String[] { "Portugal" });
            FLAND.Add("11", new String[] { "España" });
            FLAND.Add("17", new String[] { "Bélgica" });
            FLAND.Add("18", new String[] { "Luxemburgo" });
            FLAND.Add("21", new String[] { "Ceuta" });
            FLAND.Add("23", new String[] { "Melilla" });
            FLAND.Add("24", new String[] { "Islandia" });
            FLAND.Add("28", new String[] { "Noruega" });
            FLAND.Add("30", new String[] { "Suecia" });
            FLAND.Add("32", new String[] { "Finlandia" });
            FLAND.Add("37", new String[] { "Liechtenstein" });
            FLAND.Add("38", new String[] { "Austria" });
            FLAND.Add("39", new String[] { "Suiza" });
            FLAND.Add("41", new String[] { "Islas Feroe" });
            FLAND.Add("43", new String[] { "Andorra" });
            FLAND.Add("44", new String[] { "Gibraltar" });
            FLAND.Add("45", new String[] { "El Vaticano" });
            FLAND.Add("46", new String[] { "Malta" });
            FLAND.Add("47", new String[] { "San Marino" });
            FLAND.Add("52", new String[] { "Turquia" });
            FLAND.Add("53", new String[] { "Estonia" });
            FLAND.Add("54", new String[] { "Letonia" });
            FLAND.Add("55", new String[] { "Lituania" });
            FLAND.Add("60", new String[] { "Polonia" });
            FLAND.Add("61", new String[] { "República Checa" });
            FLAND.Add("63", new String[] { "Eslovaquia" });
            FLAND.Add("64", new String[] { "Hungría" });
            FLAND.Add("66", new String[] { "Rumania" });
            FLAND.Add("68", new String[] { "Bulgaria" });
            FLAND.Add("70", new String[] { "Albania" });
            FLAND.Add("72", new String[] { "Ucrania" });
            FLAND.Add("73", new String[] { "Bielorusina" });
            FLAND.Add("74", new String[] { "Moldavia" });
            FLAND.Add("75", new String[] { "Rusia" });
            FLAND.Add("76", new String[] { "Georgia" });
            FLAND.Add("77", new String[] { "Armenia" });
            FLAND.Add("78", new String[] { "Azerbaiyán" });
            FLAND.Add("79", new String[] { "Kazajstán" });
            FLAND.Add("80", new String[] { "Turkmenistán" });
            FLAND.Add("81", new String[] { "Uzbekistan" });
            FLAND.Add("82", new String[] { "Tayikistán" });
            FLAND.Add("83", new String[] { "Kirgisistan" });
            FLAND.Add("91", new String[] { "Eslovenia" });
            FLAND.Add("92", new String[] { "Croacia" });
            FLAND.Add("93", new String[] { "Bosnia" });
            FLAND.Add("94", new String[] { "Yugoslavia" });
            FLAND.Add("96", new String[] { "Ehemalige jugoslawische Republ" });
            FLAND.Add("204", new String[] { "Marruecos" });
            FLAND.Add("208", new String[] { "Argelia" });
            FLAND.Add("212", new String[] { "Tunez" });
            FLAND.Add("216", new String[] { "Libia" });
            FLAND.Add("220", new String[] { "Egipto" });
            FLAND.Add("224", new String[] { "Sudan" });
            FLAND.Add("228", new String[] { "Mauretanien" });
            FLAND.Add("232", new String[] { "Mali" });
            FLAND.Add("236", new String[] { "Burkina Faso" });
            FLAND.Add("240", new String[] { "Niger" });
            FLAND.Add("244", new String[] { "Tschad" });
            FLAND.Add("247", new String[] { "Kap Verde" });
            FLAND.Add("248", new String[] { "Senegal" });
            FLAND.Add("252", new String[] { "Gambia" });
            FLAND.Add("257", new String[] { "Guinea-Bissau" });
            FLAND.Add("260", new String[] { "Guinea" });
            FLAND.Add("264", new String[] { "Sierra Leone" });
            FLAND.Add("268", new String[] { "Liberia" });
            FLAND.Add("272", new String[] { "Côte d'Ivoire" });
            FLAND.Add("276", new String[] { "Ghana" });
            FLAND.Add("280", new String[] { "Togo" });
            FLAND.Add("284", new String[] { "Benin" });
            FLAND.Add("288", new String[] { "Nigeria" });
            FLAND.Add("302", new String[] { "Kamerun" });
            FLAND.Add("306", new String[] { "Zentralafrikanische Republik" });
            FLAND.Add("310", new String[] { "Äquatorialguinea" });
            FLAND.Add("311", new String[] { "Sao Tome und Principe" });
            FLAND.Add("314", new String[] { "Gabun" });
            FLAND.Add("318", new String[] { "Republik Kongo" });
            FLAND.Add("322", new String[] { "Demokratische Republik Kongo" });
            FLAND.Add("324", new String[] { "Ruanda" });
            FLAND.Add("328", new String[] { "Burundi" });
            FLAND.Add("329", new String[] { "St. Helena" });
            FLAND.Add("330", new String[] { "Angola" });
            FLAND.Add("334", new String[] { "Etiopía" });
            FLAND.Add("336", new String[] { "Eritrea" });
            FLAND.Add("338", new String[] { "Dschibuti" });
            FLAND.Add("342", new String[] { "Somalia" });
            FLAND.Add("346", new String[] { "Kenia" });
            FLAND.Add("350", new String[] { "Uganda" });
            FLAND.Add("352", new String[] { "Tansania" });
            FLAND.Add("355", new String[] { "Seychellen u. zugehör. Geb." });
            FLAND.Add("357", new String[] { "Brit. Gebiet im Ind. Ozean" });
            FLAND.Add("366", new String[] { "Mosambik" });
            FLAND.Add("370", new String[] { "Madagaskar" });
            FLAND.Add("373", new String[] { "Mauritius" });
            FLAND.Add("375", new String[] { "Komoren" });
            FLAND.Add("377", new String[] { "Mayotte" });
            FLAND.Add("378", new String[] { "Sambia" });
            FLAND.Add("382", new String[] { "Simbabwe" });
            FLAND.Add("386", new String[] { "Malawi" });
            FLAND.Add("388", new String[] { "Südafrika" });
            FLAND.Add("389", new String[] { "Namibia" });
            FLAND.Add("391", new String[] { "Botsuana" });
            FLAND.Add("393", new String[] { "Swasiland" });
            FLAND.Add("395", new String[] { "Lesotho" });
            FLAND.Add("400", new String[] { "Estados Unidos" });
            FLAND.Add("404", new String[] { "Canadá" });
            FLAND.Add("406", new String[] { "Grönland" });
            FLAND.Add("408", new String[] { "St. Pierre und Miquelon" });
            FLAND.Add("412", new String[] { "Mexiko" });
            FLAND.Add("413", new String[] { "Bermuda" });
            FLAND.Add("416", new String[] { "Guatemala" });
            FLAND.Add("421", new String[] { "Belize" });
            FLAND.Add("424", new String[] { "Honduras" });
            FLAND.Add("428", new String[] { "El Salvador" });
            FLAND.Add("432", new String[] { "Nicaragua" });
            FLAND.Add("436", new String[] { "Costa Rica" });
            FLAND.Add("442", new String[] { "Panama" });
            FLAND.Add("446", new String[] { "Anguilla" });
            FLAND.Add("448", new String[] { "Cuba" });
            FLAND.Add("449", new String[] { "St. Kitts und Nevis" });
            FLAND.Add("452", new String[] { "Haiti" });
            FLAND.Add("453", new String[] { "Bahamas" });
            FLAND.Add("454", new String[] { "Turks- und Caicosinseln" });
            FLAND.Add("456", new String[] { "Dominikanische Republik" });
            FLAND.Add("457", new String[] { "Amerikan. Jungferninseln" });
            FLAND.Add("459", new String[] { "Antigua und Barbuda" });
            FLAND.Add("460", new String[] { "Dominica" });
            FLAND.Add("463", new String[] { "Kaimaninseln" });
            FLAND.Add("464", new String[] { "Jamaika" });
            FLAND.Add("465", new String[] { "St. Lucia" });
            FLAND.Add("467", new String[] { "St. Vincent" });
            FLAND.Add("468", new String[] { "Britische Jungferninseln" });
            FLAND.Add("469", new String[] { "Barbados" });
            FLAND.Add("470", new String[] { "Montserrat" });
            FLAND.Add("472", new String[] { "Trinidad und Tobago" });
            FLAND.Add("473", new String[] { "Grenada" });
            FLAND.Add("474", new String[] { "Aruba" });
            FLAND.Add("478", new String[] { "Niederländische Antillen" });
            FLAND.Add("480", new String[] { "Kolumbien" });
            FLAND.Add("484", new String[] { "Venezuela" });
            FLAND.Add("488", new String[] { "Guyana" });
            FLAND.Add("492", new String[] { "Suriname" });
            FLAND.Add("500", new String[] { "Ecuador" });
            FLAND.Add("504", new String[] { "Peru" });
            FLAND.Add("508", new String[] { "Brasilien" });
            FLAND.Add("512", new String[] { "Chile" });
            FLAND.Add("516", new String[] { "Bolivien" });
            FLAND.Add("520", new String[] { "Paraguay" });
            FLAND.Add("524", new String[] { "Uruguay" });
            FLAND.Add("528", new String[] { "Argentinien" });
            FLAND.Add("529", new String[] { "Falklandinseln" });
            FLAND.Add("600", new String[] { "Chipre" });
            FLAND.Add("604", new String[] { "Libanon" });
            FLAND.Add("608", new String[] { "Arabische Republik Syrien" });
            FLAND.Add("612", new String[] { "Irak" });
            FLAND.Add("616", new String[] { "Iran" });
            FLAND.Add("624", new String[] { "Israel" });
            FLAND.Add("625", new String[] { "Westjordanland/Gazastreifen" });
            FLAND.Add("628", new String[] { "Jordanien" });
            FLAND.Add("632", new String[] { "Saudi-Arabien" });
            FLAND.Add("636", new String[] { "Kuwait" });
            FLAND.Add("640", new String[] { "Bahrain" });
            FLAND.Add("644", new String[] { "Katar" });
            FLAND.Add("647", new String[] { "Vereinigte Arabische Emirate" });
            FLAND.Add("649", new String[] { "Oman" });
            FLAND.Add("653", new String[] { "Jemen" });
            FLAND.Add("660", new String[] { "Afghanistan" });
            FLAND.Add("662", new String[] { "Pakistan" });
            FLAND.Add("664", new String[] { "Indien" });
            FLAND.Add("666", new String[] { "Bangladesch" });
            FLAND.Add("667", new String[] { "Malediven" });
            FLAND.Add("669", new String[] { "Sri Lanka" });
            FLAND.Add("672", new String[] { "Nepal" });
            FLAND.Add("675", new String[] { "Bhutan" });
            FLAND.Add("676", new String[] { "Myanmar" });
            FLAND.Add("680", new String[] { "Thailand" });
            FLAND.Add("684", new String[] { "Demokratische Volksrepublik La" });
            FLAND.Add("690", new String[] { "Vietnam" });
            FLAND.Add("696", new String[] { "Kambodscha" });
            FLAND.Add("700", new String[] { "Indonesien" });
            FLAND.Add("701", new String[] { "Malaysia" });
            FLAND.Add("703", new String[] { "Brunei Darussalam" });
            FLAND.Add("706", new String[] { "Singapur" });
            FLAND.Add("708", new String[] { "Philippinen" });
            FLAND.Add("716", new String[] { "Mongolei" });
            FLAND.Add("720", new String[] { "China" });
            FLAND.Add("724", new String[] { "Demokratische Volksrepublik Ko" });
            FLAND.Add("728", new String[] { "Republik Korea" });
            FLAND.Add("732", new String[] { "Japan" });
            FLAND.Add("736", new String[] { "Taiwan" });
            FLAND.Add("740", new String[] { "Hongkong" });
            FLAND.Add("743", new String[] { "Macau" });
            FLAND.Add("800", new String[] { "Australien" });
            FLAND.Add("801", new String[] { "Papua-Neuguinea" });
            FLAND.Add("802", new String[] { "Australisch-Ozeanien" });
            FLAND.Add("803", new String[] { "Nauru" });
            FLAND.Add("804", new String[] { "Neuseeland" });
            FLAND.Add("806", new String[] { "Salomonen" });
            FLAND.Add("807", new String[] { "Tuvalu" });
            FLAND.Add("809", new String[] { "Neukaledonien" });
            FLAND.Add("810", new String[] { "Amerikanisch-Ozeanien" });
            FLAND.Add("811", new String[] { "Wallis und Futuna" });
            FLAND.Add("812", new String[] { "Kiribati" });
            FLAND.Add("813", new String[] { "Pitcairn" });
            FLAND.Add("814", new String[] { "Neuseeländisch-Ozeanien" });
            FLAND.Add("815", new String[] { "Fidschi" });
            FLAND.Add("816", new String[] { "Vanuatu" });
            FLAND.Add("817", new String[] { "Tonga" });
            FLAND.Add("819", new String[] { "Samoa" });
            FLAND.Add("820", new String[] { "Nördliche Marianen" });
            FLAND.Add("822", new String[] { "Französisch-Polynesien" });
            FLAND.Add("823", new String[] { "Föderierte Staaten von Mikrone" });
            FLAND.Add("824", new String[] { "Marshall-Inseln" });
            FLAND.Add("825", new String[] { "Palau" });
            FLAND.Add("890", new String[] { "Polargebiete" });
            FLAND.Add("950", new String[] { "Schiffs- und Luftfahrzeugbed." });
            FLAND.Add("958", new String[] { "Nicht ermittelte Länder" });
            Caract = new Caracteristic("FLAND", "País de destino", false, FLAND, true);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FLAND", Caract);
            OrderCaractComercial.Add(900, Caract);

            //Unidad de accionamiento
            SortedList FANTREHT = new SortedList();
            FANTREHT.Add("HISTA", new String[] { "Standar alto (820/840)" });
            FANTREHT.Add("SOG", new String[] { "3C" });
            FANTREHT.Add("EM", new String[] { "EMERSON 125AR" });
            FANTREHT.Add("FJ", new String[] { "FJ 125/160" });
            FANTREHT.Add("FTJ+FJ", new String[] { "Emod FTJ + Fuma FJ" });
            FANTREHT.Add("5C", new String[] { "5C" });
            FANTREHT.Add("QC", new String[] { "QC180A" });
            FANTREHT.Add("ESP", new String[] { "ESPECIAL" });
            Caract = new Caracteristic("FANTREHT", "Unidad de accionamiento", false, FANTREHT);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FANTREHT", Caract);
            OrderCaractComercial.Add(910, Caract);

            //Potencia del motor
            Caract = new Caracteristic("TNPOTENCIAMOTOR", "Potencia del motor", true, 0.0);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNPOTENCIAMOTOR", Caract);
            OrderCaractComercial.Add(920, Caract);

            // Paso de cadena peldaño/paleta
            SortedList FSTFKT = new SortedList();
            FSTFKT.Add("101,15", new String[] { "Paso t=101,15" });
            FSTFKT.Add("101,25", new String[] { "Paso t=101,25" });
            FSTFKT.Add("135", new String[] { "Paso t=135" });
            Caract = new Caracteristic("FSTFKT", "Paso de cadena peldaño/paleta", false, FSTFKT);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractComercial.Add("FSTFKT", Caract);
            OrderCaractComercial.Add(930, Caract);

            //Incremento cables cab inferior
            Caract = new Caracteristic("TNCR_OT_INCREM_CABLES_INFERIOR", "Incremento cables cab inferior", true, 0.0, false);
            CaractComercial.Add("TNCR_OT_INCREM_CABLES_INFERIOR", Caract);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            OrderCaractComercial.Add(950, Caract);

            //Incremento cables cab superior
            Caract = new Caracteristic("TNCR_OT_INCREM_CABLES_SUPERIOR", "Incremento cables cab superior", true, 0.0, false);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractComercial.Add("TNCR_OT_INCREM_CABLES_SUPERIOR", Caract);
            OrderCaractComercial.Add(960, Caract);

            #endregion

        }
        
        private void CreateCaractIng()
        {
            CaractIng = new SortedList();
            OrderCaractIng = new SortedList();
            Caracteristic Caract;

            //Tipo de esquema                      
            //SortedList ARMARIO = new SortedList();
            //ARMARIO.Add("LOCAL", new String[] { "Esquema de armario local" });
            //ARMARIO.Add("CHINO", new String[] { "Esquema de armario chino" });
            //Caracteristic Caract = new Caracteristic("ARMARIO", "Tipo de armario", false, ARMARIO);
            //Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            //CaractIng.Add("ARMARIO", Caract);
            //OrderCaractIng.Add(10, Caract);

            //Tipo de Maniobra             
            SortedList MANIOBRA = new SortedList();
            MANIOBRA.Add("ESTANDAR", new String[] { "Maniobra Estándar" });
            MANIOBRA.Add("BASIC", new String[] { "Maniobra Basic" });
            Caract = new Caracteristic("MANIOBRA", "Tipo de maniobra", false, MANIOBRA);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("MANIOBRA", Caract);
            OrderCaractIng.Add(20, Caract);

            //Corriente térmico motor
            Caract = new Caracteristic("ITERMICO", "Corriente térmico motor", true, 0);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractIng.Add("ITERMICO", Caract);
            OrderCaractIng.Add(30, Caract);

            //Corriente motor
            Caract = new Caracteristic("IMOTOR", "Corriente motor", true, 0);
            Caract.Textboxdata += new Caracteristic.TexboxDelegate(TextboxChanged);
            CaractIng.Add("IMOTOR", Caract);
            OrderCaractIng.Add(40, Caract);

            //Conexión Motor
            SortedList CONEXMOTOR = new SortedList();
            CONEXMOTOR.Add("D", new String[] { "Triangulo sin VVF" });
            CONEXMOTOR.Add("YD", new String[] { "Estrella-Triangulo sin VVF" });
            CONEXMOTOR.Add("VVF_D", new String[] { "VVF con bypass triangulo" });
            CONEXMOTOR.Add("VVF_YD", new String[] { "VVF con bypass estrella-triangulo" });
            CONEXMOTOR.Add("VVF", new String[] { "VVF sin bypass" });
            Caract = new Caracteristic("CONEXMOTOR", "Conexión Motor", false, CONEXMOTOR);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("CONEXMOTOR", Caract);
            OrderCaractIng.Add(50, Caract);

            //Seccion cable motor
            SortedList SECCABLEMOT = new SortedList();
            SECCABLEMOT.Add("2,5", new String[] { "2,5mm2 sin apantallar" });
            SECCABLEMOT.Add("4", new String[] { "4mm2 sin apantallar" });
            SECCABLEMOT.Add("4A", new String[] { "4mm2 apantallado" });
            SECCABLEMOT.Add("6", new String[] { "6mm2 sin apantallar" });
            SECCABLEMOT.Add("6A", new String[] { "6mm2 apantallado" });
            SECCABLEMOT.Add("10", new String[] { "10mm2 sin apantallar" });
            SECCABLEMOT.Add("10A", new String[] { "10mm2 apantallado" });
            SECCABLEMOT.Add("16", new String[] { "16mm2 sin apantallar" });
            SECCABLEMOT.Add("16A", new String[] { "16mm2 apantallado" });
            Caract = new Caracteristic("SECCABLEMOT", "Seccion cable motor", false, SECCABLEMOT);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("SECCABLEMOT", Caract);
            OrderCaractIng.Add(60, Caract);

            // Controlador
            SortedList TNCR_DO_CONTROL = new SortedList();
            TNCR_DO_CONTROL.Add("GEC", new String[] { "GEC" });
            TNCR_DO_CONTROL.Add("GEC+PLC", new String[] { "GEC+PLC" });
            TNCR_DO_CONTROL.Add("PLC", new String[] { "PLC Siemens" });
            TNCR_DO_CONTROL.Add("BUS", new String[] { "Sistema BUS" });
            TNCR_DO_CONTROL.Add("CUSTOM", new String[] { "Controlador según pedido" });
            Caract = new Caracteristic("TNCR_DO_CONTROL", "Controlador", false, TNCR_DO_CONTROL);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("TNCR_DO_CONTROL", Caract);
            OrderCaractIng.Add(70, Caract);

            // Tipo display
            SortedList TNCR_DO_DISPLAY_TYPE = new SortedList();
            TNCR_DO_DISPLAY_TYPE.Add("DDU", new String[] { "DDU" });
            TNCR_DO_DISPLAY_TYPE.Add("ESCATRONIC", new String[] { "Escatronic" });
            TNCR_DO_DISPLAY_TYPE.Add("CUSTOM", new String[] { "Otro" });
            Caract = new Caracteristic("TNCR_DO_DISPLAY_TYPE", "Tipo Display", false, TNCR_DO_DISPLAY_TYPE);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("TNCR_DO_DISPLAY_TYPE", Caract);
            OrderCaractIng.Add(80, Caract);

            // Envolvente Armario Exterior
            SortedList ENVOLV_ARM_EXT = new SortedList();
            ENVOLV_ARM_EXT.Add("M_1800x800x400", new String[] { "Metalica 1800x800x400" });
            ENVOLV_ARM_EXT.Add("I_1800x800x400", new String[] { "Inox. 1800x800x400" });
            ENVOLV_ARM_EXT.Add("M_1800x1000x400", new String[] { "Metalica 1800x1000x400" });
            ENVOLV_ARM_EXT.Add("I_1800x1000x400", new String[] { "Inox. 1800x1000x400" });
            ENVOLV_ARM_EXT.Add("M_1800x1200x400", new String[] { "Metalica 1800x1200x400" });
            ENVOLV_ARM_EXT.Add("I_1800x1200x400", new String[] { "Inox. 1800x1200x400" });
            ENVOLV_ARM_EXT.Add("CUSTOM", new String[] { "Customizado" });
            Caract = new Caracteristic("ENVOLV_ARM_EXT", "Envolvente Armario Exterior", false, ENVOLV_ARM_EXT);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("ENVOLV_ARM_EXT", Caract);
            OrderCaractIng.Add(90, Caract);

            // Paquete especial
            SortedList PAQUETE_ESP = new SortedList();
            PAQUETE_ESP.Add("NO", new String[] { "Ninguno" });
            PAQUETE_ESP.Add("FRANCIA", new String[] { "Paquete Francia" });
            PAQUETE_ESP.Add("MERCADONA", new String[] { "Paquete Mercadona" });
            PAQUETE_ESP.Add("SUBIDAS", new String[] { "Paquete Subidas" });
            Caract = new Caracteristic("PAQUETE_ESP", "Paquete especial", false, PAQUETE_ESP);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("PAQUETE_ESP", Caract);
            OrderCaractIng.Add(100, Caract);

            //Rearranque tras corte tensión
            SortedList POWER_OUTAGE_RESTART = new SortedList();
            POWER_OUTAGE_RESTART.Add("NO", new String[] { "No" });
            POWER_OUTAGE_RESTART.Add("SI", new String[] { "Si" });
            Caract = new Caracteristic("POWER_OUTAGE_RESTART", "Rearranque tras corte tensión", false, POWER_OUTAGE_RESTART);
            Caract.Comboboxdata += new Caracteristic.ComboboxDelegate(ComboboxChanged);
            CaractIng.Add("POWER_OUTAGE_RESTART", Caract);
            OrderCaractIng.Add(110, Caract);

        }

        public void createDefaultGECParam()
        {
            Dictionary<string, GEC> oElectricList = new Dictionary<string, GEC>();
            PathInfo EPLANPaths = new PathInfo();
            String path = String.Concat(EPLANPaths.Documents, "\\GEC_Default_Parameters.csv");

            using (Microsoft.VisualBasic.FileIO.TextFieldParser parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(path))
            {
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(";");
                while (!parser.EndOfData)
                {
                    //Processing row
                    string[] fields = parser.ReadFields();
                    oElectricList.Add(fields[0], new GEC(fields[0], fields[1]));

                }
            }

            GECParameterList=oElectricList;

            IDFunctions = new Dictionary<uint, string>();

            //ID Functions
            path = String.Concat(EPLANPaths.Documents, "\\ID_Functions.csv");

            using (Microsoft.VisualBasic.FileIO.TextFieldParser parser = new Microsoft.VisualBasic.FileIO.TextFieldParser(path))
            {
                parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
                parser.SetDelimiters(";");
                while (!parser.EndOfData)
                {
                    uint val;
                    //Processing row
                    string[] fields = parser.ReadFields();
                    uint.TryParse(fields[0], out val);
                    IDFunctions.Add(val, fields[1]);

                };
            }
        }

        public void createGECParamList()
        {
            PathInfo EPLANPaths = new PathInfo();
            string path;
            string[] filas;
            uint val;

            //GECParameterList = new Dictionary<string, GEC>();
            IDFunctions = new Dictionary<uint, string>();

            //ID Functions
            path = String.Concat(EPLANPaths.Documents, "\\ID_Functions.csv");
            filas = File.ReadAllLines(path);
            foreach (var fila in filas)
            {
                string[] sfila = fila.Split(';');
                if (sfila[0] != "")
                {
                    uint.TryParse(sfila[0], out val);
                    IDFunctions.Add(val, sfila[1]);
                }
            }
        }

        public bool checkValues()
        {
            bool res = true;

            foreach (Caracteristic c in OrderCaractComercial.Values)
            {
                if (c.combobox != null &&
                    c.combobox.SelectedIndex < 0)
                    return false;
                else if (c.textBox != null &&
                    c.textBox.Text == "")
                    return false;
            }

            foreach (Caracteristic c in OrderCaractIng.Values)
            {
                if (c.combobox != null &&
                    c.combobox.SelectedIndex < 0)
                    return false;
                else if (c.textBox != null &&
                    c.textBox.Text == "")
                    return false;
            }

            return res;
        }

        private void ComboboxChanged(string data, string reference)
        {
            ComboboxDataToConfigurador(data, reference);
        }

        private void TextboxChanged(string data, string reference)
        {
            TextBoxDataToConfigurador(data, reference);
        }

    }
}
