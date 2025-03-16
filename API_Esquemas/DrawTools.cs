using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.DataModel.MasterData;
using Eplan.EplApi.HEServices;
using EPLAN_API.SAP;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Media3D;
using Eplan.EplApi.DataModel.E3D;
using Eplan.EplApi.DataModel.Graphics;
using System.IO;
using Eplan.EplApi.Base.Enums;
using System.Globalization;
using System.Windows.Forms;
using Eplan.EplApi.DataModel.EObjects;


namespace EPLAN_API.User
{
    public class DrawTools
    {

        #region Interfaces
        // Primera interfaz
        public interface IAdvancedDraw
        {
            void SetGECParameter(Project oProject, Electric oElectric, string address, uint value, bool changeText = false);
        }
        #endregion

        #region Metodos auxiliares
        public void insertNewPage(Project oProject, string pageName, string pageBefore)
        {
            Dictionary<int, string> dictPages = GetPageTable(oProject);

            //Compruebo si ya esta insertada
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageName);

            if (key == 0)
            {
                //Despues de la página de
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageBefore);

                //Renumeramos páginas
                Renumber renumber = new Renumber();
                renumber.Pages(oProject.Pages, false, 1, 1, false, true, Renumber.Enums.SubPages.Retain);

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

            }

        }

        public Dictionary<int, string> GetPageTable(Project oProject)
        {

            Dictionary<int, string>  dictPages = new Dictionary<int, string>();


            for (int i = 0; i < oProject.Pages.Length; i++)
            {
                dictPages.Add(i, oProject.Pages[i].Properties.PAGE_NOMINATIOMN.ToMultiLangString().GetStringToDisplay(ISOCode.Language.L_en_US));
            }

            return dictPages;
        }

        public void insertWindowMacro(Project oProject, string pathMacro, char variante, string page, double x, double y)
        {
            int key;
            Insert oInsert = new Insert();

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);

            int nVariante = -1;
            

            if (variante >= 'A' && variante <= 'Z')
            {
                nVariante = variante - 'A';
            }
            else
            {
                nVariante = -1; // Manejo de error
            }

            oInsert.WindowMacro(pathMacro, nVariante, oProject.Pages[key], new PointD(x, y), Insert.MoveKind.Absolute);
        }

        public StorableObject[] insertWindowMacro_ObjCont(Project oProject, string pathMacro, char variante, string page, double x, double y)
        {
            int key;
            Insert oInsert = new Insert();
            Dictionary<int, string> dictPages = GetPageTable(oProject);

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

        public void insertPageMacro(Project oProject, string pageMacroPath, string pageBefore, string pageName)
        {
            int key;
            Insert oInsert = new Insert();
            PagePropertyList pagePropList;
            Dictionary<int, string> dictPages = GetPageTable(oProject);

            //Compruebo si ya esta insertada la página
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageName);

            if (key == 0)
            {
                //Despues de la página de "Control II"
                key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == pageBefore);

                //Renumeramos páginas
                Renumber renumber = new Renumber();
                renumber.Pages(oProject.Pages, false, 1, 1, false, true, Renumber.Enums.SubPages.Retain);

                pagePropList = new PagePropertyList();

                for (int i = oProject.Pages.Length - 1; i > key; i--)
                {
                    pagePropList.DESIGNATION_PLANT = oProject.Pages[i].Properties.DESIGNATION_PLANT;
                    pagePropList.DESIGNATION_LOCATION = oProject.Pages[i].Properties.DESIGNATION_LOCATION;
                    pagePropList.PAGE_COUNTER = i + 2;
                    oProject.Pages[i].SetName(pagePropList);
                }

                oInsert.PageMacro(pageMacroPath, oProject.Pages[key], oProject, false);
            }
        }

        public void insertLayoutMacro(Project oProject, string pathMacro, char variante, string page, double x, double y, string IME)
        {
            int key;
            Insert oInsert = new Insert();
            Dictionary<int, string> dictPages = GetPageTable(oProject);

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

        public StorableObject[] insert3DMacro(Project oProject, string pathMacro, char variante, InstallationSpace oInstallationSpace, double x, double y, double z, double angle)
        {
            int nVariant = -1;
            switch (variante)
            {
                case 'A':
                    nVariant = 0;
                    break;

                case 'B':
                    nVariant = 1;
                    break;

                case 'C':
                    nVariant = 2;
                    break;

                case 'D':
                    nVariant = 3;
                    break;

                case 'E':
                    nVariant = 4;
                    break;

                case 'F':
                    nVariant = 5;
                    break;

                case 'G':
                    nVariant = 6;
                    break;

                case 'H':
                    nVariant = 7;
                    break;

                case 'I':
                    nVariant = 8;
                    break;

                case 'J':
                    nVariant = 9;
                    break;

                case 'K':
                    nVariant = 10;
                    break;

                case 'L':
                    nVariant = 11;
                    break;

                case 'M':
                    nVariant = 12;
                    break;

                case 'N':
                    nVariant = 13;
                    break;

                case 'O':
                    nVariant = 14;
                    break;

                case 'P':
                    nVariant = 15;
                    break;

            }

            //preparing transformation

            Matrix3D oMatrix = new Matrix3D();
            Quaternion oQaternion = new Quaternion(new Vector3D(x, y, z), 0.2);
            oMatrix.Rotate(oQaternion);

            //preparing WindowMacro object                                                                  
            string strWindowMacroName = pathMacro;
            WindowMacro oWMacro = new WindowMacro();
            oWMacro.Open(strWindowMacroName, oProject, 0);

            //insert macro into an InstallationSpace
            Insert3D oInsert3D = new Insert3D();
            StorableObject[] arrStorableObjects = oInsert3D.WindowMacro(oWMacro, nVariant, oInstallationSpace,
            oMatrix, Insert3D.MoveKind.Absolute, WindowMacro.Enums.NumerationMode.None);
            return arrStorableObjects;
        }

        public StorableObject[] insert3DMacro(Project oProject, string pathMacro, char variante, InstallationSpace oInstallationSpace, double x, double y, double z)
        {
            int nVariant = -1;
            switch (variante)
            {
                case 'A':
                    nVariant = 0;
                    break;

                case 'B':
                    nVariant = 1;
                    break;

                case 'C':
                    nVariant = 2;
                    break;

                case 'D':
                    nVariant = 3;
                    break;

                case 'E':
                    nVariant = 4;
                    break;

                case 'F':
                    nVariant = 5;
                    break;

                case 'G':
                    nVariant = 6;
                    break;

                case 'H':
                    nVariant = 7;
                    break;

                case 'I':
                    nVariant = 8;
                    break;

                case 'J':
                    nVariant = 9;
                    break;

                case 'K':
                    nVariant = 10;
                    break;

                case 'L':
                    nVariant = 11;
                    break;

                case 'M':
                    nVariant = 12;
                    break;

                case 'N':
                    nVariant = 13;
                    break;

                case 'O':
                    nVariant = 14;
                    break;

                case 'P':
                    nVariant = 15;
                    break;

            }

            //preparing transformation

            Matrix3D oMatrix = new Matrix3D();
            oMatrix.Translate(new Vector3D(x, y, z));

            //preparing WindowMacro object                                                                  
            string strWindowMacroName = pathMacro;
            WindowMacro oWMacro = new WindowMacro();
            oWMacro.Open(strWindowMacroName, oProject, 0);

            //insert macro into an InstallationSpace
            Insert3D oInsert3D = new Insert3D();
            StorableObject[] arrStorableObjects = oInsert3D.WindowMacro(oWMacro, nVariant, oInstallationSpace,
            oMatrix, Insert3D.MoveKind.Absolute, WindowMacro.Enums.NumerationMode.None);
            return arrStorableObjects;
        }


        public long insertSymbol(Project oProject, string symbol, string symbolLibrary, char variante, string page, double x, double y)
        {
            int key;
            Dictionary<int, string> dictPages = GetPageTable(oProject);
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);
            Page oPage = oProject.Pages[key];
            string strSymbolLibName = symbolLibrary;
            string strSymbolName = symbol;
            SymbolLibrary oSymbolLibrary = new SymbolLibrary(oProject, strSymbolLibName);
            Symbol oSymbol = new Symbol(oSymbolLibrary, strSymbolName);

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

            SymbolVariant oSymbolVariant = new SymbolVariant(oSymbol, nVariante);
            SymbolReference sr = oSymbolVariant.Create(oPage);
            sr.Location = new PointD(x, y);

            return sr.ObjectIdentifier;
        }

        public long insertInterruptionPoint(Project oProject, string symbol, string symbolLibrary, char variante, string page, string deviceName, string propertySchema, double x, double y)
        {

            int key;
            Insert oInsert = new Insert();

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);
            
            Page oPage = oProject.Pages[key];
            string strSymbolLibName = symbolLibrary;
            string strSymbolName = symbol;
            SymbolLibrary oSymbolLibrary = new SymbolLibrary(oProject, strSymbolLibName);
            Symbol oSymbol = new Symbol(oSymbolLibrary, strSymbolName);

            int nVariante = -1;
            if (variante >= 'A' && variante <= 'Z')
            {
                nVariante = variante - 'A';
            }
            else
            {
                nVariante = -1; // Manejo de error
            }

            SymbolVariant oSymbolVariant = new SymbolVariant(oSymbol, nVariante);
            SymbolReference sr = oSymbolVariant.Create(oPage);
            sr.Location = new PointD(x, y);
            InterruptionPoint ip = (InterruptionPoint)sr;
            ip.Name = deviceName;
            ip.VisibleName = deviceName.Split('-')[1];
            foreach (SymbolReference.PropertyPlacementsSchema property in ip.PropertyPlacementsSchemas.All)
            {
                if (property.Name.Contains(propertySchema))
                {
                    ip.PropertyPlacementsSchemas.Selected = property;
                    break;
                }
            }

            return sr.ObjectIdentifier;
        }

        public void setRecordContactor(StorableObject[] oInsertedObjects, Caracteristic iMotor)
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

        public void changeFunctionTextPLCInput(Project oProject, Page page, string address, string newText)
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

        public void changeFunctionTextPLCInput(Project oProject, string spage, string address, string newText)
        {
            int key;
            Insert oInsert = new Insert();
            Dictionary<int, string> dictPages = GetPageTable(oProject);

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

        public void changeFunctionTextPLCInput(Project oProject, string address, string newText)
        {
            //other way to do the same:
            FunctionsFilter ff = new FunctionsFilter();
            ff.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.PLCTerminal;
            DMObjectsFinder objFinder = new DMObjectsFinder(oProject);

            //now we have all functions having category 'MOTOR' placed on page p
            Function[] functions = objFinder.GetFunctions(ff);

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

        public void deleteSymbol(Project oProject, string page, string name, string type, int posX = 0, int posY = 0)
        {
            Dictionary<int, string> dictPages = GetPageTable(oProject);
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                string pType = placement.GetType().Name;
                switch (pType)
                {
                    case "InterruptionPoint":
                        if (((InterruptionPoint)placement).Name.Equals(name))
                        {
                            placement.Remove();
                        }
                        break;

                    case "SymbolReference":
                        if (placement.Location.X == posX && placement.Location.Y == posY)
                        {
                            placement.Remove();
                        }
                        break;

                }
            }
        }

        public void moveSymbol(Project oProject, string page, int offsetX, int offsetY, int iniposX, int iniposY, int endposX, int endposY)
        {
            Dictionary<int, string> dictPages = GetPageTable(oProject);
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.X >= iniposX && placement.Location.X <= endposX)
                {
                    if (placement.Location.Y >= iniposY && placement.Location.Y <= endposY)
                    {
                        placement.Location = new PointD(placement.Location.X + offsetX, placement.Location.Y + offsetY);

                    }
                }
            }
        }

        public void deleteArea(Project oProject, string page, int iniposX, int iniposY, int endposX, int endposY)
        {
            Dictionary<int, string> dictPages = GetPageTable(oProject);
            int key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);
            Placement[] placements = oProject.Pages[key].AllPlacements;
            foreach (Placement placement in placements)
            {
                if (placement.Location.X >= iniposX && placement.Location.X <= endposX)
                {
                    if (placement.Location.Y >= iniposY && placement.Location.Y <= endposY)
                    {
                        placement.Remove();

                    }
                }
            }
        }

        public void SetCableLength(Project oProject, string page, string name, double length)
        {
            int key;
            Dictionary<int, string> dictPages = GetPageTable(oProject);
            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == page);
            Page oPage = oProject.Pages[key];

            FunctionsFilter functionsFilter = new FunctionsFilter();
            functionsFilter.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.Cable;
            functionsFilter.Page=oPage;

            DMObjectsFinder DMObjectsFinder = new DMObjectsFinder(oProject);

            Function[] functions = DMObjectsFinder.GetFunctions(functionsFilter);

            foreach (Function function in functions) 
            {
                if (function.VisibleName.Trim('-') == name) 
                {
                    PropertyValue CableLength = function.Properties.FUNC_CABLELENGTH;
                    CableLength.Set(length);
                }
            }
        }

        public void insertArticle(Project oProject, string ipage, string ifunction, string articleref, uint cantidad)
        {
            int key;
            Insert oInsert = new Insert();

            Dictionary<int, string> dictPages = GetPageTable(oProject);


            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == ipage);

            Page opage = oProject.Pages[key];

            opage.Filter.FunctionCategory = FunctionCategory.ArticleDefinitionPoint;
            Function[] f = opage.Functions;
            Function ofunction = null;

            foreach (Function function in f)
            {
                if (function.VisibleName.Equals(ifunction))
                {
                    ofunction = function;
                    break;
                }
            }

            if (ofunction == null)
                return;

            ofunction.AddArticleReference(articleref, "1", cantidad);

        }

        public Function insertDeviceLayout(Project oProject, string functionName, string functionPage, string mountingPlate, int posArticle, char macroVariant, string variant, string layouyPage, int posX, int posY)
        {
            int key;
            Insert oInsert = new Insert();

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            Function deviceFuncion = new Function();
            Function articlePlacememt = new Function();

            int nVariante = -1;
            switch (macroVariant)
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

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == layouyPage);
            Page oLayoutPage = oProject.Pages[key];

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == functionPage);

            Page ofunctionPage = oProject.Pages[key];

            foreach (Function function in ofunctionPage.Functions)
            {
                if (function.VisibleName.Contains(functionName) &&
                    function.IsMainFunction)
                {
                    deviceFuncion = function;
                    break;
                }
            }

            if (deviceFuncion.VisibleName == null)
                return null;

            FunctionsFilter ff = new FunctionsFilter();
            ff.FunctionCategory=FunctionCategory.MountingPlate;
            ff.Page = oLayoutPage;
            DMObjectsFinder objFinder = new DMObjectsFinder(oProject);
            Function[] functions = objFinder.GetFunctions(ff);


            foreach (Function function in functions)
            {
                if (function.VisibleName.Contains(mountingPlate))
                {
                    try
                    {
                        new Settings().SetStringSetting("USER.PanelLayoutGui.Settings.Gripper", "UpperLeft");
                        MountingPanelService mountingPanelService = new MountingPanelService();
                        //mountingPanelService.CreateArticlePlacement(function, deviceFuncion.Articles[posArticle].PartNr, new PointD(posX, posY + deviceFuncion.Articles[posArticle].Properties.ARTICLE_HEIGHT.ToDouble()), ref articlePlacememt);
                        mountingPanelService.CreateArticlePlacement(function, deviceFuncion.Articles[posArticle].PartNr, nVariante, new PointD(posX, posY), ref articlePlacememt);
                        articlePlacememt.Name = deviceFuncion.Name;
                        articlePlacememt.VisibleName = deviceFuncion.VisibleName;
                        //articlePlacememt.IdentifyingNameParts.FUNC_VISIBLEDEVICETAG = deviceFuncion.IdentifyingNameParts.FUNC_VISIBLEDEVICETAG;
                        articlePlacememt.PropertyPlacementsSchemas.Selected = articlePlacememt.PropertyPlacementsSchemas.DefaultScheme;

                    }

                    catch (Exception ex)
                    {
                        return null;
                    }
                    break;
                }
            }

            return articlePlacememt;

        }

        public Function insertDeviceLayout(Project oProject, string functionName, string functionPage, string mountingPlate, int posArticle, char macroVariant, string variant, string layouyPage, string functionNextComponent)
        {
            int key;
            Insert oInsert = new Insert();

            Dictionary<int, string> dictPages = GetPageTable(oProject);

            int posX = 0;
            int posY = 0;
            Function deviceFuncion = new Function();
            Function articlePlacememt = new Function();

            int nVariante = -1;
            switch (macroVariant)
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

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == layouyPage);
            Page oLayoutPage = oProject.Pages[key];

            key = dictPages.Keys.OfType<int>().FirstOrDefault(s => dictPages[s] == functionPage);

            Page ofunctionPage = oProject.Pages[key];

            foreach (Function function in ofunctionPage.Functions)
            {
                if (function.VisibleName.Contains(functionName) &&
                    function.IsMainFunction)
                {
                    deviceFuncion = function;
                    break;
                }
            }

            if (deviceFuncion.VisibleName == null)
                return null;



            foreach (Function function in oLayoutPage.Functions)
            {
                if (function.VisibleName.Contains(functionNextComponent))
                {
                    posX = function.Properties.INSTANCE_XCOORD.ToInt() +
                        function.Articles[0].Properties.ARTICLE_WIDTH.ToInt();
                    posY = function.Properties.INSTANCE_YCOORD.ToInt();
                }
            }

            FunctionsFilter ff = new FunctionsFilter();
            ff.FunctionCategory=FunctionCategory.MountingPlate;
            ff.Page = oLayoutPage;
            DMObjectsFinder objFinder = new DMObjectsFinder(oProject);
            Function[] functions = objFinder.GetFunctions(ff);

            foreach (Function function in functions)
            {
                if (function.VisibleName.Contains(mountingPlate))
                {
                    try
                    {
                        new Settings().SetStringSetting("USER.PanelLayoutGui.Settings.Gripper", "UpperLeft");
                        MountingPanelService mountingPanelService = new MountingPanelService();
                        mountingPanelService.CreateArticlePlacement(function, deviceFuncion.Articles[posArticle].PartNr, nVariante, new PointD(posX, posY), ref articlePlacememt);
                        articlePlacememt.Name = deviceFuncion.Name;
                        articlePlacememt.VisibleName = deviceFuncion.VisibleName;
                        articlePlacememt.PropertyPlacementsSchemas.Selected = articlePlacememt.PropertyPlacementsSchemas.DefaultScheme;
                    }
                    catch (Exception ex)
                    {
                        return null;
                    }
                    break;
                }
            }

            return articlePlacememt;

        }

        public void Insert3DDeviceIntoDINRail(Project oProject, string DINRailName, string deviceName, double offset, bool forcedPos=false, double forcedPosOffset=0, string rightTo=null)
        {

            /*
             *  There are two different types of mates: stored in database and generated in runtime. 
                First can be created by user and assign to a Placement3D. 
                Name of generated mate depends on type of object. 
                Mate named 'V1'is placed on first vertex (lowest left in local coordinate system of the object). 
                Numeration of those mates is done in counterclockwise direction. 
                Very similar are mates which name starts with 'M', only those mates are placed in the middle of edges and the first on is the lowest one (in the local coordinate system). 
                Third type of mate that can be found is named "C". This one is placed in the center of object. 
                Check mate description for more information about current mate.
             */


            //searching DIN Rail
            string str3DFunction = DINRailName;
            Functions3DFilter oFunctions3DFilter = new Functions3DFilter();
            Function3DPropertyList oFunction3DPropertyList = new Function3DPropertyList();
            oFunction3DPropertyList.FUNC_VISIBLEDEVICETAG = str3DFunction;
            oFunctions3DFilter.SetFilteredPropertyList(oFunction3DPropertyList);
            Function3D[] oFunctions3D = new DMObjectsFinder(oProject).GetFunctions3D(oFunctions3DFilter);
            MountingRail mountingRail = oFunctions3D[0] as MountingRail;

            //searching device
            FunctionsFilter oFunctionsFilter = new FunctionsFilter();
            FunctionPropertyList functionPropertyList = new FunctionPropertyList();
            functionPropertyList.FUNC_VISIBLEDEVICETAG = deviceName;
            functionPropertyList.FUNC_MAINFUNCTION = true;
            oFunctionsFilter.SetFilteredPropertyList(functionPropertyList);
            Function[] oFunctions = new DMObjectsFinder(oProject).GetFunctions(oFunctionsFilter);
            Function device = oFunctions[0];

            //searching rightToDevice
            str3DFunction = rightTo;
            oFunctions3DFilter = new Functions3DFilter();
            oFunction3DPropertyList = new Function3DPropertyList();
            oFunction3DPropertyList.FUNC_VISIBLEDEVICETAG = str3DFunction;
            oFunctions3DFilter.SetFilteredPropertyList(oFunction3DPropertyList);
            oFunctions3D = new DMObjectsFinder(oProject).GetFunctions3D(oFunctions3DFilter);
            Component rightToDevice3D = null;
            try
            {
                rightToDevice3D = oFunctions3D[0] as Component;
            }
            catch (Exception ex) 
            {
                ;
            }


            double pos = 0;
            double iniPos = 0;
            bool isfirst = true;
            if (!forcedPos && rightTo == null)
            {
                foreach (Component c in mountingRail.Children)
                {
                    //comprobar si es el primer elemento
                    double M4Pos = c.FindSourceMate("M4", Mate.Enums.PlacementOptions.None).Transformation.OffsetX;
                    if (isfirst || M4Pos < iniPos)
                        iniPos = M4Pos;

                    isfirst = false;

                    double M2pos = c.FindSourceMate("M2", Mate.Enums.PlacementOptions.None).Transformation.OffsetX;
                    if (M2pos > pos)
                        pos = M2pos;
                }
            }
            else if (forcedPos)
            {
                iniPos = 0;
                pos = forcedPosOffset;
            }

            Component oComponent = new Component();
            oComponent.Create(oProject, device.ArticleReferences[0].Article.PartNr, "1");
            oComponent.Name = device.Name;
            oComponent.VisibleName = deviceName;
            oComponent.Parent = mountingRail;
            if (rightTo != null)
                oComponent.FindSourceMate("M4", Mate.Enums.PlacementOptions.None).SnapTo(rightToDevice3D.FindTargetMate("M2", false));
            else
                oComponent.FindSourceMate("M4", Mate.Enums.PlacementOptions.None).SnapTo(mountingRail.BaseMate, pos - iniPos + offset);
        }

        public void Insert3DDeviceIntoMountingPlate(Project oProject, string deviceName, double x, double y, string rightTo = null, double rightToOffset=0, string varianteRef="1", int articleRef=0, double planeoffset=0)
        {
            /*
             *  There are two different types of mates: stored in database and generated in runtime. 
                First can be created by user and assign to a Placement3D. 
                Name of generated mate depends on type of object. 
                Mate named 'V1'is placed on first vertex (lowest left in local coordinate system of the object). 
                Numeration of those mates is done in counterclockwise direction. 
                Very similar are mates which name starts with 'M', only those mates are placed in the middle of edges and the first on is the lowest one (in the local coordinate system). 
                Third type of mate that can be found is named "C". This one is placed in the center of object. 
                Check mate description for more information about current mate.
             */


            //searching Mouning plate
            Functions3DFilter oFunctions3DFilter = new Functions3DFilter();
            Function3DPropertyList oFunction3DPropertyList = new Function3DPropertyList();
            oFunctions3DFilter.FunctionCategory = FunctionCategory.MountingPlate;
            //oFunctions3DFilter.InstallationSpace.VisibleName = "MAIN";
            //oFunctions3DFilter.SetFilteredPropertyList(oFunction3DPropertyList
            Function3D[] oFunctions3D = new DMObjectsFinder(oProject).GetFunctions3D(oFunctions3DFilter);
            MountingPanel mountingPanel = oProject.InstallationSpaces[0].Children[0] as MountingPanel;

            //searching device
            FunctionsFilter oFunctionsFilter = new FunctionsFilter();
            FunctionPropertyList functionPropertyList = new FunctionPropertyList();
            functionPropertyList.FUNC_VISIBLEDEVICETAG = deviceName;
            functionPropertyList.FUNC_MAINFUNCTION = true;
            oFunctionsFilter.SetFilteredPropertyList(functionPropertyList);
            Function[] oFunctions = new DMObjectsFinder(oProject).GetFunctions(oFunctionsFilter);
            Function device = oFunctions[0];

            //searching rightToDevice
            string str3DFunction = rightTo;
            oFunctions3DFilter = new Functions3DFilter();
            oFunction3DPropertyList = new Function3DPropertyList();
            oFunction3DPropertyList.FUNC_VISIBLEDEVICETAG = str3DFunction;
            oFunctions3DFilter.SetFilteredPropertyList(oFunction3DPropertyList);
            oFunctions3D = new DMObjectsFinder(oProject).GetFunctions3D(oFunctions3DFilter);
            Component rightToDevice3D = null;
            try
            {
                rightToDevice3D = oFunctions3D[0] as Component;
            }
            catch (Exception ex)
            {
                ;
            }

            Component oComponent = new Component();
            oComponent.Create(oProject, device.ArticleReferences[articleRef].Article.PartNr, varianteRef);
            oComponent.Name = device.Name;
            oComponent.VisibleName = deviceName;
            oComponent.Parent = mountingPanel;
            if (rightTo==null)
            {
                oComponent.FindSourceMate("V4", Mate.Enums.PlacementOptions.None).SnapTo(mountingPanel.Planes[0].BaseMate, 0, x, y);
            }
            if (planeoffset > 0)
            {
                oComponent.MoveRelative(0,0,planeoffset);
            }
        }

        #region GEC Parameters

        public void SetGECParameter(Project oProject, Electric oElectric, string address, uint value, bool changeText=false)
        {
            oElectric.GECParameterList[address].setValue(value);
            if (changeText)
            {
                changeFunctionTextPLCInput(oProject, address, oElectric.IDFunctions[oElectric.GECParameterList[address].value]);
            }
        }

        public void SetGECParameter(Project oProject, Electric oElectric, string address, string value, bool changeText = false)
        {
            oElectric.GECParameterList[address].setValue(value);
            if (changeText)
            {
                changeFunctionTextPLCInput(oProject, address, oElectric.IDFunctions[oElectric.GECParameterList[address].value]);
            }
        }

        public void calcParmGEC_Basic(Project oProject, Electric electric)
        {
            //electric.GECParameterList = createDefaultGECParam();    // Instancia lista parametros

            // Funcion para generar parametros GEC
            Caracteristic Velocidad = (Caracteristic)electric.CaractComercial["FGESCHW"];          // Velocidad  (m/s)
            Caracteristic MotorRPM = (Caracteristic)electric.CaractComercial["FACH"];              // RPM Motor
            Caracteristic FrenoAux = (Caracteristic)electric.CaractComercial["FZUSBREMSE"];        // Freno auxiliar en eje princip.
            Caracteristic Producto = (Caracteristic)electric.CaractComercial["FMODELL"];           // Modelo de producto
            Caracteristic PasoCadena = (Caracteristic)electric.CaractComercial["FSTFKT"];          // Paso de cadena (mm)
            Caracteristic Inclinacion = (Caracteristic)electric.CaractComercial["FNEIGUNG"];       // Angulo de inclinacion (º)
            Caracteristic motorConnection = (Caracteristic)electric.CaractIng["CONEXMOTOR"];
            Caracteristic bypass = (Caracteristic)electric.CaractComercial["TNCR_OT_BYPASS_VARIADOR"];
            Caracteristic deteccionPersonas = (Caracteristic)electric.CaractComercial["FLICHTINT"];
            Caracteristic modofuncionamiento = (Caracteristic)electric.CaractComercial["FBETRART"];
            Caracteristic sCadena = (Caracteristic)electric.CaractComercial["TNCR_S_DRIVE_CHAIN"];
            Caracteristic sRoturaPasamanos = (Caracteristic)electric.CaractComercial["F09ZUB1"];
            Caracteristic sDesgasteFrenos = (Caracteristic)electric.CaractComercial["F01ZUB"];
            Caracteristic sDobleFreno = (Caracteristic)electric.CaractComercial["FBREMSE2"];
            Caracteristic bombaEngrase = (Caracteristic)electric.CaractComercial["TNCR_ENGRASE_AUTOMATICO"];
            Caracteristic detectorAgua = (Caracteristic)electric.CaractComercial["TNCR_OT_NIVEL_AGUA"];
            Caracteristic aceiteReductor = (Caracteristic)electric.CaractComercial["TNCR_SENSOR_ACEITE_REDUC"];
            Caracteristic sistAhorro = (Caracteristic)electric.CaractComercial["TNCR_SD_SIST_AHORRO"];
            Caracteristic llavinAutoCont = (Caracteristic)electric.CaractComercial["LLAVES_AUT_CONT"];
            Caracteristic llavinLocalRemoto = (Caracteristic)electric.CaractComercial["LLAVES_LOCAL_REM"];
            Caracteristic llavinParo = (Caracteristic)electric.CaractComercial["LLAVES_PARO"];
            Caracteristic sistemaAnden = (Caracteristic)electric.CaractComercial["FWIEDERB"];
            Caracteristic desarrollo = (Caracteristic)electric.CaractComercial["TNCR_OT_DESARROLLO"];
            Caracteristic normativa = (Caracteristic)electric.CaractComercial["FNORM"];
            Caracteristic ubicacionControlador = (Caracteristic)electric.CaractComercial["F53ZUB7"];
            Caracteristic stopCarritos = (Caracteristic)electric.CaractComercial["TNCR_POSTE_STOP_CARRITOS"];
            Caracteristic sMicrosZocalo = (Caracteristic)electric.CaractComercial["TNCR_OT_NUM_MICROCONT"];
            Caracteristic sBuggy = (Caracteristic)electric.CaractComercial["F04ZUB"];
            Caracteristic sPeines = (Caracteristic)electric.CaractComercial["FKAMMPLHK"];
            Caracteristic desnivel = (Caracteristic)electric.CaractComercial["FHOEHEV"];
            Caracteristic lCabezaSup = (Caracteristic)electric.CaractComercial["FOT"];
            Caracteristic lCabezaInf = (Caracteristic)electric.CaractComercial["FUT"];
            Caracteristic contactoFuego = (Caracteristic)electric.CaractComercial["TNCR_CONTACTO_FUEGO"];
            Caracteristic paqueteEspecial = (Caracteristic)electric.CaractIng["PAQUETE_ESP"];
            Caracteristic cerrojo = (Caracteristic)electric.CaractComercial["TNCR_OT_CERROJO_MANTENIMIENTO"];
            Caracteristic stopAdicional = (Caracteristic)electric.CaractComercial["TNCR_OT_E_STOP_ADICIONAL"];
            Caracteristic powerOutage = (Caracteristic)electric.CaractIng["POWER_OUTAGE_RESTART"];

            //uint PulsosNAB;                              // Pulsos/rev eje ppal NAB
            //uint PulsosTrinqMagn;                        // Pulsos/rev eje ppal trinquete magnetico


            #region Safety 
            //******************************************************************************
            //------------- SAFETY PARAMETERS ----------------------------------------------
            //******************************************************************************

            #region Safety Parameters
            //S1	MANUFACTURING_ORDER
            SetGECParameter(oProject, electric, "S1", "11500xxxxx");

            //S2 CODE
            if (normativa.CurrentReference.Equals("EN"))
                //EN115
                SetGECParameter(oProject, electric, "S2", (uint)GEC.Code.EN115);
            else
                //ASME/B44
                SetGECParameter(oProject, electric, "S2", (uint)GEC.Code.ASME_B44);

            //S3	NOMINAL_SPEED
            SetGECParameter(oProject, electric, "S3", (uint)(Velocidad.NumVal * 100));

            //S4	MOTOR_RPM
            SetGECParameter(oProject, electric, "S4", (uint)MotorRPM.NumVal);

            //S5	MOTOR_PULSE_PER_REV
            if (Producto.CurrentReference.Equals("IWALK"))
                SetGECParameter(oProject, electric, "S5", 4);
            else
                SetGECParameter(oProject, electric, "S5", 10);

            //S6	MAIN_SHAFT_RPH
            uint rph = (uint)calcRPH(Producto.CurrentReference, PasoCadena.CurrentReference, Velocidad.NumVal);
            if (rph != 0)
            {
                SetGECParameter(oProject, electric, "S6", rph);
            }

            //S7	MAIN_SHAFT_PULSE_PER_REV
            if (FrenoAux.CurrentReference != null)
            {
                if (FrenoAux.CurrentReference.Equals("NAB"))
                    SetGECParameter(oProject, electric, "S7", calcPPR(870, 5));
                else if (FrenoAux.CurrentReference.Equals("HWSPERRKMAGN"))
                    SetGECParameter(oProject, electric, "S7", calcPPR(580, 5));
            }


            //S8	ROLLER_HR_RADIUS
            if (Producto.CurrentReference == "ORINOCO" && Inclinacion.NumVal == 12)
                SetGECParameter(oProject, electric, "S8", 35);
            else
                SetGECParameter(oProject, electric, "S8", 50);

            //S9	HR_PULSES_PER_REV
            if (Producto.CurrentReference.Equals("VELINO_CLASSIC") ||
               Producto.CurrentReference.Equals("TUGELA_CLASSIC"))
                SetGECParameter(oProject, electric, "S9", 2);
            else
                SetGECParameter(oProject, electric, "S9", 4);

            //S10	HR_FAULT_TIME
            SetGECParameter(oProject, electric, "S10", 5);

            //S11 STEP_WIDTH
            if (Producto.CurrentReference.Equals("IWALK"))
                SetGECParameter(oProject, electric, "S11", 127);
            else
                SetGECParameter(oProject, electric, "S11", 405);

            //S12 DELAY_NO_PULSE_CHECKING
            SetGECParameter(oProject, electric, "S12", 5000);

            //S13 SPEED_SENSOR_INSTALLATION
            if (FrenoAux.CurrentReference.Equals("HWSPERRKMAGN") ||
                FrenoAux.CurrentReference.Equals("NAB"))
                SetGECParameter(oProject, electric, "S13", (uint)GEC.MotorSensor.Two_Sensors_in_motor_Two_Main_Shaft);
            else
                SetGECParameter(oProject, electric, "S13", (uint)GEC.MotorSensor.Two_Sensors_in_motor);

            //S14 UNDERSPEED_TIME
            SetGECParameter(oProject, electric, "S14", 5000);

            //S15 END_SAFETY_STRING_FAULT_TYPE
            SetGECParameter(oProject, electric, "S15", 2);

            //S16 CONTACTOR_FEEDBACK_FILTER
            SetGECParameter(oProject, electric, "S16", 2000);

            //S17 CONTACTORS_TIMEOUT
            SetGECParameter(oProject, electric, "S17", 3);

            //S18	CONTACTOR_FB1_MASK
            //K1.1 / K1.2 star order motor 1
            SetGECParameter(oProject, electric, "S18", (uint)GEC.ContactorFB.K1_1_K1_2);

            //S19	CONTACTOR_FB2_MASK
            //K2.1 / K2.1.1 star order motor 1
            //K2.2 / K2.2.1 delta order motor 1
            if (motorConnection.CurrentReference.Equals("YD") ||
            motorConnection.CurrentReference.Equals("VVF_YD"))
                SetGECParameter(oProject, electric, "S19", (uint)GEC.ContactorFB.K2_1_K2_2);


            //S20	CONTACTOR_FB3_MASK
            //K10.1 / K10.2 VVVF operation
            if (motorConnection.CurrentReference.Equals("VVF_D") ||
                motorConnection.CurrentReference.Equals("VVF_YD") ||
                motorConnection.CurrentReference.Equals("VVF"))
                SetGECParameter(oProject, electric, "S20", (uint)GEC.ContactorFB.K10_2_K10_1);

            //S21	CONTACTOR_FB4_MASK
            //K10 delta operation / bypass of VVVF
            if (motorConnection.CurrentReference.Equals("VVF_D") ||
                motorConnection.CurrentReference.Equals("VVF_YD"))
                SetGECParameter(oProject, electric, "S21", (uint)GEC.ContactorFB.K10);

            //S22 CONTACTOR_FB5_MASK
            SetGECParameter(oProject, electric, "S22", (uint)GEC.ContactorFB.empty);

            //S23 CONTACTOR_FB6_MASK
            SetGECParameter(oProject, electric, "S23", (uint)GEC.ContactorFB.empty);

            //S24 CONTACTOR_FB7_MASK
            SetGECParameter(oProject, electric, "S24", (uint)GEC.ContactorFB.empty);

            //S25 CONTACTOR_FB8_MASK
            SetGECParameter(oProject, electric, "S25", (uint)GEC.ContactorFB.empty);

            //S26 KEY_MINIMUM_TIME
            SetGECParameter(oProject, electric, "S26", 500);

            //S27 UP_DOWN_ALLOWED
            SetGECParameter(oProject, electric, "S27", 0);

            //S28	AUTCONT_OPTIONS
            if (modofuncionamiento.CurrentReference.Equals("INTERM") ||
               modofuncionamiento.CurrentReference.Equals("SG") ||
               modofuncionamiento.CurrentReference.Equals("SGBV"))
            {
                if (sistAhorro.CurrentReference.Contains("VA"))
                    SetGECParameter(oProject, electric, "S28", (uint)GEC.Mode.Intermittent_Standby);
                else
                    SetGECParameter(oProject, electric, "S28", (uint)GEC.Mode.Intermittent);

            }
            else if (modofuncionamiento.CurrentReference.Equals("BV"))
                SetGECParameter(oProject, electric, "S28", (uint)GEC.Mode.Standby);

            //S29   DIAGNOSTIC_BOARD_L2_QUANTITY
            /*  No available    */

            //S30   TANDEM
            SetGECParameter(oProject, electric, "S30", (uint)GEC.Tandem.No_Tandem);

            //S31    INSPECTION_CATCH_THE_MOTOR
            SetGECParameter(oProject, electric, "S31", (uint)GEC.Active.Enable);

            //S32   RESET_FROM_INSPECTION_CONTROL
            SetGECParameter(oProject, electric, "S32", (uint)GEC.Active.Enable);

            //S33	AUX_BRAKE_SUPERVISION_TIME
            SetGECParameter(oProject, electric, "S33", 10);

            //S34	AUX_BRAKE_ENABLE
            if (FrenoAux.CurrentReference.Equals("HWSPERRKMAGN") ||
                FrenoAux.CurrentReference.Equals("NAB"))
                SetGECParameter(oProject, electric, "S34", (uint)GEC.Active.Enable);
            else
                SetGECParameter(oProject, electric, "S34", (uint)GEC.Active.Disable);

            //S35   CAPACITOR_TIME_MEASUREMENT
            /*  No available    */

            //S36   RADAR_TYPE
            SetGECParameter(oProject, electric, "S36", (uint)GEC.PeopleDetectionSensor.One_Input);

            //S37   LIGHT_BARRIER _COMBS_AREA_TYPE
            SetGECParameter(oProject, electric, "S37", (uint)GEC.PeopleDetectionSensor.One_Input);

            //S38   LIGHT_BARRIER_ENTRY_TYPE
            SetGECParameter(oProject, electric, "S38", (uint)GEC.PeopleDetectionSensor.One_Input);

            //S39	TIME_TRANSPORTATION (TIME_LONG)
            double fDesarrollo;
            //Calculate length
            if (desarrollo.NumVal == 0)
            {
                fDesarrollo = desnivel.NumVal / Math.Sin(Math.PI * Inclinacion.NumVal / 180.0);
                if (lCabezaInf.NumVal != 0 && lCabezaSup.NumVal != 0)
                {
                    fDesarrollo = fDesarrollo + (lCabezaInf.NumVal + lCabezaSup.NumVal) / 1000;
                }
                //estimate length
                else
                    fDesarrollo = fDesarrollo + 6.0;
            }
            else
                fDesarrollo = desarrollo.NumVal;

            if (normativa.CurrentReference.Equals("EN"))
                SetGECParameter(oProject, electric, "S39", (uint)(fDesarrollo / Velocidad.NumVal));
            else
                SetGECParameter(oProject, electric, "S39", (uint)(fDesarrollo / Velocidad.NumVal * 3));

            //S40	TIME_DIRECTION_INDICATION (TIME_SHORT)
            if (normativa.CurrentReference.Equals("EN"))
                SetGECParameter(oProject, electric, "S40", 10);
            else
                SetGECParameter(oProject, electric, "S40", (uint)(fDesarrollo / Velocidad.NumVal * 3));

            //S41   TIME_REVERSING
            /*  No available    */

            //S42   SAFETY_CURTAIN_LONG_TIME
            SetGECParameter(oProject, electric, "S42", 60);

            //S43   PULSE_SIGNALS_MINIMUM_LAG
            SetGECParameter(oProject, electric, "S43", 300);

            //S44   DRIVE_CHAIN_DELAY
            SetGECParameter(oProject, electric, "S44", 2000);

            //S45	DRIVE_CHAIN_AUX_BRAKE
            if (FrenoAux.CurrentReference.Equals("HWSPERRKMAGN") ||
                   FrenoAux.CurrentReference.Equals("NAB"))
                SetGECParameter(oProject, electric, "S45", (uint)GEC.Active.Enable);
            else
                SetGECParameter(oProject, electric, "S45", (uint)GEC.Active.Disable);

            //S46   AUX_BRAKE_ACTIVATION_DELAY_AFTER_STOP
            SetGECParameter(oProject, electric, "S46", 3);

            //S47   MOTOR_TRUNDLE
            SetGECParameter(oProject, electric, "S47", (uint)GEC.Active.Disable);

            //S48   AUX_BRAKE_UNBLOCK
            if (FrenoAux.CurrentReference.Equals("HWSPERRKMAGN") ||
                   FrenoAux.CurrentReference.Equals("NAB"))
                SetGECParameter(oProject, electric, "S47", (uint)GEC.Active.Enable);
            else
                SetGECParameter(oProject, electric, "S47", (uint)GEC.Active.Disable);

            //S49 TO S64 Not available

            //S65   AUTO_RESTART_WITH_SAFETY_CURTAIN
            if (sistemaAnden.CurrentReference.Equals("WBEREITSCH"))
                SetGECParameter(oProject, electric, "S65", (uint)GEC.Active.Enable);
            else
                SetGECParameter(oProject, electric, "S65", (uint)GEC.Active.Disable);

            //S66   ELECTRICAL BRAKING
            if (FrenoAux.CurrentReference.Equals("FRENO_VA"))
                SetGECParameter(oProject, electric, "S66", (uint)GEC.Active.Enable);
            else
                SetGECParameter(oProject, electric, "S66", (uint)GEC.Active.Disable);

            //S67   VFD TYPE
            SetGECParameter(oProject, electric, "S67", 0);

            //S68   ELEC BRAKING TIMEOUT
            SetGECParameter(oProject, electric, "S68", 1000);

            //S69   VFD STOPPING TIME
            SetGECParameter(oProject, electric, "S69", 1900);

            //S70   DECELERATION DISTANCE ERROR
            SetGECParameter(oProject, electric, "S70", 500);

            //S71   MAX DECELERATION
            SetGECParameter(oProject, electric, "S71", 80);

            //S72   AUTOMATIC_RESTART_AFTER_POWER_OUTAGE
            if (powerOutage.CurrentReference.Equals("SI"))
                SetGECParameter(oProject, electric, "S72", (uint)GEC.Active.Enable);
            else
                SetGECParameter(oProject, electric, "S72", (uint)GEC.Active.Disable);

            #endregion

            #region Safety Inputs
            //SI15	SF Safety Input 7
            SetGECParameter(oProject, electric, "SI15", (uint)GEC.Param.Empty);

            //SI16	SF Safety Input 8
            SetGECParameter(oProject, electric, "SI16", (uint)GEC.Param.Empty);

            //SI17	SF Safety Input 9
            SetGECParameter(oProject, electric, "SI17", (uint)GEC.Param.Empty);

            //SI18	SF Safety Input 10
            SetGECParameter(oProject, electric, "SI18", (uint)GEC.Param.Empty);

            //SI19	SF Safety Input 11
            SetGECParameter(oProject, electric, "SI19", (uint)GEC.Param.Up_maint_order);

            //SI20	SF Safety Input 12
            SetGECParameter(oProject, electric, "SI20", (uint)GEC.Param.Down_maint_order);

            //SI21	SF Safety Input 13
            SetGECParameter(oProject, electric, "SI21", (uint)GEC.Param.Empty);

            //SI22	SF Safety Input 14
            SetGECParameter(oProject, electric, "SI22", (uint)GEC.Param.Empty);

            //SI23 SF Safety Input 15 X23
            SetGECParameter(oProject, electric, "SI23", (uint)GEC.Param.Empty);

            //SI24    SF Safety Input 16
            SetGECParameter(oProject, electric, "SI24", (uint)GEC.Param.Empty);

            //SI25 SF Safety Input 17
            SetGECParameter(oProject, electric, "SI25", (uint)GEC.Param.Empty);

            //SI26 SF Safety Input 18
            SetGECParameter(oProject, electric, "SI26", (uint)GEC.Param.Empty);
            #endregion

            #endregion

            #region Control
            //******************************************************************************
            //------------- CONTROL PARAMETERS ---------------------------------------------
            //******************************************************************************

            #region Control Parameter
            //C1	MOTOR_CONNECTION 
            switch (motorConnection.CurrentReference)
            {
                case "VVF_YD":
                    //GEC Parameter
                    SetGECParameter(oProject, electric, "C1", (uint)GEC.MotorConnection.Nominal_load_inverter_Y_D);
                    break;
                case "VVF_D":
                    //GEC Parameter
                    SetGECParameter(oProject, electric, "C1", (uint)GEC.MotorConnection.Nominal_load_inverter_Delta);
                    break;
                case "YD":
                    //GEC Parameter
                    SetGECParameter(oProject, electric, "C1", (uint)GEC.MotorConnection.Y_D);
                    break;
                case "D":
                    //GEC Parameter
                    SetGECParameter(oProject, electric, "C1", (uint)GEC.MotorConnection.Delta);
                    break;
                case "VVF":
                    //GEC Parameter
                    SetGECParameter(oProject, electric, "C1", (uint)GEC.MotorConnection.Nominal_load_inverter_no_wired_bypass);
                    break;
            }

            //C2	TIME_LOW_SPEED
            SetGECParameter(oProject, electric, "C2", 20);

            //C3	ADDITIONAL_DIRECTION_INDICATION_TIME
            SetGECParameter(oProject, electric, "C3", 0);

            //C4	ADDITIONAL_TRANSPORTATION_TIME
            SetGECParameter(oProject, electric, "C4", 0);

            //C5	ADDITIONAL_REVERSING_TIME
            SetGECParameter(oProject, electric, "C5", 0);

            //C6	OIL_PUMP_CONTROL
            SetGECParameter(oProject, electric, "C6", (uint)GEC.Active.Enable);

            //C7	OIL_PUMP1_TIMER_ON
            SetGECParameter(oProject, electric, "C7", Calc_OIL_PUMP1_TIMER_ON(electric));

            //C8	OIL_PUMP1_CYCLE_TIME
            SetGECParameter(oProject, electric, "C8", Calc_OIL_PUMP1_CYCLE_TIME(electric));

            //C9	STAR_DELTA_DELAY
            SetGECParameter(oProject, electric, "C9", 5);

            //C10	TIMETABLE_MODE
            SetGECParameter(oProject, electric, "C10", (uint)GEC.Active.Disable);

            //C11	TIME_CHANGE_TO_MODE1_HOUR
            SetGECParameter(oProject, electric, "C11", 0);

            //C12	TIME_CHANGE_TO_MODE1_MIN
            SetGECParameter(oProject, electric, "C12", 0);

            //C13	TIME_CHANGE_TO_MODE2_HOUR
            SetGECParameter(oProject, electric, "C13", 0);

            //C14	TIME_CHANGE_TO_MODE2_MIN
            SetGECParameter(oProject, electric, "C14", 0);

            //C15	LANGUAGE
            SetGECParameter(oProject, electric, "C15", (uint)GEC.Language.English);

            //C16	CHANGE SUMERWINTER TIME
            SetGECParameter(oProject, electric, "C16", (uint)GEC.Active.Enable);

            //C17	SPEED_MEASUREMENT_UNIT
            SetGECParameter(oProject, electric, "C17", 0);

            //C18	TIMETABLE_MODE1
            SetGECParameter(oProject, electric, "C18", (uint)GEC.Mode.No_People_Detection);

            //C19	TIMETABLE_MODE2
            SetGECParameter(oProject, electric, "C19", (uint)GEC.Mode.No_People_Detection);

            //C20	HEATER_LOW_LEVEL
            SetGECParameter(oProject, electric, "C20", 0);

            //C21	HEATER_HIGH_LEVEL
            SetGECParameter(oProject, electric, "C21", 0);

            //C22	IP_ADDRESS_BYTE1
            SetGECParameter(oProject, electric, "C22", 192);

            //C23	IP_ADDRESS_BYTE2
            SetGECParameter(oProject, electric, "C23", 168);

            //C24	IP_ADDRESS_BYTE3
            SetGECParameter(oProject, electric, "C24", 0);

            //C25	IP_ADDRESS_BYTE4
            //string str_IP = tB_OE.Text.Substring(tB_OE.Text.Length - 2, 2);
            string str_IP = "10";
            uint IP = (uint)Int32.Parse(str_IP);
            if (IP == 00)
            {
                IP = 100;
            }
            else if (IP == 01)
            {
                IP = 101;
            }
            SetGECParameter(oProject, electric, "C25", IP);

            //C26	SUBNET_MASK_BYTE1
            SetGECParameter(oProject, electric, "C26", 255);

            //C27	SUBNET_MASK_BYTE2
            SetGECParameter(oProject, electric, "C27", 255);

            //C28	SUBNET_MASK_BYTE3
            SetGECParameter(oProject, electric, "C28", 255);

            //C29	SUBNET_MASK_BYTE4
            SetGECParameter(oProject, electric, "C29", 0);

            //C30	GATEWAY_BYTE1
            SetGECParameter(oProject, electric, "C30", 192);

            //C31	GATEWAY_BYTE2
            SetGECParameter(oProject, electric, "C31", 168);

            //C32	GATEWAY_BYTE3
            SetGECParameter(oProject, electric, "C32", 0);

            //C33	GATEWAY_BYTE4
            SetGECParameter(oProject, electric, "C33", 1);

            //C34	NODE_NUMBER
            SetGECParameter(oProject, electric, "C34", 0);

            //C35	RGB
            SetGECParameter(oProject, electric, "C35", (uint)GEC.Active.Disable);

            //C36	RGB_MOVING_LIGTH
            SetGECParameter(oProject, electric, "C36", (uint)GEC.Active.Disable);

            //C37	DIAGNOSTIC_BOARD_L1_QUANTITY
            SetGECParameter(oProject, electric, "C37", 2);

            //C38	UPPER_DIAG_SS_LENGTH
            SetGECParameter(oProject, electric, "C38", 17);

            //C39	LOWER_DIAG_SS_LENGTH
            SetGECParameter(oProject, electric, "C39", 18);

            //C40   INTERM1_DIAG_SS_LENGTH
            SetGECParameter(oProject, electric, "C40", 0);

            //C41   INTERM2_DIAG_SS_LENGTH
            SetGECParameter(oProject, electric, "C41", 0);

            //C42   DIAG3_ENABLE
            SetGECParameter(oProject, electric, "C42", (uint)GEC.Active.Disable);

            //C43   LABEL_FAULT1
            SetGECParameter(oProject, electric, "C43", 0);

            //C44   LABEL_FAULT2
            SetGECParameter(oProject, electric, "C44", 0);

            //C45   LABEL_FAULT3
            SetGECParameter(oProject, electric, "C45", 0);

            //C46   LABEL_FAULT4
            SetGECParameter(oProject, electric, "C46", 0);

            //C47   LABEL_FAULT5
            SetGECParameter(oProject, electric, "C47", 0);

            //C48   LABEL_FAULT6
            SetGECParameter(oProject, electric, "C48", 0);

            //C49   LABEL_FAULT7
            SetGECParameter(oProject, electric, "C49", 0);

            //C50   LABEL_FAULT8
            SetGECParameter(oProject, electric, "C50", 0);

            //C51   LABEL_FAULT9
            SetGECParameter(oProject, electric, "C51", 0);

            //C52   LABEL_FAULT10
            SetGECParameter(oProject, electric, "C52", 0);

            //C53   RGB_COLOR
            SetGECParameter(oProject, electric, "C53", 0);

            //C54   LIGHTING_1_MODE
            SetGECParameter(oProject, electric, "C54", (uint)GEC.Lighting.auto);

            //C55	LIGHTING_2_MODE
            SetGECParameter(oProject, electric, "C55", (uint)GEC.Lighting.auto);

            //C56	LIGHTING_3_MODE
            SetGECParameter(oProject, electric, "C56", (uint)GEC.Lighting.auto);

            //C57	COVER SIGNAL
            SetGECParameter(oProject, electric, "C57", 0);

            //C58	BELL TIME
            SetGECParameter(oProject, electric, "C58", 5);

            //C59   MAX ORDER
            SetGECParameter(oProject, electric, "C59", (uint)GEC.Active.Enable);

            //C60	MAX CAN ID
            SetGECParameter(oProject, electric, "C60", 0);

            //C61	SPARE_PARAMETER_19
            /*  No available    */

            //C62   SPARE_PARAMETER_20
            /*  No available    */

            //C63   CONTINUOS_KEY_FUNCTION
            SetGECParameter(oProject, electric, "C63", (uint)GEC.Mode.No_People_Detection);
            if (modofuncionamiento.CurrentReference.Equals("SGBV"))
                SetGECParameter(oProject, electric, "C63", (uint)GEC.Mode.Standby);

            //C64   AUTOMATIC_KEY_FUNCTION
            if (modofuncionamiento.CurrentReference.Equals("INTERM") ||
              modofuncionamiento.CurrentReference.Equals("SG") ||
              modofuncionamiento.CurrentReference.Equals("SGBV"))
            {
                if (sistAhorro.CurrentReference.Contains("VA"))
                    SetGECParameter(oProject, electric, "C64", (uint)GEC.Mode.Intermittent_Standby);
                else
                    SetGECParameter(oProject, electric, "C64", (uint)GEC.Mode.Intermittent);

            }

            //C65   BRAKE_WEAR_RUN
            SetGECParameter(oProject, electric, "C65", (uint)GEC.Active.Enable);

            //C66   OIL_PUMP_RUN
            SetGECParameter(oProject, electric, "C66", (uint)GEC.Active.Disable);
            
            //C67   OIL_GEARBOX_RUN
            SetGECParameter(oProject, electric, "C67", (uint)GEC.Active.Enable);

            //C68   OIL_PUMP2_TIMER_ON
            SetGECParameter(oProject, electric, "C68", Calc_OIL_PUMP1_TIMER_ON(electric));

            //C69   OIL_PUMP2_CYCLE_TIME
            SetGECParameter(oProject, electric, "C69", Calc_OIL_PUMP1_CYCLE_TIME(electric));

            //C70	PREDEFINE_MODE
            if (modofuncionamiento.CurrentReference.Equals("INTERM") ||
               modofuncionamiento.CurrentReference.Equals("SG") ||
               modofuncionamiento.CurrentReference.Equals("SGBV"))
            {
                if (sistAhorro.CurrentReference.Contains("VA"))
                    SetGECParameter(oProject, electric, "C70", (uint)GEC.Mode.Intermittent_Standby);
                else
                    SetGECParameter(oProject, electric, "C70", (uint)GEC.Mode.Intermittent);

            }
            else if (modofuncionamiento.CurrentReference.Equals("BV"))
                SetGECParameter(oProject, electric, "C70", (uint)GEC.Mode.Standby);

            //C89	DELAY_TIME_STOP_IF_FIRE_ALARM
            SetGECParameter(oProject, electric, "C89", 15);

            #endregion

            #region Modbus
            //C100 MODBUS000
            SetGECParameter(oProject, electric, "C100", 1);
            //C101 MODBUS001
            SetGECParameter(oProject, electric, "C101", 5);
            //C102 MODBUS002
            SetGECParameter(oProject, electric, "C102", 7);
            //C103 MODBUS003
            SetGECParameter(oProject, electric, "C103", 12);
            //C104 MODBUS004
            SetGECParameter(oProject, electric, "C104", 13);
            //C105 MODBUS005
            SetGECParameter(oProject, electric, "C105", 14);
            //C106 MODBUS006
            SetGECParameter(oProject, electric, "C106", 15);
            //C107 MODBUS007
            SetGECParameter(oProject, electric, "C107", 16);
            //C108 MODBUS008
            SetGECParameter(oProject, electric, "C108", 20);
            //C109 MODBUS009
            SetGECParameter(oProject, electric, "C109", 21);
            //C110 MODBUS010
            SetGECParameter(oProject, electric, "C110", 24);
            //C111 MODBUS011
            SetGECParameter(oProject, electric, "C111", 25);
            //C112 MODBUS012
            SetGECParameter(oProject, electric, "C112", 26);
            //C113 MODBUS013
            SetGECParameter(oProject, electric, "C113", 28);
            //C114 MODBUS014
            SetGECParameter(oProject, electric, "C114", 30);
            //C115 MODBUS015
            SetGECParameter(oProject, electric, "C115", 32);
            //C116 MODBUS016
            SetGECParameter(oProject, electric, "C116", 34);
            //C117 MODBUS017
            SetGECParameter(oProject, electric, "C117", 36);
            //C118 MODBUS018
            SetGECParameter(oProject, electric, "C118", 38);
            //C119 MODBUS019
            SetGECParameter(oProject, electric, "C119", 40);
            //C120 MODBUS020
            SetGECParameter(oProject, electric, "C120", 42);
            //C121 MODBUS021
            SetGECParameter(oProject, electric, "C121", 44);
            //C122 MODBUS022
            SetGECParameter(oProject, electric, "C122", 46);
            //C123 MODBUS023
            SetGECParameter(oProject, electric, "C123", 48);
            //C124 MODBUS024
            SetGECParameter(oProject, electric, "C124", 50);
            //C125 MODBUS025
            SetGECParameter(oProject, electric, "C125", 52);
            //C126 MODBUS026
            SetGECParameter(oProject, electric, "C126", 54);
            //C127 MODBUS027
            SetGECParameter(oProject, electric, "C127", 56);
            //C128 MODBUS028
            SetGECParameter(oProject, electric, "C128", 0);
            //C129 MODBUS029
            SetGECParameter(oProject, electric, "C129", 0);
            //C130 MODBUS030
            SetGECParameter(oProject, electric, "C130", 0);
            //C131 MODBUS031
            SetGECParameter(oProject, electric, "C131", 0);
            //C132 MODBUS032
            SetGECParameter(oProject, electric, "C132", 57);
            //C133 MODBUS033
            SetGECParameter(oProject, electric, "C133", 0);
            //C134 MODBUS034
            SetGECParameter(oProject, electric, "C134", 58);
            //C135 MODBUS035
            SetGECParameter(oProject, electric, "C135", 0);
            //C136 MODBUS036
            SetGECParameter(oProject, electric, "C136", 0);
            //C137 MODBUS037
            SetGECParameter(oProject, electric, "C137", 59);
            //C138 MODBUS038
            SetGECParameter(oProject, electric, "C138", 0);
            //C139 MODBUS039
            SetGECParameter(oProject, electric, "C139", 0);
            //C140 MODBUS TYPE
            SetGECParameter(oProject, electric, "C140", 2);
            //C141    LOG_ERASE
            SetGECParameter(oProject, electric, "C141", (uint)GEC.Active.Enable);
            //C142    UNIT ID
            SetGECParameter(oProject, electric, "C142", 1);
            //C143 RESERVED
            //C144 RUN REMOTE
            SetGECParameter(oProject, electric, "C144", (uint)GEC.Active.Enable);
            //C145 GENERAL 1 BIT 0
            SetGECParameter(oProject, electric, "C145", 1007);
            //C146 GENERAL 1 BIT 1
            SetGECParameter(oProject, electric, "C146", 1003);
            //C147 GENERAL 1 BIT 2
            SetGECParameter(oProject, electric, "C147", 1013);
            //C148 GENERAL 1 BIT 3
            SetGECParameter(oProject, electric, "C148", 1014);
            //C149 GENERAL 1 BIT 4
            SetGECParameter(oProject, electric, "C149", 48);
            //C150 GENERAL 1 BIT 5
            SetGECParameter(oProject, electric, "C150", 244);
            //C151 GENERAL 1 BIT 6
            SetGECParameter(oProject, electric, "C151", 1009);
            //C152 GENERAL 1 BIT 7
            SetGECParameter(oProject, electric, "C152", 1006);
            //C153 GENERAL 1 BIT 8
            SetGECParameter(oProject, electric, "C153", 1015);
            //C154 GENERAL 1 BIT 9
            SetGECParameter(oProject, electric, "C154", 324);
            //C155 GENERAL 1 BIT 10
            SetGECParameter(oProject, electric, "C155", 483);
            //C156 GENERAL 1 BIT 11
            SetGECParameter(oProject, electric, "C156", 9);
            //C157 GENERAL 1 BIT 12
            SetGECParameter(oProject, electric, "C157", 67);
            //C158 GENERAL 1 BIT 13
            SetGECParameter(oProject, electric, "C158", 4);
            //C159 GENERAL 1 BIT 14
            SetGECParameter(oProject, electric, "C159", 26);
            //C160 GENERAL 1 BIT 15
            SetGECParameter(oProject, electric, "C160", 28);
            //C161 GENERAL 1 BIT 16
            SetGECParameter(oProject, electric, "C161", 259);
            //C162 GENERAL 1 BIT 17
            SetGECParameter(oProject, electric, "C162", 258);
            //C163 GENERAL 1 BIT 18
            SetGECParameter(oProject, electric, "C163", 328);
            //C164 GENERAL 1 BIT 19
            SetGECParameter(oProject, electric, "C164", 23);
            //C165 GENERAL 1 BIT 20
            SetGECParameter(oProject, electric, "C165", 24);
            //C166 GENERAL 1 BIT 21
            SetGECParameter(oProject, electric, "C166", 40);
            //C167 GENERAL 1 BIT 22
            SetGECParameter(oProject, electric, "C167", 41);
            //C168 GENERAL 1 BIT 23
            SetGECParameter(oProject, electric, "C168", 32);
            //C169 GENERAL 1 BIT 24
            SetGECParameter(oProject, electric, "C169", 406);
            //C170 GENERAL 1 BIT 25
            SetGECParameter(oProject, electric, "C170", 33);
            //C171 GENERAL 1 BIT 26
            SetGECParameter(oProject, electric, "C171", 434);
            //C172 GENERAL 1 BIT 27
            SetGECParameter(oProject, electric, "C172", 21);
            //C173 GENERAL 1 BIT 28
            SetGECParameter(oProject, electric, "C173", 22);
            //C174 GENERAL 1 BIT 29
            SetGECParameter(oProject, electric, "C174", 95);
            //C175 GENERAL 1 BIT 30
            SetGECParameter(oProject, electric, "C175", 96);
            //C176 GENERAL 1 BIT 31
            SetGECParameter(oProject, electric, "C176", 38);
            //C177 GENERAL 1 BIT 32
            SetGECParameter(oProject, electric, "C177", 39);
            //C178 GENERAL 1 BIT 33
            SetGECParameter(oProject, electric, "C178", 97);
            //C179 GENERAL 1 BIT 34
            SetGECParameter(oProject, electric, "C179", 98);
            //C180 GENERAL 1 BIT 35
            SetGECParameter(oProject, electric, "C180", 44);
            //C181 GENERAL 1 BIT 36
            SetGECParameter(oProject, electric, "C181", 45);
            //C182 GENERAL 1 BIT 37
            SetGECParameter(oProject, electric, "C182", 50);
            //C183 GENERAL 1 BIT 38
            SetGECParameter(oProject, electric, "C183", 51);
            //C184 GENERAL 1 BIT 39
            SetGECParameter(oProject, electric, "C184", 46);
            //C185 GENERAL 1 BIT 40
            SetGECParameter(oProject, electric, "C185", 47);
            //C186 GENERAL 1 BIT 41
            SetGECParameter(oProject, electric, "C186", 35);
            //C187 GENERAL 1 BIT 42
            SetGECParameter(oProject, electric, "C187", 36);
            //C188 GENERAL 1 BIT 43
            SetGECParameter(oProject, electric, "C188", 11);
            //C189 GENERAL 1 BIT 44
            SetGECParameter(oProject, electric, "C189", 12);
            //C190 GENERAL 1 BIT 45
            SetGECParameter(oProject, electric, "C190", 286);
            //C191 GENERAL 1 BIT 46
            SetGECParameter(oProject, electric, "C191", 0);
            //C192 GENERAL 1 BIT 47
            SetGECParameter(oProject, electric, "C192", 0);
            //C193 GENERAL 1 BIT 48
            SetGECParameter(oProject, electric, "C193", 0);
            //C194 GENERAL 1 BIT 49
            SetGECParameter(oProject, electric, "C194", 0);
            //C195 GENERAL 1 BIT 50
            SetGECParameter(oProject, electric, "C195", 0);
            //C196 GENERAL 1 BIT 51
            SetGECParameter(oProject, electric, "C196", 0);
            //C197 GENERAL 1 BIT 52
            SetGECParameter(oProject, electric, "C197", 0);
            //C198 GENERAL 1 BIT 53
            SetGECParameter(oProject, electric, "C198", 0);
            //C199 GENERAL 1 BIT 54
            SetGECParameter(oProject, electric, "C199", 0);
            //C200 GENERAL 1 BIT 55
            SetGECParameter(oProject, electric, "C200", 0);
            //C201 GENERAL 1 BIT 56
            SetGECParameter(oProject, electric, "C201", 0);
            //C202 GENERAL 1 BIT 57
            SetGECParameter(oProject, electric, "C202", 0);
            //C203 GENERAL 1 BIT 58
            SetGECParameter(oProject, electric, "C203", 0);
            //C204 GENERAL 1 BIT 59
            SetGECParameter(oProject, electric, "C204", 0);
            //C205 GENERAL 1 BIT 60
            SetGECParameter(oProject, electric, "C205", 0);
            //C206 GENERAL 1 BIT 61
            SetGECParameter(oProject, electric, "C206", 0);
            //C207 GENERAL 1 BIT 62
            SetGECParameter(oProject, electric, "C207", 0);
            //C208 GENERAL 1 BIT 63
            SetGECParameter(oProject, electric, "C208", 2014);
            //C209 GENERAL 2 BIT 0
            SetGECParameter(oProject, electric, "C209", 2014);
            //C210 GENERAL 2 BIT 1
            SetGECParameter(oProject, electric, "C210", 2030);
            //C211 GENERAL 2 BIT 2
            SetGECParameter(oProject, electric, "C211", 2031);
            //C212 GENERAL 2 BIT 3
            SetGECParameter(oProject, electric, "C212", 2005);
            //C213 GENERAL 2 BIT 4
            SetGECParameter(oProject, electric, "C213", 2004);
            //C214 GENERAL 2 BIT 5
            SetGECParameter(oProject, electric, "C214", 2000);
            //C215 GENERAL 2 BIT 6
            SetGECParameter(oProject, electric, "C215", 2001);
            //C216 GENERAL 2 BIT 7
            SetGECParameter(oProject, electric, "C216", 2015);
            //C217 GENERAL 2 BIT 8
            SetGECParameter(oProject, electric, "C217", 2024);
            //C218 GENERAL 2 BIT 9
            SetGECParameter(oProject, electric, "C218", 2016);
            //C219 GENERAL 2 BIT 10
            SetGECParameter(oProject, electric, "C219", 2003);
            //C220 GENERAL 2 BIT 11
            SetGECParameter(oProject, electric, "C220", 2032);
            //C221 GENERAL 2 BIT 12
            SetGECParameter(oProject, electric, "C221", 2033);
            //C222 GENERAL 2 BIT 13
            SetGECParameter(oProject, electric, "C222", 2022);
            //C223 GENERAL 2 BIT 14
            SetGECParameter(oProject, electric, "C223", 2029);
            //C224 GENERAL 2 BIT 15
            SetGECParameter(oProject, electric, "C224", 2026);
            //C225 GENERAL 2 BIT 16
            SetGECParameter(oProject, electric, "C225", 2027);
            //C226 GENERAL 2 BIT 17
            SetGECParameter(oProject, electric, "C226", 2011);
            //C227 GENERAL 2 BIT 18
            SetGECParameter(oProject, electric, "C227", 2012);
            //C228 GENERAL 2 BIT 19
            SetGECParameter(oProject, electric, "C228", 2013);
            //C229 GENERAL 2 BIT 20
            SetGECParameter(oProject, electric, "C229", 2036);
            //C230 GENERAL 2 BIT 21
            SetGECParameter(oProject, electric, "C230", 2017);
            //C231 GENERAL 2 BIT 22
            SetGECParameter(oProject, electric, "C231", 2018);
            //C232 GENERAL 2 BIT 23
            SetGECParameter(oProject, electric, "C232", 2019);
            //C233 GENERAL 2 BIT 24
            SetGECParameter(oProject, electric, "C233", 2025);
            //C234 GENERAL 2 BIT 25
            SetGECParameter(oProject, electric, "C234", 2014);
            //C235 GENERAL 2 BIT 26
            SetGECParameter(oProject, electric, "C235", 2039);
            //C236 GENERAL 2 BIT 27
            SetGECParameter(oProject, electric, "C236", 0);
            //C237 GENERAL 2 BIT 28
            SetGECParameter(oProject, electric, "C237", 0);
            //C238 GENERAL 2 BIT 29
            SetGECParameter(oProject, electric, "C238", 0);
            //C239 GENERAL 2 BIT 30
            SetGECParameter(oProject, electric, "C239", 0);
            //C240 GENERAL 2 BIT 31
            SetGECParameter(oProject, electric, "C240", 0);
            //C241 ORDER BIT 0
            SetGECParameter(oProject, electric, "C241", 3);
            //C242 ORDER  BIT 1
            SetGECParameter(oProject, electric, "C242", 4);
            //C243 ORDER BIT 2
            SetGECParameter(oProject, electric, "C243", 5);
            //C244 ORDER  BIT 3
            SetGECParameter(oProject, electric, "C244", 6);
            //C245 ORDER  BIT 4
            SetGECParameter(oProject, electric, "C245", 7);
            //C246 ORDER  BIT 5
            SetGECParameter(oProject, electric, "C246", 8);
            //C247 ORDER  BIT 6
            SetGECParameter(oProject, electric, "C247", 9);
            //C248 ORDER  BIT 7
            SetGECParameter(oProject, electric, "C248", 10);
            //C249 ORDER  BIT 8
            SetGECParameter(oProject, electric, "C249", 11);
            //C250 ORDER  BIT 9
            SetGECParameter(oProject, electric, "C250", 12);
            //C251 ORDER  BIT 10
            SetGECParameter(oProject, electric, "C251", 13);
            //C252 ORDER  BIT 11
            SetGECParameter(oProject, electric, "C252", 14);
            //C253 ORDER  BIT 12
            SetGECParameter(oProject, electric, "C253", 0);
            //C254 ORDER  BIT 13
            SetGECParameter(oProject, electric, "C254", 0);
            //C255 ORDER  BIT 14
            SetGECParameter(oProject, electric, "C255", 0);
            //C256 ORDER  BIT 15
            SetGECParameter(oProject, electric, "C256", 0);
            //C257 GENERAL 3 BIT 0
            SetGECParameter(oProject, electric, "C257", 3008);
            //C258 GENERAL 3 BIT 1
            SetGECParameter(oProject, electric, "C258", 0);
            //C259 GENERAL 3 BIT 2
            SetGECParameter(oProject, electric, "C259", 0);
            //C260 GENERAL 3 BIT 3
            SetGECParameter(oProject, electric, "C260", 0);
            //C261 GENERAL 3 BIT 4
            SetGECParameter(oProject, electric, "C261", 3023);
            //C262 GENERAL 3 BIT 5
            SetGECParameter(oProject, electric, "C262", 3024);
            //C263 GENERAL 3 BIT 6
            SetGECParameter(oProject, electric, "C263", 0);
            //C264 GENERAL 3 BIT 7
            SetGECParameter(oProject, electric, "C264", 0);
            //C265 GENERAL 3 BIT 8
            SetGECParameter(oProject, electric, "C265", 3025);
            //C266 GENERAL 3 BIT 9
            SetGECParameter(oProject, electric, "C266", 3026);
            //C267 GENERAL 3 BIT 10
            SetGECParameter(oProject, electric, "C267", 0);
            //C268 GENERAL 3 BIT 11
            SetGECParameter(oProject, electric, "C268", 0);
            //C269 GENERAL 3 BIT 12
            SetGECParameter(oProject, electric, "C269", 0);
            //C270 GENERAL 3 BIT 13
            SetGECParameter(oProject, electric, "C270", 0);
            //C271 GENERAL 3 BIT 14
            SetGECParameter(oProject, electric, "C271", 0);
            //C272 GENERAL 3 BIT 15
            SetGECParameter(oProject, electric, "C272", 0);

            #endregion

            #region Control Outputs
            //O12	C NO2 K2.1
            //K2.1
            if (motorConnection.CurrentReference.Equals("YD") ||
            motorConnection.CurrentReference.Equals("VVF_YD"))
                SetGECParameter(oProject, electric, "O12", (uint)GEC.Param.Star);
            else
                SetGECParameter(oProject, electric, "O12", (uint)GEC.Param.Empty);

            //O13	C NO1 K2.2
            //K2.2
            if (motorConnection.CurrentReference.Equals("YD") ||
                motorConnection.CurrentReference.Equals("VVF_YD"))
                SetGECParameter(oProject, electric, "O13", (uint)GEC.Param.Delta);
            else
                SetGECParameter(oProject, electric, "O13", (uint)GEC.Param.Empty);


            #endregion

            #region Upper Diagnostic Outputs
            //UO4 UDL1 Relay output 1 NO Q4/ 4L
            SetGECParameter(oProject, electric, "UO4", (uint)GEC.Param.Empty);
            //UO5 UDL1 Relay output 2NO Q5/ 5L
            SetGECParameter(oProject, electric, "UO5", (uint)GEC.Param.Empty);
            //UO6 UDL1 Relay output 3 NO Q6/ 6L
            SetGECParameter(oProject, electric, "UO6", (uint)GEC.Param.Empty);
            //UO7 UDL1 Relay output 4 NO Q7/ 7L
            SetGECParameter(oProject, electric, "UO7", (uint)GEC.Param.Empty);
            //UO8 UDL1 Relay output 5 NO Q8/ 8L
            SetGECParameter(oProject, electric, "UO8", (uint)GEC.Param.Empty);
            #endregion

            #region Control Inputs

            #region Control Board Inputs
            //I1 C Standard input 1 X01
            SetGECParameter(oProject, electric, "I1", (uint)GEC.Param.Asymmetry_phase_control_relay);

            //I2  C Standard input 2
            SetGECParameter(oProject, electric, "I2", (uint)GEC.Param.Overtemperature_M1);

            //I3 C Standard input 3
            SetGECParameter(oProject, electric, "I3", (uint)GEC.Param.Protection_switch_drive_M1);

            #endregion

            #region Upper Diagnostic Inputs
            #endregion

            #region Lower Diagnostic Inputs

            #endregion

            #endregion

            #endregion


        }

        public uint calcPPR(uint lcinta, uint dpolos)
        {
            // Funcion para calcular pulsos por revolucion (PPR)
            //      Entradas:
            //          tapelenght: longitud cinta magnetica
            //          dpolos:  distancia entre polos

            uint npolos_std = 174;   // Numero de polos para longitud de cinta standard
            uint ppr_std = 522;      // Periodos de onda cuadrada por revolucion con cinta de 174 polos

            uint npolos = lcinta / dpolos;               // Numero de polos
            uint ppr = npolos * ppr_std / npolos_std;    // Periodos de onda cuadrada por vuelta

            return ppr;
        }
        
        public double calcRPH(string m, string p, double v)
        {
            int z = 0;

            switch (m)
            {
                case ("TUGELA"):
                    if (p == "135")
                        z = 16;
                    else if (p == "101,25")
                        z = 22;
                    else
                    {
                        p = "0";
                    }
                    break;
                case ("VELINO"):
                    if (p == "135")
                        z = 16;
                    else if (p == "101,25")
                        z = 22;
                    else
                    {
                        p = "0";
                    }
                    break;
                case ("ORINOCO"):
                    if (p == "135" && (v.ToString() == "0,5" || v.ToString() == "0,65"))
                        z = 16;
                    else
                    {
                        p = "0";
                    }
                    break;
                case ("VICTORIA"):
                    if (p == "101,15")
                        z = 23;
                    else
                    {
                        p = "0";
                    }
                    break;
                case ("VELINO_CLASSIC"):
                case ("TUGELA_CLASSIC"):
                    if (p == "135")
                        z = 16;
                    else if (p == "101,25")
                        z = 22;
                    else
                    {
                        p = "0";
                    }
                    break;
            }

            //float primdiam = Convert.ToSingle(Convert.ToSingle(p) / (Math.Sin(Math.PI / z))); // Primitive diameter
            //float rph = (float)1.2 * Convert.ToSingle(v / (primdiam / (2 * 1000)) / (2 * Math.PI) * 60 * 60); // Main shaft rph
            float rph = (float)1.2 * Convert.ToSingle(v) / Convert.ToSingle(z) / Convert.ToSingle(p) * 3600000;

            return rph;
        }

        public uint Calc_OIL_PUMP1_TIMER_ON(Electric electric)
        {
            Caracteristic Velocidad = (Caracteristic)electric.CaractComercial["FGESCHW"];
            Caracteristic Producto = (Caracteristic)electric.CaractComercial["FMODELL"];
            Caracteristic Inclinacion = (Caracteristic)electric.CaractComercial["FNEIGUNG"];
            Caracteristic desnivel = (Caracteristic)electric.CaractComercial["FHOEHEV"];

            uint res = 5;

            if (Producto.CurrentReference.Equals("VELINO") ||
                Producto.CurrentReference.Equals("TUGELA") ||
                Producto.CurrentReference.Contains("CLASSIC"))

            {
                switch (desnivel.NumVal)
                {
                    #region Item 1
                    case double n when (n <= 2000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 42;
                                        break;

                                    case 0.65:
                                        res = 35;
                                        break;

                                    case 0.75:
                                        res = 31;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 40;
                                        break;

                                    case 0.65:
                                        res = 33;
                                        break;

                                    case 0.75:
                                        res = 30;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 38;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 2
                    case double n when (n <= 2500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 45;
                                        break;

                                    case 0.65:
                                        res = 38;
                                        break;

                                    case 0.75:
                                        res = 34;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 44;
                                        break;

                                    case 0.65:
                                        res = 36;
                                        break;

                                    case 0.75:
                                        res = 32;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 41;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 3
                    case double n when (n <= 3000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 50;
                                        break;

                                    case 0.65:
                                        res = 41;
                                        break;

                                    case 0.75:
                                        res = 37;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 48;
                                        break;

                                    case 0.65:
                                        res = 39;
                                        break;

                                    case 0.75:
                                        res = 35;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 44;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 4
                    case double n when (n <= 3500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 54;
                                        break;

                                    case 0.65:
                                        res = 44;
                                        break;

                                    case 0.75:
                                        res = 39;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 52;
                                        break;

                                    case 0.65:
                                        res = 42;
                                        break;

                                    case 0.75:
                                        res = 38;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 48;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 5
                    case double n when (n <= 4000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 58;
                                        break;

                                    case 0.65:
                                        res = 47;
                                        break;

                                    case 0.75:
                                        res = 42;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 56;
                                        break;

                                    case 0.65:
                                        res = 45;
                                        break;

                                    case 0.75:
                                        res = 40;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 52;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 6
                    case double n when (n <= 4500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 61;
                                        break;

                                    case 0.65:
                                        res = 51;
                                        break;

                                    case 0.75:
                                        res = 45;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 60;
                                        break;

                                    case 0.65:
                                        res = 48;
                                        break;

                                    case 0.75:
                                        res = 43;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 55;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 7
                    case double n when (n <= 5000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 66;
                                        break;

                                    case 0.65:
                                        res = 54;
                                        break;

                                    case 0.75:
                                        res = 48;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 64;
                                        break;

                                    case 0.65:
                                        res = 51;
                                        break;

                                    case 0.75:
                                        res = 46;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 59;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 8
                    case double n when (n <= 5500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 70;
                                        break;

                                    case 0.65:
                                        res = 57;
                                        break;

                                    case 0.75:
                                        res = 51;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 68;
                                        break;

                                    case 0.65:
                                        res = 54;
                                        break;

                                    case 0.75:
                                        res = 48;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 62;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 9
                    case double n when (n <= 5800):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 72;
                                        break;

                                    case 0.65:
                                        res = 59;
                                        break;

                                    case 0.75:
                                        res = 53;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 70;
                                        break;

                                    case 0.65:
                                        res = 56;
                                        break;

                                    case 0.75:
                                        res = 50;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 65;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 10
                    case double n when (n <= 6000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 77;
                                        break;

                                    case 0.65:
                                        res = 63;
                                        break;

                                    case 0.75:
                                        res = 56;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 76;
                                        break;

                                    case 0.65:
                                        res = 60;
                                        break;

                                    case 0.75:
                                        res = 53;
                                        break;
                                }
                                break;

                            case 35:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 67;
                                        break;

                                    case 0.65:
                                        break;

                                    case 0.75:
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 11
                    case double n when (n <= 6500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 84;
                                        break;

                                    case 0.65:
                                        res = 67;
                                        break;

                                    case 0.75:
                                        res = 59;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 80;
                                        break;

                                    case 0.65:
                                        res = 63;
                                        break;

                                    case 0.75:
                                        res = 56;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 12
                    case double n when (n <= 7000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 88;
                                        break;

                                    case 0.65:
                                        res = 70;
                                        break;

                                    case 0.75:
                                        res = 62;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 84;
                                        break;

                                    case 0.65:
                                        res = 66;
                                        break;

                                    case 0.75:
                                        res = 59;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 13
                    case double n when (n <= 7500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 92;
                                        break;

                                    case 0.65:
                                        res = 73;
                                        break;

                                    case 0.75:
                                        res = 65;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 88;
                                        break;

                                    case 0.65:
                                        res = 69;
                                        break;

                                    case 0.75:
                                        res = 61;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 14
                    case double n when (n <= 8000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 97;
                                        break;

                                    case 0.65:
                                        res = 77;
                                        break;

                                    case 0.75:
                                        res = 68;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 92;
                                        break;

                                    case 0.65:
                                        res = 72;
                                        break;

                                    case 0.75:
                                        res = 64;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 15
                    case double n when (n <= 8500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 101;
                                        break;

                                    case 0.65:
                                        res = 80;
                                        break;

                                    case 0.75:
                                        res = 71;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 96;
                                        break;

                                    case 0.65:
                                        res = 75;
                                        break;

                                    case 0.75:
                                        res = 67;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 16
                    case double n when (n <= 9000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 105;
                                        break;

                                    case 0.65:
                                        res = 83;
                                        break;

                                    case 0.75:
                                        res = 74;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 100;
                                        break;

                                    case 0.65:
                                        res = 78;
                                        break;

                                    case 0.75:
                                        res = 69;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 17
                    case double n when (n <= 9500):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 110;
                                        break;

                                    case 0.65:
                                        res = 87;
                                        break;

                                    case 0.75:
                                        res = 77;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 103;
                                        break;

                                    case 0.65:
                                        res = 81;
                                        break;

                                    case 0.75:
                                        res = 72;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 18
                    case double n when (n <= 10000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 114;
                                        break;

                                    case 0.65:
                                        res = 90;
                                        break;

                                    case 0.75:
                                        res = 79;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 108;
                                        break;

                                    case 0.65:
                                        res = 85;
                                        break;

                                    case 0.75:
                                        res = 75;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 19
                    case double n when (n <= 11000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 123;
                                        break;

                                    case 0.65:
                                        res = 97;
                                        break;

                                    case 0.75:
                                        res = 85;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 115;
                                        break;

                                    case 0.65:
                                        res = 91;
                                        break;

                                    case 0.75:
                                        res = 80;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 20
                    case double n when (n <= 12000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 132;
                                        break;

                                    case 0.65:
                                        res = 104;
                                        break;

                                    case 0.75:
                                        res = 91;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 123;
                                        break;

                                    case 0.65:
                                        res = 97;
                                        break;

                                    case 0.75:
                                        res = 85;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 21
                    case double n when (n <= 13000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 140;
                                        break;

                                    case 0.65:
                                        res = 110;
                                        break;

                                    case 0.75:
                                        res = 97;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 132;
                                        break;

                                    case 0.65:
                                        res = 103;
                                        break;

                                    case 0.75:
                                        res = 91;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 22
                    case double n when (n <= 14000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 149;
                                        break;

                                    case 0.65:
                                        res = 117;
                                        break;

                                    case 0.75:
                                        res = 103;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 140;
                                        break;

                                    case 0.65:
                                        res = 109;
                                        break;

                                    case 0.75:
                                        res = 96;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 23
                    case double n when (n <= 15000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 158;
                                        break;

                                    case 0.65:
                                        res = 124;
                                        break;

                                    case 0.75:
                                        res = 108;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 148;
                                        break;

                                    case 0.65:
                                        res = 115;
                                        break;

                                    case 0.75:
                                        res = 101;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 24
                    case double n when (n <= 16000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 166;
                                        break;

                                    case 0.65:
                                        res = 130;
                                        break;

                                    case 0.75:
                                        res = 114;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 155;
                                        break;

                                    case 0.65:
                                        res = 121;
                                        break;

                                    case 0.75:
                                        res = 107;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 25
                    case double n when (n <= 17000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 175;
                                        break;

                                    case 0.65:
                                        res = 137;
                                        break;

                                    case 0.75:
                                        res = 120;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 163;
                                        break;

                                    case 0.65:
                                        res = 128;
                                        break;

                                    case 0.75:
                                        res = 112;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 26
                    case double n when (n <= 18000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 184;
                                        break;

                                    case 0.65:
                                        res = 144;
                                        break;

                                    case 0.75:
                                        res = 126;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 171;
                                        break;

                                    case 0.65:
                                        res = 134;
                                        break;

                                    case 0.75:
                                        res = 117;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 27
                    case double n when (n <= 19000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 193;
                                        break;

                                    case 0.65:
                                        res = 150;
                                        break;

                                    case 0.75:
                                        res = 132;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 180;
                                        break;

                                    case 0.65:
                                        res = 140;
                                        break;

                                    case 0.75:
                                        res = 123;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 28
                    case double n when (n <= 20000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 201;
                                        break;

                                    case 0.65:
                                        res = 157;
                                        break;

                                    case 0.75:
                                        res = 138;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 187;
                                        break;

                                    case 0.65:
                                        res = 146;
                                        break;

                                    case 0.75:
                                        res = 128;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 29
                    case double n when (n <= 21000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 210;
                                        break;

                                    case 0.65:
                                        res = 164;
                                        break;

                                    case 0.75:
                                        res = 143;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 195;
                                        break;

                                    case 0.65:
                                        res = 152;
                                        break;

                                    case 0.75:
                                        res = 133;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 30
                    case double n when (n <= 22000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 219;
                                        break;

                                    case 0.65:
                                        res = 171;
                                        break;

                                    case 0.75:
                                        res = 149;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 204;
                                        break;

                                    case 0.65:
                                        res = 158;
                                        break;

                                    case 0.75:
                                        res = 139;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 31
                    case double n when (n <= 23000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 227;
                                        break;

                                    case 0.65:
                                        res = 177;
                                        break;

                                    case 0.75:
                                        res = 155;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 212;
                                        break;

                                    case 0.65:
                                        res = 165;
                                        break;

                                    case 0.75:
                                        res = 144;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 32
                    case double n when (n <= 24000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 236;
                                        break;

                                    case 0.65:
                                        res = 184;
                                        break;

                                    case 0.75:
                                        res = 161;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 220;
                                        break;

                                    case 0.65:
                                        res = 171;
                                        break;

                                    case 0.75:
                                        res = 149;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 33
                    case double n when (n <= 25000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 245;
                                        break;

                                    case 0.65:
                                        res = 191;
                                        break;

                                    case 0.75:
                                        res = 167;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 228;
                                        break;

                                    case 0.65:
                                        res = 177;
                                        break;

                                    case 0.75:
                                        res = 155;
                                        break;
                                }
                                break;
                        }
                        break;
                    #endregion

                    #region Item 34
                    case double n when (n <= 26000):
                        switch (Inclinacion.NumVal)
                        {
                            case 27.3:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 254;
                                        break;

                                    case 0.65:
                                        res = 197;
                                        break;

                                    case 0.75:
                                        res = 172;
                                        break;
                                }
                                break;

                            case 30:
                                switch (Velocidad.NumVal)
                                {
                                    case 0.5:
                                        res = 236;
                                        break;

                                    case 0.65:
                                        res = 183;
                                        break;

                                    case 0.75:
                                        res = 160;
                                        break;
                                }
                                break;
                        }
                        break;
                        #endregion
                }
            }
            return res;
        }

        public uint Calc_OIL_PUMP1_CYCLE_TIME(Electric electric)
        {
            Caracteristic Velocidad = (Caracteristic)electric.CaractComercial["FGESCHW"];
            Caracteristic condClimatica = (Caracteristic)electric.CaractComercial["FKLIMAKLS"];

            uint res = 2;

            if (condClimatica.CurrentReference.Equals("1") ||
                condClimatica.CurrentReference.Equals("2"))
            {
                switch (Velocidad.NumVal)
                {
                    case 0.5:
                        res = 36;
                        break;

                    case 0.65:
                        res = 26;
                        break;

                    case 0.75:
                        res = 18;
                        break;
                }
            }
            else
            {
                res = 12;
            }

            return res;
        }

        public void paramGEC(Project oProject, Electric oElectric)
        {
            String path = String.Concat(oProject.DocumentDirectory.Substring(0, oProject.DocumentDirectory.Length - 3), "GEC\\GEC.csv");
            String data = "";
            foreach (GEC gEC in oElectric.GECParameterList.Values)
            {
                if (gEC.active)
                    data = String.Concat(data, "GEC_", gEC.ID, ";", gEC.getValue(), "\r\n");
            }
            File.WriteAllText(path, data);
        }

        public void writeGECtoEPLAN(Project oProject, Electric oElectric)
        {
            foreach (GEC gec in oElectric.GECParameterList.Values)
            {
                try
                {
                    string controlName = "GEC.CONTROL."+ gec.ID;
                    string safetyName = "GEC.CONTROL." + gec.ID;
                    AnyPropertyId propertyId = new AnyPropertyId(oProject, controlName);
                    if (propertyId == null)
                        propertyId= new AnyPropertyId(oProject, safetyName);
                    PropertyValue propertyValue = oProject.Properties[propertyId];

                    if (!gec.isNumeric)
                    {
                        MultiLangString langString = new MultiLangString();
                        langString.AddString(ISOCode.Language.L_en_US, gec.sValue);
                        propertyValue.Set(langString);
                    }
                    else
                    {
                        CultureInfo esES = CultureInfo.CreateSpecificCulture("es-ES");
                        MultiLangString value = new MultiLangString();
                        value.AddString(ISOCode.Language.L_en_US, String.Format(esES, "{0:0}", gec.value));
                        propertyValue.Set(value);
                    }
                }
                catch (Exception ex)
                {
                    ;
                }
            }

        }

        public void readGECfromEPLAN(Project oProject, Electric oElectric)
        {
            foreach (GEC gec in oElectric.GECParameterList.Values)
            {
                try
                {
                    string controlName = "GEC.CONTROL." + gec.ID;
                    string safetyName = "GEC.CONTROL." + gec.ID;
                    AnyPropertyId propertyId = new AnyPropertyId(oProject, controlName);
                    if (propertyId == null)
                        propertyId = new AnyPropertyId(oProject, safetyName);
                    PropertyValue propertyValue = oProject.Properties[propertyId];

                    string value = propertyValue.ToMultiLangString().GetString(ISOCode.Language.L_en_US);

                    if (int.TryParse(value, out int result))
                    {
                        gec.setValue((uint)result);
                    }
                    else
                    {
                        gec.setValue(value);
                    }

                }
                catch (Exception ex)
                {
                    ;
                }

            }
        }

        #endregion

        public void deleteAllDummyConnections(Project oProject) 
        {
            FunctionsFilter functionsFilter = new FunctionsFilter();
            functionsFilter.FunctionCategory = Eplan.EplApi.Base.Enums.FunctionCategory.DeviceEndTerminal;
            FunctionPropertyList functionsPropertyList = new FunctionPropertyList();
            functionsPropertyList.FUNC_VISIBLENAME = "";
            functionsFilter.SetFilteredPropertyList(functionsPropertyList);

            DMObjectsFinder DMObjectsFinder = new DMObjectsFinder(oProject);

            Function[] functions = DMObjectsFinder.GetFunctions(functionsFilter);
            foreach (Function function in functions)
            {
                function.Remove();
            }
            ;
        }
        #endregion
    }
}
