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

namespace EPLAN_API.User
{
    public class DrawTools
    {

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

            //Compruebo si ya esta insertada la página "VVF Power"
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
            page.Filter.Category = Function.Enums.Category.PLCTerminal;

            //now we have all functions having category 'MOTOR' placed on page p
            Function[] functions = page.Functions;

            //other way to do the same:
            FunctionsFilter ff = new FunctionsFilter();
            ff.Category = Function.Enums.Category.PLCTerminal;
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

            page.Filter.Category = Function.Enums.Category.PLCTerminal;

            //now we have all functions having category 'MOTOR' placed on page p
            Function[] functions = page.Functions;

            //other way to do the same:
            FunctionsFilter ff = new FunctionsFilter();
            ff.Category = Function.Enums.Category.PLCTerminal;
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
            ff.Category = Function.Enums.Category.PLCTerminal;
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
