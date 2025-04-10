using Eplan.EplAddin.Actions;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using EPLAN_API.API_Basic;
using EPLAN_API.User;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API_2024
{
    class Generate_Doc_Service : IActionService
    {
        public Generate_Doc_Service()
        {
        }
        public void Execute()
        {


            try
            {
                string pagestr;
                Project oProject = new ProjectManager().CurrentProject;
                Reports reports = new Reports();

                ProjectSettings projectSettings = new ProjectSettings(oProject);
                string sProjectsPath = projectSettings.GetStringSetting("FormGeneratorGui.PxfForm_INTERCONNECTDIAGRAM.FormName", 0);


                //Part List
                var filteredPagesMainPartList = oProject.Pages.Where(p => p.PageType == DocumentTypeManager.DocumentType.PartsList);

                foreach (Page page in filteredPagesMainPartList)
                {
                    page.Remove();
                }

                var filteredPagesMain = oProject.Pages.Where(p => p.NameParts.DESIGNATION_LOCATION.ToString().Equals("MAIN", StringComparison.OrdinalIgnoreCase)).ToArray();

                ReportBlock oReportPartList = new ReportBlock();
                oReportPartList.Create(oProject);
                oReportPartList.FormName = projectSettings.GetStringSetting("FormGeneratorGui.PxfForm_PARTSLIST.FormName", 0);
                oReportPartList.Type = DocumentTypeManager.DocumentType.PartsList;
                oReportPartList.FilterSchemaName = "MAIN";
                oReportPartList.IsAutomaticPageDescription = true;
                pagestr = filteredPagesMain.Last().Name.Split('/')[0] + "/" + (Convert.ToInt32(filteredPagesMain.Last().Name.Split('/')[1]) + 1).ToString();
                reports.CreateReport(oReportPartList, pagestr);

                //Cable Overview
                var filteredPagesCableOverview = oProject.Pages.Where(p => p.PageType == DocumentTypeManager.DocumentType.CableOverview);
                foreach (Page page in filteredPagesCableOverview)
                {
                    page.Remove();
                }
                ReportBlock oReportCableOverview = new ReportBlock();
                oReportCableOverview.Create(oProject);
                oReportCableOverview.FormName = projectSettings.GetStringSetting("FormGeneratorGui.PxfForm_CABLEOVERVIEW.FormName", 0);
                oReportCableOverview.Type = DocumentTypeManager.DocumentType.CableOverview;
                oReportCableOverview.IsAutomaticPageDescription = true;
                reports.CreateReport(oReportCableOverview, "=DOC+CBL/1");

                //Cable diagram
                var filteredPagesCableDiagram = oProject.Pages.Where(p => p.PageType == DocumentTypeManager.DocumentType.InterconnectDiagram);
                foreach (Page page in filteredPagesCableDiagram)
                {
                    page.Remove();
                }

                filteredPagesCableOverview = oProject.Pages.Where(p => p.PageType == DocumentTypeManager.DocumentType.CableOverview);
                
                ReportBlock oReportCableDigram = new ReportBlock();
                oReportCableDigram.Create(oProject);
                oReportCableDigram.FormName = projectSettings.GetStringSetting("FormGeneratorGui.PxfForm_INTERCONNECTDIAGRAM.FormName", 0);
                oReportCableDigram.Type = DocumentTypeManager.DocumentType.InterconnectDiagram;
                oReportCableDigram.IsAutomaticPageDescription = true;
                pagestr = filteredPagesCableOverview.Last().Name.Split('/')[0] + "/" + (Convert.ToInt32(filteredPagesCableOverview.Last().Name.Split('/')[1]) + 1).ToString();
                reports.CreateReport(oReportCableDigram, pagestr);

                



            }
            catch (Exception ex)
            {
                ;
            }
        }


            ////copy a form with placeholder texts processing action to the master data directory
            //File.Copy("c:\\temp\\PlugDiagramReportActionFormular.f22", new ProjectManager().Paths.Forms + "\\PlugDiagramReportActionFormular.f22", true);
            ////... and add it to project master data
            //StringCollection oProjectNewEntries = new StringCollection();
            //oProjectNewEntries.Add(@"PlugDiagramReportActionFormular.f22");
            //System.Collections.Hashtable oResult = new Masterdata().AddToProjectEx(m_oReportActionProject, oProjectNewEntries);
            ////prepare ReportBlock object
            //ReportBlock oReportBlock = new ReportBlock();
            //oReportBlock.Create(m_oReportActionProject);
            ////set a form with a placeholder texts processing action
            //oReportBlock.FormName = "PlugDiagramReportActionFormular";
            //oReportBlock.Type = DocumentTypeManager.DocumentType.PlugDiagram;
            ////set report processing action
            //oReportBlock.Action = "PlugDiagramReportAction";
            ////generate embedded report
            //ReportBlockReference oReportBlockReference = new Reports().CreateEmbeddedReport(oReportBlock, oPage, new PointD(10.0, 300.0));


        }
    }
