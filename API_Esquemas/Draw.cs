﻿using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using EPLAN_API.Forms;
using EPLAN_API.SAP;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API.User
{
    public class Draw
    {
        private Electric electric;
        private string ProjectPath;
        private PathInfo EPLANPaths;
        private Project oProject;
        private string drawingNumber;
        public delegate void ProejctOpenedDelegate(Project project);
        public event ProejctOpenedDelegate ProjectOpenedToConfigurador;
        public delegate void ProgressChangedDelegate(int progress);
        public event ProgressChangedDelegate ProgressChangedToConfigurador;
        DrawArmStandardIntEN drawStandardArmIntEN;
        DrawArmExteriorEN drawArmExteriorEN;
        DrawArmStandardIntASME drawStandardIntASME;
        DrawBasicArmIntEN drawBasicArmIntEN;

        public Draw(Electric oElectric)
        {
            electric = oElectric;
            EPLANPaths = new PathInfo();
        }

        public void StartDrawing()
        {
            //Check if all caracteristic are filled
            if (electric.checkValues())
            {
                if (GenerateProjectDirectory())
                {
                    RestoreProject();
                    OpenProject();
                    SelectDrawing();
                }
            }
            else
            {
                MessageBox.Show("No están valoradas todas las caracteristicas comerciales y de ingeniería", "Error");
            }
        }

        private bool GenerateProjectDirectory()
        {
            string denominacionProyecto = (electric.CaractComercial["TNCR_COM_NOMBREOBRA_VBACK"] as Caracteristic).TextVal;
            String templateProjectPath;

            DialogResult dialogResult = MessageBox.Show("¿Quieres crear nuevo proyecto?", "Proyecto", MessageBoxButtons.YesNo);

            if (dialogResult == DialogResult.Yes)
            {
                NumeroEsquema numeroEsquema = new NumeroEsquema(EPLANPaths.Schemes, denominacionProyecto);
                numeroEsquema.ShowDialog();

                if (numeroEsquema.strNumeroEsquema != null)
                {
                    drawingNumber = numeroEsquema.strNumeroEsquema;
                    ProjectPath = String.Concat(EPLANPaths.Projects, "\\01_Obra_Nueva\\", numeroEsquema.strNumeroEsquema, "_", denominacionProyecto);
                    templateProjectPath = String.Concat(EPLANPaths.Projects, "\\01_Obra_Nueva\\11500xxxx_PLANTILLA");
                    CopyFilesRecursively(templateProjectPath, ProjectPath);
                    return true;
                }
            }

            return false;
        }

        private void RestoreProject()
        {
            string[] fileTemplatePaths = Directory.GetFiles(EPLANPaths.Templates);
            string TemplateFileName = null;
            Caracteristic ubicacionArmario = electric.CaractComercial["F53ZUB7"] as Caracteristic;
            Caracteristic norma = electric.CaractComercial["FNORM"] as Caracteristic;
            Caracteristic maniobra = electric.CaractIng["MANIOBRA"] as Caracteristic;

            Restore oRestore = new Restore();
            StringCollection oArchives = new StringCollection();

            //Armario Standard
            if (!maniobra.CurrentReference.Equals("BASIC"))
            {
                if (norma.CurrentReference.Equals("EN"))
                {
                    //Armario interior EN115
                    if (ubicacionArmario.CurrentReference.Equals("INNENOBEN"))
                        TemplateFileName = "GEC_EN115_Arm_Int_Base";
                    //Armario exterior EN115
                    else
                        TemplateFileName = "GEC_EN115_Arm_Ext_Base";
                }
                //Armario interior ASME
                else if (norma.CurrentReference.Equals("ASME") ||
                    norma.CurrentReference.Equals("CSA"))
                    TemplateFileName = "GEC_ASME_Arm_Int_Base";
            }
            //Armario Basic
            else
            {
                TemplateFileName = "GEC_EN115_Base_Basic";
            }

            //Restore base project
            foreach (string file in fileTemplatePaths)
            {
                if (file.Contains(TemplateFileName))
                {
                    oArchives.Add(file);
                    oRestore.Project(oArchives,
                                    String.Concat(ProjectPath, "\\1_Electrico\\1_Esquema\\"),
                                    "Base",
                                    false,
                                    false);

                    break;
                }
            }

            ProjectPath = String.Concat(ProjectPath,
                            "\\1_Electrico\\1_Esquema\\Base");
        }

        private void OpenProject()
        {
            string denominacionProyecto = (electric.CaractComercial["TNCR_COM_NOMBREOBRA_VBACK"] as Caracteristic).TextVal;

            oProject = new ProjectManager().OpenProject(ProjectPath);
            oProject.ProjectName = String.Concat("0-55.1-3.", drawingNumber, ".00", "_", denominacionProyecto.Replace(' ', '_'));
            ProjectOpened(oProject);
        }

        private void SelectDrawing()
        {
            Caracteristic ubicacionArmario = electric.CaractComercial["F53ZUB7"] as Caracteristic;
            Caracteristic norma = electric.CaractComercial["FNORM"] as Caracteristic;
            Caracteristic maniobra = electric.CaractIng["MANIOBRA"] as Caracteristic;

            //Armario Standard
            if (!maniobra.CurrentReference.Equals("BASIC"))
            {
                if (norma.CurrentReference.Equals("EN"))
                {
                    //Armario interior EN115
                    if (ubicacionArmario.CurrentReference.Equals("INNENOBEN"))
                    {
                        drawStandardArmIntEN = new DrawArmStandardIntEN(oProject, electric);
                        drawStandardArmIntEN.ProgressChangedToDraw += new DrawArmStandardIntEN.ProgressChangedDelegate(UpdateProgressBar);
                        drawStandardArmIntEN.DrawMacro();
                    }

                    //Armario exterior EN115
                    else
                    {
                        drawArmExteriorEN = new DrawArmExteriorEN(oProject, electric);
                        drawArmExteriorEN.ProgressChangedToDraw += new DrawArmExteriorEN.ProgressChangedDelegate(UpdateProgressBar);
                        drawArmExteriorEN.DrawMacro();
                        //drawArmExteriorEN.ProgressChanged += UpdateProgressBar;
                    }
                }
                //Armario interior ASME
                else if (norma.CurrentReference.Equals("ASME") ||
                    norma.CurrentReference.Equals("CSA"))
                {
                    drawStandardIntASME = new DrawArmStandardIntASME(oProject, electric);
                    drawStandardIntASME.ProgressChangedToDraw += new DrawArmStandardIntASME.ProgressChangedDelegate(UpdateProgressBar);
                    drawStandardIntASME.DrawMacro();
                }
            }
            //Armario Basic
            else
            {
                drawBasicArmIntEN = new DrawBasicArmIntEN(oProject, electric);
                drawBasicArmIntEN.ProgressChangedToDraw += new DrawBasicArmIntEN.ProgressChangedDelegate(UpdateProgressBar);
                drawBasicArmIntEN.DrawMacro();
            }
        }

        private void UpdateProgressBar(int value)
        {
            ProgressChangedToConfigurador(value);
        }

        private void CopyFilesRecursively(string sourcePath, string targetPath)
        {
            //Now Create all of the directories
            foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourcePath, targetPath));
            }

            //Copy all the files & Replaces any files with the same name
            foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
            {
                File.Copy(newPath, newPath.Replace(sourcePath, targetPath), true);
            }
        }

        private void ProjectOpened(Project project)
        {
            ProjectOpenedToConfigurador(project);
        }

    }
}
