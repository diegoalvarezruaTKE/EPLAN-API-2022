using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.API_Basic
{
    public class ReleaseBasic
    {
        public Project oProject;

        public ReleaseBasic(Project oProject, string ubicacion)
        {

            this.oProject = oProject;
            Label oLabel = new Label();

            //Crear Estuctura de carpetas
            string path = oProject.ProjectDirectoryPath + "\\DOC_Basic";
            if (!Directory.Exists(path)) // Verifica si la carpeta ya existe
            {
                Directory.CreateDirectory(path);
            }

            path = oProject.ProjectDirectoryPath + "\\DOC_Basic\\Compras";
            if (!Directory.Exists(path)) // Verifica si la carpeta ya existe
            {
                Directory.CreateDirectory(path);
            }


            switch (ubicacion)
            {
                case "MAIN":
                    oLabel.DoLabel(oProject,
                            "11200004271300-ARMARIO INTERIOR BASIC GEC EN Compras", // use implicitly last used scheme
                            ubicacion,                 // don't filter
                            "",                 // don't sort
                            "en_EN",            // use English language  
                            path + "\\BOM_11200004271300-ARMARIO INTERIOR BASIC GEC EN.xlsx", // destination file
                            1,                  // Record Repeat = 2
                            1);                 // Task Repeat = 1
                    break;

                case "CDS":
                    oLabel.DoLabel(oProject,
                            "11200004366200_ENSAMBLAJE_CAJA_DERIV_SUP_BASIC_Compras", // use implicitly last used scheme
                            ubicacion,                 // don't filter
                            "",                 // don't sort
                            "en_EN",            // use English language  
                            path + "\\BOM_11200004366200_ENSAMBLAJE_CAJA_DERIV_SUP_BASIC_Compras.xlsx", // destination file
                            1,                  // Record Repeat = 2
                            1);                 // Task Repeat = 1
                    break;

                case "CDI":
                    oLabel.DoLabel(oProject,
                            "11200004366100_ENSAMBLAJE_CAJA_DERIV_INF_BASIC_Compras", // use implicitly last used scheme
                            ubicacion,                 // don't filter
                            "",                 // don't sort
                            "en_EN",            // use English language  
                            path + "\\BOM_11200004366100_ENSAMBLAJE_CAJA_DERIV_INF_BASIC_Compras.xlsx", // destination file
                            1,                  // Record Repeat = 2
                            1);                 // Task Repeat = 1
                    break;
            }





        }
    }
}
