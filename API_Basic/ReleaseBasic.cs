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
            string pathBasic = oProject.ProjectDirectoryPath + "\\DOC_Basic";
            if (!Directory.Exists(pathBasic)) // Verifica si la carpeta ya existe
            {
                Directory.CreateDirectory(pathBasic);
            }

            String pathCompras = oProject.ProjectDirectoryPath + "\\DOC_Basic\\Compras";
            if (!Directory.Exists(pathCompras)) // Verifica si la carpeta ya existe
            {
                Directory.CreateDirectory(pathCompras);
            }

            String pathProveedor = oProject.ProjectDirectoryPath + "\\DOC_Basic\\Proveedor";
            if (!Directory.Exists(pathProveedor)) // Verifica si la carpeta ya existe
            {
                Directory.CreateDirectory(pathProveedor);
            }


            switch (ubicacion)
            {
                case "MAIN":
                    //Compras
                    oLabel.DoLabel(oProject,
                            "11200004271300-ARMARIO INTERIOR BASIC GEC EN Compras", // use implicitly last used scheme
                            ubicacion,                
                            "",                 // don't sort
                            "en_EN",            // use English language  
                            pathCompras + "\\BOM_11200004271300-ARMARIO INTERIOR BASIC GEC EN.xlsx", // destination file
                            1,                  // Record Repeat = 2
                            1);                 // Task Repeat = 1


                    //Proveedor
                    //01_DOC PROV_LDM_MAIN
                    oLabel.DoLabel(oProject,
                            "01_DOC PROV_LDM_MAIN",
                            "TKN_LDM_MAIN",                 
                            "",                 
                            "en_EN",
                            pathProveedor + "\\MAIN_1_LISTA DE MATERIALES.xlsx", 
                            1,                  
                            1);

                    //02_DOC PROV_LDM MANGUERAS+CONECTORES_MAIN
                    oLabel.DoLabel(oProject,
                            "02_DOC PROV_LDM MANGUERAS+CONECTORES_MAIN",
                            "02_DOC PROV_LDM MANGUERAS+CONECTORES_MAIN",
                            "",
                            "en_EN",
                            pathProveedor + "\\MAIN_2_LISTA DE MATERIALES MANGUERAS Y CONECTORES.xlsx",
                            1,
                            1);

                    //06_DOC PROV_LIST CABLE_MAIN
                    oLabel.DoLabel(oProject,
                            "06_DOC PROV_LIST CABLE_MAIN",
                            "06_DOC PROV_MAIN",
                            "",
                            "en_EN",
                            pathProveedor + "\\MAIN_6_LISTA CORTE MANGUERAS.xlsx",
                            1,
                            1);

                    //07_DOC PROV_LIST CABLE+FC_MAIN
                    oLabel.DoLabel(oProject,
                            "07_DOC PROV_LIST CABLE+FC_MAIN",
                            "07_DOC PROV_CAB+FC_MAIN",
                            "",
                            "en_EN",
                            pathProveedor + "\\MAIN_7_TABLA RESUMEN CABLE + FC.xlsx",
                            1,
                            1);

                    //08_DOC PROV_LIST CORTE CAB UNIF_MAIN
                    oLabel.DoLabel(oProject,
                            "08_DOC PROV_LIST CORTE CAB UNIF_MAIN",
                            "INTERNAL CONECTIONS_MAIN",
                            "COLOR + SECCIÓN",
                            "en_EN",
                            pathProveedor + "\\MAIN_8_PLANO DE CORTE DE CABLE UNIFILAR.xlsx",
                            1,
                            1);

                    break;

                case "CDS":
                    oLabel.DoLabel(oProject,
                            "11200004366200_ENSAMBLAJE_CAJA_DERIV_SUP_BASIC_Compras", // use implicitly last used scheme
                            ubicacion,                 // don't filter
                            "",                 // don't sort
                            "en_EN",            // use English language  
                            pathCompras + "\\BOM_11200004366200_ENSAMBLAJE_CAJA_DERIV_SUP_BASIC_Compras.xlsx", // destination file
                            1,                  // Record Repeat = 2
                            1);                 // Task Repeat = 1

                    //Proveedor
                    //01_DOC PROV_LDM_CDS
                    oLabel.DoLabel(oProject,
                            "01_DOC PROV_LDM_MAIN",
                            "TKN_LDM_CDS",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDS_1_LISTA DE MATERIALES.xlsx",
                            1,
                            1);

                    //02_DOC PROV_LDM MANGUERAS+CONECTORES_CDS
                    oLabel.DoLabel(oProject,
                            "02_DOC PROV_LDM MANGUERAS+CONECTORES_MAIN",
                            "02_DOC PROV_LDM MANGUERAS+CONECTORES_CDS",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDS_2_LISTA DE MATERIALES MANGUERAS Y CONECTORES.xlsx",
                            1,
                            1);

                    //06_DOC PROV_LIST CABLE_CDS
                    oLabel.DoLabel(oProject,
                            "06_DOC PROV_LIST CABLE_MAIN",
                            "06_DOC PROV_CDS",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDS_6_LISTA CORTE MANGUERAS.xlsx",
                            1,
                            1);

                    //07_DOC PROV_LIST CABLE+FC_CDS
                    oLabel.DoLabel(oProject,
                            "07_DOC PROV_LIST CABLE+FC_MAIN",
                            "07_DOC PROV_CAB+FC_CDS",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDS_7_TABLA RESUMEN CABLE + FC.xlsx",
                            1,
                            1);

                    //08_DOC PROV_LIST CORTE CAB UNIF_CDS
                    oLabel.DoLabel(oProject,
                            "08_DOC PROV_LIST CORTE CAB UNIF_MAIN",
                            "INTERNAL CONECTIONS_CDS",
                            "COLOR + SECCIÓN",
                            "en_EN",
                            pathProveedor + "\\CDS_8_PLANO DE CORTE DE CABLE UNIFILAR.xlsx",
                            1,
                            1);

                    break;

                case "CDI":
                    oLabel.DoLabel(oProject,
                            "11200004366100_ENSAMBLAJE_CAJA_DERIV_INF_BASIC_Compras", // use implicitly last used scheme
                            ubicacion,                 // don't filter
                            "",                 // don't sort
                            "en_EN",            // use English language  
                            pathCompras + "\\BOM_11200004366100_ENSAMBLAJE_CAJA_DERIV_INF_BASIC_Compras.xlsx", // destination file
                            1,                  // Record Repeat = 2
                            1);                 // Task Repeat = 1

                    //Proveedor
                    //01_DOC PROV_LDM_CDI
                    oLabel.DoLabel(oProject,
                            "01_DOC PROV_LDM_MAIN",
                            "TKN_LDM_CDI",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDI_1_LISTA DE MATERIALES.xlsx",
                            1,
                            1);

                    //02_DOC PROV_LDM MANGUERAS+CONECTORES_CDI
                    oLabel.DoLabel(oProject,
                            "02_DOC PROV_LDM MANGUERAS+CONECTORES_MAIN",
                            "02_DOC PROV_LDM MANGUERAS+CONECTORES_CDI",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDI_2_LISTA DE MATERIALES MANGUERAS Y CONECTORES.xlsx",
                            1,
                            1);

                    //06_DOC PROV_LIST CABLE_CDI
                    oLabel.DoLabel(oProject,
                            "06_DOC PROV_LIST CABLE_MAIN",
                            "06_DOC PROV_CDI",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDI_6_LISTA CORTE MANGUERAS.xlsx",
                            1,
                            1);

                    //07_DOC PROV_LIST CABLE+FC_CDI
                    oLabel.DoLabel(oProject,
                            "07_DOC PROV_LIST CABLE+FC_MAIN",
                            "07_DOC PROV_CAB+FC_CDI",
                            "",
                            "en_EN",
                            pathProveedor + "\\CDI_7_TABLA RESUMEN CABLE + FC.xlsx",
                            1,
                            1);

                    //08_DOC PROV_LIST CORTE CAB UNIF_CDI
                    oLabel.DoLabel(oProject,
                            "08_DOC PROV_LIST CORTE CAB UNIF_MAIN",
                            "INTERNAL CONECTIONS_CDI",
                            "COLOR + SECCIÓN",
                            "en_EN",
                            pathProveedor + "\\CDI_8_PLANO DE CORTE DE CABLE UNIFILAR.xlsx",
                            1,
                            1);

                    break;
            }





        }
    }
}
