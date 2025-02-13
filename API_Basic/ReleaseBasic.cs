using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.API_Basic
{
    public class ReleaseBasic
    {
        public Project oProject;

        public ReleaseBasic()
        {

            oProject = new ProjectManager().GetCurrentProjectWithDialog();
            Label oLabel = new Label();
            oLabel.DoLabel(oProject,
                "11200004271300-ARMARIO INTERIOR BASIC GEC EN Compras", // use implicitly last used scheme
                "MAIN",                 // don't filter
                "",                 // don't sort
                "en_EN",            // use English language  
                oProject.ProjectDirectoryPath + "\\LabelledPrj_2.txt", // destination file
                1,                  // Record Repeat = 2
                1);                 // Task Repeat = 1


        }
    }
}
