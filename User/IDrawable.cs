using Eplan.EplApi.DataModel;
using EPLAN_API.SAP;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.User
{
    public interface IDrawable
    {
        void SetGECParameter(Project oProject, Electric oElectric, string address, uint value, bool changeText = false);

        void SetGECParameter(Project oProject, Electric oElectric, string address, string value, bool changeText = false);
    }
}
