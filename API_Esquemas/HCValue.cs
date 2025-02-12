using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.User
{
    public class HCValue
    {

        //Caracteristic reference in SAP
        public string SAPReference { get; set; }

        //Name in SAP in Spanish languaje
        public string SAPNameES { get; set; }

        //Possible names in HC
        public string [] HCNames { get; set; }

        public HCValue(string SAPReference, string SAPNameES) 
        {
            this.SAPReference = SAPReference;
            this.SAPNameES = SAPNameES;
        }

        public HCValue(string SAPReference, string SAPNameES, string[] HCNames)
        {
            this.SAPReference = SAPReference;
            this.SAPNameES = SAPNameES;
            this.HCNames = HCNames;
        }
    }
}
