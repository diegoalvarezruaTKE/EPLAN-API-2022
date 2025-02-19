using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.SAP
{
    public class Cable
    {
        public string IME { get; set; }

        public string Name { get; set; }

        public string Code { get; set; }

        public string Description { get; set; }
        public double Lenght { get; set; }
        public string ParentName { get; set; }

        public string ParentCode { get; set; }
        public Cable(string IME, string name, string code, string description, double lenght, string parentName, string parentCode)
        {
            this.Name = name;
            this.Description = description;
            this.Lenght = lenght;
            ParentName = parentName;
            ParentCode = parentCode;
            Code = code;
            this.IME = IME;
        }
    }
}
