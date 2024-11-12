using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPLAN_API.User
{
    class MotorCurrentData
    {
        public double Power { get; set; }
        public String Voltage { get; set; }
        public int Frequency { get; set; }
        public double Current { get; set; }

        public MotorCurrentData(double power, string voltage, int frequency, double current)
        {
            Power = power;
            Voltage = voltage;
            Frequency = frequency;
            Current = current;
        }
    }
}
