using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API.Forms
{
    public partial class NumeroEsquema : Form
    {

        String currentDirectory;
        String strPath;
        public string strNumeroEsquema { get; set; }
        public string strProjectName { get; set; }

        public NumeroEsquema(string strPath, string strProjectName)
        {
            this.strPath = strPath;
            InitializeComponent();
            this.strProjectName = strProjectName;
            strNumeroEsquema = null;
        }

        public int checkLastDrawing()
        {
            int res = 0;


            string[] strArr = Directory.GetDirectories(strPath);

            int highNum = 0;
            string subPath = null;
            foreach (string str in strArr)
            {
                string[] strArrAux = str.Split('\\');
                String auxStr = strArrAux[strArrAux.Length - 1];
                int num;
                int.TryParse(auxStr, out num);
                if (num > highNum)
                {
                    highNum = num;
                    subPath = str;
                }
            }

            currentDirectory = subPath;

            if (subPath != null)
            {
                strArr = Directory.GetDirectories(subPath);
                highNum = 0;
                foreach (string str in strArr)
                {
                    string[] strArrAux = str.Split('\\');
                    String auxStr = strArrAux[strArrAux.Length - 1].Substring(0, 4);
                    int num;
                    int.TryParse(auxStr, out num);
                    if (num > highNum)
                    {
                        highNum = num;
                        subPath = str;
                    }
                }
            }

            res = highNum + 1;

            return res;

        }

        private void check_NewDrawin_Clik(object sender, EventArgs e)
        {

        }

        private void buttonAcept_Click(object sender, EventArgs e)
        {
            if (tB_Drawing_Number.Text != null)
            {
                strNumeroEsquema = tB_Drawing_Number.Text;
            }
            if (checkBox_NewDrawing.Checked)
                Directory.CreateDirectory(String.Concat(currentDirectory, "\\", checkLastDrawing().ToString(), "_", strProjectName));
            this.Close();
        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
