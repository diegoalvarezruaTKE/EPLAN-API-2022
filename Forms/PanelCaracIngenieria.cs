using EPLAN_API.SAP;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EPLAN_API_2022.Forms
{
    public partial class PanelCaracIngenieria : Form
    {
        public List<Electric> oElectricList;

        public PanelCaracIngenieria(List<Electric> electricList)
        {
            oElectricList = electricList;

            InitializeComponent();
            LinkComponents();
        }

        private void LinkComponents()
        {
            //General
            GroupBox groupGeneral;
            groupGeneral = new GroupBox()
            {
                Location = new Point(12, 48),
                Name = "groupGeneral",
                Size = new Size(10, 10),
                TabIndex = 1,
                TabStop = false,
                Text = "GENERAL",
                Visible = true,
            };
            groupGeneral.SuspendLayout();
            Controls.Add(groupGeneral);

            int xLocationOther, yLocationOther,
                 rowCountOther;

            const int xOffset = 280;
            const int yOffset = 54;
            const int ylabelOffset = 22;
            const int nRows = 15;
            const int xInit = 12;
            const int yInit = 24;

            //Init values
            xLocationOther = xInit;
            yLocationOther = yInit;
            rowCountOther = 1;

            foreach (Caracteristic c in oElectricList[0].OrderCaractIng.Values)
            {
                if (c.isVisible)
                {
                    int pTORindex;
                    if (c.pTORReference != null)
                        int.TryParse(c.pTORReference.Split('.')[0], out pTORindex);
                    else
                        pTORindex = 0;

                    switch (pTORindex)
                    {
                        default:
                            groupGeneral.Controls.Add(c.label);
                            c.label.Location = new Point(xLocationOther, yLocationOther);

                            if (!c.IsNumeric && c.combobox != null)
                            {
                                groupGeneral.Controls.Add(c.combobox);

                                c.combobox.Location = new Point(xLocationOther, yLocationOther + ylabelOffset);

                                if (groupGeneral.Size.Width < c.combobox.Location.X)
                                    groupGeneral.Size = new Size(c.combobox.Location.X + c.combobox.Size.Width + 12,
                                        groupGeneral.Size.Height);

                                if (groupGeneral.Size.Height < c.combobox.Location.Y)
                                    groupGeneral.Size = new Size(groupGeneral.Size.Width,
                                        c.combobox.Location.Y + c.combobox.Size.Height + 12);

                            }
                            else if (c.textBox != null)
                            {
                                groupGeneral.Controls.Add(c.textBox);

                                c.textBox.Location = new Point(xLocationOther, yLocationOther + ylabelOffset);

                                if (groupGeneral.Size.Width < c.textBox.Location.X)
                                    groupGeneral.Size = new Size(c.textBox.Location.X + c.textBox.Size.Width + 12,
                                        groupGeneral.Size.Height);

                                if (groupGeneral.Size.Height < c.textBox.Location.Y)
                                    groupGeneral.Size = new Size(groupGeneral.Size.Width,
                                        c.textBox.Location.Y + c.textBox.Size.Height + 12);
                            }

                            yLocationOther += yOffset;

                            if (rowCountOther >= nRows)
                            {
                                rowCountOther = 1;
                                xLocationOther += xOffset;
                                yLocationOther = yInit;
                            }
                            else
                                rowCountOther++;

                            break;

                    }

                }
            }

        }
    }
}
