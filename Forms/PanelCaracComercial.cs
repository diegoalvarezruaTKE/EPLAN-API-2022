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
    public partial class PanelCaracComercial : Form
    {
        public List<Electric> oElectricList;

        public PanelCaracComercial(List<Electric> electricList)
        {
            oElectricList = electricList;

            InitializeComponent();
            LinkComponets();

        }

        private void LinkComponets()
        {
            //PTOR 5 - LIGHTING
            GroupBox groupLighting;
            groupLighting = new GroupBox()
            {
                Location = new Point(12, 48),
                Name = "safetyDevices",
                Size = new Size(10, 10),
                TabIndex = 1,
                TabStop = false,
                Text = "LIGHTING",
                Visible = true
            };
            groupLighting.SuspendLayout();
            Controls.Add(groupLighting);

            //PTOR 9 - SAFETY DEVICES
            GroupBox groupSafetyDevices;
            groupSafetyDevices = new GroupBox()
            {
                Location = new Point(12, 48),
                Name = "safetyDevices",
                Size = new Size(10, 10),
                TabIndex = 1,
                TabStop = false,
                Text = "SAFETY DEVICES",
                Visible = true,
            };
            groupSafetyDevices.SuspendLayout();
            Controls.Add(groupSafetyDevices);
            
            //PTOR 10 - DRIVE SYSTEM
            GroupBox groupDriveSystem;
            groupDriveSystem = new GroupBox() {
                Location = new Point(12, 48),
                Name = "groupDriveSystem",
                Size = new Size(10, 10),
                TabIndex = 1,
                TabStop = false,
                Text = "DRIVE SYSTEM",
                Visible = true,
            };
            groupDriveSystem.SuspendLayout();
            Controls.Add(groupDriveSystem);

            //Other
            GroupBox groupOther;
            groupOther = new GroupBox()
            {
                Location = new Point(12, 48),
                Name = "groupOther",
                Size = new Size(10, 10),
                TabIndex = 1,
                TabStop = false,
                Text = "GENERAL",
                Visible = true,
            };
            groupOther.SuspendLayout();
            Controls.Add(groupOther);

            int xLocationLighting, yLocationLighting,
                xLocationSafety, yLocationSafety,
                xLocationDrive, yLocationDrive,
                xLocationOther, yLocationOther,
                rowCountLighting, rowCountSafety, rowCountDrive, rowCountOther;

            const int xOffset = 280;
            const int yOffset = 54;
            const int ylabelOffset = 22;
            const int nRows = 15;
            const int xInit = 12;
            const int yInit = 24;

            //Init values
            xLocationLighting = xInit;
            yLocationLighting = yInit;
            xLocationSafety = xInit;
            yLocationSafety = yInit;
            xLocationDrive = xInit;
            yLocationDrive = yInit;
            xLocationOther = xInit;
            yLocationOther = yInit;
            rowCountSafety = 1;
            rowCountLighting = 1;
            rowCountDrive = 1;
            rowCountOther = 1;

            foreach (Caracteristic c in oElectricList[0].OrderCaractComercial.Values)
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
                        //PTOR 5 - LIGHTING
                        case 5:
                            groupLighting.Controls.Add(c.label);
                            c.label.Location = new Point(xLocationLighting, yLocationLighting);

                            if (!c.IsNumeric && c.combobox != null)
                            {
                                groupLighting.Controls.Add(c.combobox);

                                c.combobox.Location = new Point(xLocationLighting, yLocationLighting + ylabelOffset);

                                if (groupLighting.Size.Width < c.combobox.Location.X)
                                    groupLighting.Size = new Size(c.combobox.Location.X + c.combobox.Size.Width + 12,
                                        groupLighting.Size.Height);

                                if (groupLighting.Size.Height < c.combobox.Location.Y)
                                    groupLighting.Size = new Size(groupLighting.Size.Width,
                                        c.combobox.Location.Y + c.combobox.Size.Height + 12);

                            }

                            else if (c.textBox != null)
                            {
                                groupLighting.Controls.Add(c.textBox);

                                c.textBox.Location = new Point(xLocationLighting, yLocationLighting + ylabelOffset);

                                if (groupLighting.Size.Width < c.textBox.Location.X)
                                    groupLighting.Size = new Size(c.textBox.Location.X + c.textBox.Size.Width + 12,
                                        groupLighting.Size.Height);

                                if (groupLighting.Size.Height < c.textBox.Location.Y)
                                    groupLighting.Size = new Size(groupLighting.Size.Width,
                                        c.textBox.Location.Y + c.textBox.Size.Height + 12);
                            }

                            yLocationLighting += yOffset;

                            if (rowCountLighting >= nRows)
                            {
                                rowCountLighting = 1;
                                xLocationLighting += + xOffset;
                                yLocationLighting = yInit;
                            }
                            else
                                rowCountLighting ++;

                            break;

                        //PTOR 9 - SAFETY DEVICES
                        case 9:
                            groupSafetyDevices.Controls.Add(c.label);
                            c.label.Location = new Point(xLocationSafety, yLocationSafety);

                            if (!c.IsNumeric && c.combobox != null)
                            {
                                groupSafetyDevices.Controls.Add(c.combobox);

                                c.combobox.Location = new Point(xLocationSafety, yLocationSafety + ylabelOffset);

                                if (groupSafetyDevices.Size.Width < c.combobox.Location.X)
                                    groupSafetyDevices.Size = new Size(c.combobox.Location.X + c.combobox.Size.Width + 12,
                                        groupSafetyDevices.Size.Height);

                                if (groupSafetyDevices.Size.Height < c.combobox.Location.Y)
                                    groupSafetyDevices.Size = new Size(groupSafetyDevices.Size.Width,
                                        c.combobox.Location.Y + c.combobox.Size.Height + 12);

                            }

                            else if (c.textBox != null)
                            {
                                groupSafetyDevices.Controls.Add(c.textBox);

                                c.textBox.Location = new Point(xLocationSafety, yLocationSafety + ylabelOffset);

                                if (groupSafetyDevices.Size.Width < c.textBox.Location.X)
                                    groupSafetyDevices.Size = new Size(c.textBox.Location.X + c.textBox.Size.Width + 12,
                                        groupSafetyDevices.Size.Height);

                                if (groupSafetyDevices.Size.Height < c.textBox.Location.Y)
                                    groupSafetyDevices.Size = new Size(groupSafetyDevices.Size.Width,
                                        c.textBox.Location.Y + c.textBox.Size.Height + 12);
                            }

                            yLocationSafety += yOffset;

                            if (rowCountSafety >= nRows)
                            {
                                rowCountSafety = 1;
                                xLocationSafety += xOffset;
                                yLocationSafety = yInit;
                            }
                            else
                                rowCountSafety ++;


                            break;

                        //PTOR 10 - DRIVE SYSTEM
                        case 10:
                            groupDriveSystem.Controls.Add(c.label);
                            c.label.Location = new Point(xLocationDrive, yLocationDrive);

                            if (!c.IsNumeric && c.combobox != null)
                            {
                                groupDriveSystem.Controls.Add(c.combobox);

                                c.combobox.Location = new Point(xLocationDrive, yLocationDrive + ylabelOffset);

                                if (groupDriveSystem.Size.Width < c.combobox.Location.X)
                                    groupDriveSystem.Size = new Size(c.combobox.Location.X + c.combobox.Size.Width + 12,
                                        groupDriveSystem.Size.Height);

                                if (groupDriveSystem.Size.Height < c.combobox.Location.Y)
                                    groupDriveSystem.Size = new Size(groupDriveSystem.Size.Width,
                                        c.combobox.Location.Y + c.combobox.Size.Height + 12);

                            }
                            else if (c.textBox != null)
                            {
                                groupDriveSystem.Controls.Add(c.textBox);

                                c.textBox.Location = new Point(xLocationDrive, yLocationDrive + ylabelOffset);

                                if (groupDriveSystem.Size.Width < c.textBox.Location.X)
                                    groupDriveSystem.Size = new Size(c.textBox.Location.X + c.textBox.Size.Width + 12,
                                        groupDriveSystem.Size.Height);

                                if (groupDriveSystem.Size.Height < c.textBox.Location.Y)
                                    groupDriveSystem.Size = new Size(groupDriveSystem.Size.Width,
                                        c.textBox.Location.Y + c.textBox.Size.Height + 12);
                            }

                            yLocationDrive += yOffset;

                            if (rowCountDrive >= nRows)
                            {
                                rowCountDrive = 1;
                                xLocationDrive += xOffset;
                                yLocationDrive = yInit;
                            }
                            else
                                rowCountDrive++;
                            break;

                        default:
                            groupOther.Controls.Add(c.label);
                            c.label.Location = new Point(xLocationOther, yLocationOther);

                            if (!c.IsNumeric && c.combobox != null)
                            {
                                groupOther.Controls.Add(c.combobox);

                                c.combobox.Location = new Point(xLocationOther, yLocationOther + ylabelOffset);

                                if (groupOther.Size.Width < c.combobox.Location.X)
                                    groupOther.Size = new Size(c.combobox.Location.X + c.combobox.Size.Width + 12,
                                        groupOther.Size.Height);

                                if (groupOther.Size.Height < c.combobox.Location.Y)
                                    groupOther.Size = new Size(groupOther.Size.Width,
                                        c.combobox.Location.Y + c.combobox.Size.Height + 12);

                            }
                            else if (c.textBox != null)
                            {
                                groupOther.Controls.Add(c.textBox);

                                c.textBox.Location = new Point(xLocationOther, yLocationOther + ylabelOffset);

                                if (groupOther.Size.Width < c.textBox.Location.X)
                                    groupOther.Size = new Size(c.textBox.Location.X + c.textBox.Size.Width + 12,
                                        groupOther.Size.Height);

                                if (groupOther.Size.Height < c.textBox.Location.Y)
                                    groupOther.Size = new Size(groupOther.Size.Width,
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

            groupLighting.Location = new Point(groupOther.Location.X + groupOther.Width + 12,
                groupOther.Location.Y);

            groupSafetyDevices.Location = new Point(groupLighting.Location.X + groupLighting.Width + 12,
                groupLighting.Location.Y);

            groupDriveSystem.Location = new Point(groupSafetyDevices.Location.X + groupSafetyDevices.Width + 12,
                groupSafetyDevices.Location.Y);

        }

    }
}
