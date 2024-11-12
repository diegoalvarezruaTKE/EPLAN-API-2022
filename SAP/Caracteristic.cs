using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace EPLAN_API.SAP
{
    public class Caracteristic
    {
        //Reference of the caracteristic
        public string NameReference { get; set; }

        //Name of the caracteristc in the language Spanish
        public string NameES { get; set; }

        //Definition of Numerical value
        public bool IsNumeric { get; set; }

        //Definition of Text value
        public bool IsText { get; set; }

        //Values of the caracteristic
        public SortedList Values { get; set; }

        //Numerical Reference of the value
        public double NumVal { get; set; }

        //Text Reference of the value
        public string TextVal { get; set; }

        //Selected Reference of the value
        public String CurrentReference { get; set; }

        //Name of UI Control associated
        public string uiName { get; set; }

        //UI Control enabled
        public bool enabled { get; set; }

        //Combobox associated (in case list of values)
        public ComboBox combobox { get; set; }

        //Textbox associated (in case of numerical value)
        public TextBox textBox { get; set; }

        //Label asociated associated
        public Label label { get; set; }

        //Label asociated associated
        public int ProjectPropertyID { get; set; }

        //Mostrar
        public bool isVisible { get; set; }

        //Ptor reference
        public string pTORReference { get; set; }

        //Delegate to pass info to configurador
        public delegate void ComboboxDelegate(string data, string reference);

        //Event of delegate
        public event ComboboxDelegate Comboboxdata;

        //Delegate to pass info to configurador
        public delegate void TexboxDelegate(string data, string reference);

        //Event of delegate
        public event TexboxDelegate Textboxdata;

        public Caracteristic(string reference,
            string name,
            bool isNumeric,
            SortedList values,
            bool isVisible = true,
            string pTORReference = null)
        {
            NameReference = reference;
            NameES = name;
            IsNumeric = isNumeric;
            Values = values;
            this.isVisible = isVisible;
            this.pTORReference = pTORReference;
            IsText = false;

            CreateVisualComponents();
        }

        public Caracteristic(string reference,
            string name,
            bool isNumeric,
            double numVal,
            bool isVisible = true,
            string pTORReference = null)
        {
            NameReference = reference;
            NameES = name;
            IsNumeric = isNumeric;
            NumVal = numVal;
            this.isVisible = isVisible;
            this.pTORReference = pTORReference;
            IsText = false;

            CreateVisualComponents();

        }

        public Caracteristic(string reference,
            string name,
            bool isText,
            string textVal,
            bool isVisible = true,
            string pTORReference = null)
        {
            NameReference = reference;
            NameES = name;
            IsNumeric = false;
            IsText= isText;
            this.TextVal = textVal;
            this.isVisible = isVisible;
            this.pTORReference = pTORReference;

            CreateVisualComponents();

        }

        private void CreateVisualComponents()
        {
            if (isVisible)
            {
                //Create Label
                label = new Label()
                {
                    Name = String.Concat("Label_", NameReference),
                    AutoSize = true,
                    Size = new Size(35, 13),
                    Text = NameES,
                    BackColor = Color.Tomato,
                };

                if (!IsNumeric && !IsText)
                {
                    combobox = new ComboBox()
                    {
                        Name = String.Concat("ComboBox_", NameReference),
                        Size = new Size(220, 21),
                        DropDownStyle = ComboBoxStyle.DropDownList,
                        Enabled = true,
                    };
                    combobox.Items.AddRange(getListOfValueNames());
                    combobox.SelectedIndexChanged += new System.EventHandler(comboBoxCaractSelIndexChanged);

                    uiName = NameES;
                    enabled = true;
                }
                else
                {
                    textBox = new TextBox()
                    {
                        Name = String.Concat("Textbox_", NameReference),
                        Size = new Size(220, 21),
                        Enabled = true,
                    };

                    textBox.TextChanged += new EventHandler(TextboxTextChanged);
                    uiName = NameES;
                    enabled = true;
                }
            }

        }

        public void setActualNumVal(double NumValue)
        {   // Funcion para actualizar valor numerico y para modificar el color de subrayado
            NumVal = NumValue;
            if (enabled) //en proceso, iniciar numval en null y solo cambiar subrayado cuando se modifique a un valor numerico
            {
                CultureInfo culture = CultureInfo.CreateSpecificCulture("en-US");
                textBox.Text = String.Format(culture, "{0:0.00}", NumVal);
                label.BackColor = Color.Transparent;
            }
        }

        public bool setActualValue(String reference)
        {   
            if (IsNumeric)
                return false;

            if (!IsText)
            {
                if (Values == null)
                    return false;

                if (Values.ContainsKey(reference))
                {
                    //CurrentReference = reference;

                    if (combobox != null)
                    {
                        combobox.SelectedIndex = Values.IndexOfKey(reference);

                        if (enabled)
                        {
                            label.BackColor = Color.Transparent;
                        }
                    }
                    else
                    {
                        CurrentReference = reference;
                    }
                    return true;
                }
                else
                    return false;
            }
            else 
            {
                if (textBox != null) 
                {
                    textBox.Text= reference;
                }
                else
                {
                    TextVal = reference;
                }
                return true ;
            }
        }

        public string[] getListOfValueNames()
        {
            List<string> resList = new List<string>();
            foreach (String[] v in Values.Values)
            {
                resList.Add(v[0]);
            }

            return resList.ToArray();
        }

        private Double parseNumericValue(string str)
        {
            Double res = new Double();
            string[] auxStr = str.TrimStart(' ').Split(' ');
            if (auxStr != null)
            {
                Double.TryParse(auxStr[0], out res);
            }
            return res;
        }

        private void comboBoxCaractSelIndexChanged(object sender, EventArgs e)
        {
            CurrentReference=Values.GetKey((sender as ComboBox).SelectedIndex) as string;
            if (Comboboxdata != null)
            {
                Comboboxdata(CurrentReference, NameReference);
                label.BackColor= Color.Transparent;
            }
            else
                return;
        }

        private void TextboxTextChanged(object sender, EventArgs e)
        {
            if (IsNumeric)
            {
                double val;
                Double.TryParse((sender as TextBox).Text, out val);
                NumVal = val;
                if (Textboxdata != null)
                {
                    Textboxdata(NumVal.ToString(), NameReference);
                    label.BackColor = Color.Transparent;
                }
                else
                    return;
            }
            else
            {
                TextVal = (sender as TextBox).Text;
                if (Textboxdata != null)
                    Textboxdata(TextVal, NameReference);
                else
                    return;
            }
        }

    }
}
