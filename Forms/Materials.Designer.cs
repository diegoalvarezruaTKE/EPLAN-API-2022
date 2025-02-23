namespace EPLAN_API.Forms
{
    partial class Materials
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.b_List = new System.Windows.Forms.Button();
            this.cB_Locations = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // b_List
            // 
            this.b_List.Location = new System.Drawing.Point(12, 67);
            this.b_List.Name = "b_List";
            this.b_List.Size = new System.Drawing.Size(75, 23);
            this.b_List.TabIndex = 0;
            this.b_List.Text = "Lista";
            this.b_List.UseVisualStyleBackColor = true;
            this.b_List.Click += new System.EventHandler(this.b_List_Click);
            // 
            // cB_Locations
            // 
            this.cB_Locations.FormattingEnabled = true;
            this.cB_Locations.Location = new System.Drawing.Point(94, 67);
            this.cB_Locations.Name = "cB_Locations";
            this.cB_Locations.Size = new System.Drawing.Size(185, 21);
            this.cB_Locations.TabIndex = 1;
            // 
            // Materials
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(478, 199);
            this.Controls.Add(this.cB_Locations);
            this.Controls.Add(this.b_List);
            this.Name = "Materials";
            this.Text = "Lista de Materiales";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button b_List;
        private System.Windows.Forms.ComboBox cB_Locations;
    }
}