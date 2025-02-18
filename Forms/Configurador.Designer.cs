namespace EPLAN_API_2022.Forms
{
    partial class Configurador
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Configurador));
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.export_Button = new System.Windows.Forms.Button();
            this.tB_OE = new System.Windows.Forms.TextBox();
            this.b_ReadCaract = new System.Windows.Forms.Button();
            this.b_EntradaManual = new System.Windows.Forms.Button();
            this.progressBar_Draw = new System.Windows.Forms.ProgressBar();
            this.b_Draw = new System.Windows.Forms.Button();
            this.b_Cables = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.panelApp = new System.Windows.Forms.Panel();
            this.MenuVertical = new System.Windows.Forms.Panel();
            this.buttonFormHC = new System.Windows.Forms.Button();
            this.buttonFormGEC = new System.Windows.Forms.Button();
            this.logoTKE = new System.Windows.Forms.PictureBox();
            this.buttonFormCaracComer = new System.Windows.Forms.Button();
            this.buttonFormCaracIng = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.PanelTitulo = new System.Windows.Forms.Panel();
            this.cBox_Obra = new System.Windows.Forms.ComboBox();
            this.ImportButton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tbStatus = new System.Windows.Forms.TextBox();
            this.MenuVertical.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.logoTKE)).BeginInit();
            this.PanelTitulo.SuspendLayout();
            this.SuspendLayout();
            // 
            // export_Button
            // 
            this.export_Button.BackColor = System.Drawing.SystemColors.HighlightText;
            this.export_Button.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.export_Button.Cursor = System.Windows.Forms.Cursors.Hand;
            this.export_Button.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.export_Button.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.export_Button.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.export_Button.Font = new System.Drawing.Font("Arial", 8F);
            this.export_Button.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.export_Button.Location = new System.Drawing.Point(420, 31);
            this.export_Button.Name = "export_Button";
            this.export_Button.Size = new System.Drawing.Size(75, 23);
            this.export_Button.TabIndex = 9;
            this.export_Button.Text = "Exportar";
            this.export_Button.UseVisualStyleBackColor = false;
            this.export_Button.Click += new System.EventHandler(this.export_Button_Click);
            // 
            // tB_OE
            // 
            this.tB_OE.Font = new System.Drawing.Font("Arial", 9F);
            this.tB_OE.Location = new System.Drawing.Point(182, 6);
            this.tB_OE.Name = "tB_OE";
            this.tB_OE.Size = new System.Drawing.Size(150, 21);
            this.tB_OE.TabIndex = 7;
            // 
            // b_ReadCaract
            // 
            this.b_ReadCaract.BackColor = System.Drawing.SystemColors.HighlightText;
            this.b_ReadCaract.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.b_ReadCaract.Cursor = System.Windows.Forms.Cursors.Hand;
            this.b_ReadCaract.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.b_ReadCaract.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.b_ReadCaract.Font = new System.Drawing.Font("Arial", 8F);
            this.b_ReadCaract.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.b_ReadCaract.Location = new System.Drawing.Point(99, 3);
            this.b_ReadCaract.Name = "b_ReadCaract";
            this.b_ReadCaract.Size = new System.Drawing.Size(75, 23);
            this.b_ReadCaract.TabIndex = 6;
            this.b_ReadCaract.Text = "Leer OE";
            this.b_ReadCaract.UseVisualStyleBackColor = false;
            this.b_ReadCaract.Click += new System.EventHandler(this.BRead_Click);
            // 
            // b_EntradaManual
            // 
            this.b_EntradaManual.BackColor = System.Drawing.SystemColors.HighlightText;
            this.b_EntradaManual.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.b_EntradaManual.Cursor = System.Windows.Forms.Cursors.Hand;
            this.b_EntradaManual.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.b_EntradaManual.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.b_EntradaManual.Font = new System.Drawing.Font("Arial", 8F);
            this.b_EntradaManual.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.b_EntradaManual.Location = new System.Drawing.Point(339, 3);
            this.b_EntradaManual.Name = "b_EntradaManual";
            this.b_EntradaManual.Size = new System.Drawing.Size(75, 23);
            this.b_EntradaManual.TabIndex = 5;
            this.b_EntradaManual.Text = "Manual";
            this.b_EntradaManual.UseVisualStyleBackColor = false;
            // 
            // progressBar_Draw
            // 
            this.progressBar_Draw.Location = new System.Drawing.Point(182, 35);
            this.progressBar_Draw.Margin = new System.Windows.Forms.Padding(2);
            this.progressBar_Draw.Name = "progressBar_Draw";
            this.progressBar_Draw.Size = new System.Drawing.Size(149, 20);
            this.progressBar_Draw.Step = 1;
            this.progressBar_Draw.TabIndex = 8;
            // 
            // b_Draw
            // 
            this.b_Draw.BackColor = System.Drawing.SystemColors.HighlightText;
            this.b_Draw.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.b_Draw.Cursor = System.Windows.Forms.Cursors.Hand;
            this.b_Draw.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.b_Draw.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.b_Draw.Font = new System.Drawing.Font("Arial", 8F);
            this.b_Draw.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.b_Draw.Location = new System.Drawing.Point(99, 31);
            this.b_Draw.Name = "b_Draw";
            this.b_Draw.Size = new System.Drawing.Size(75, 23);
            this.b_Draw.TabIndex = 3;
            this.b_Draw.Text = "Dibujar";
            this.b_Draw.UseVisualStyleBackColor = false;
            this.b_Draw.Click += new System.EventHandler(this.BDraw_Click);
            // 
            // b_Cables
            // 
            this.b_Cables.BackColor = System.Drawing.SystemColors.HighlightText;
            this.b_Cables.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.b_Cables.Cursor = System.Windows.Forms.Cursors.Hand;
            this.b_Cables.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.b_Cables.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.b_Cables.Font = new System.Drawing.Font("Arial", 8F);
            this.b_Cables.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.b_Cables.Location = new System.Drawing.Point(511, 31);
            this.b_Cables.Name = "b_Cables";
            this.b_Cables.Size = new System.Drawing.Size(75, 23);
            this.b_Cables.TabIndex = 4;
            this.b_Cables.Text = "Test";
            this.b_Cables.UseVisualStyleBackColor = false;
            this.b_Cables.Click += new System.EventHandler(this.BTest_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.HighlightText;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button1.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Arial", 8F);
            this.button1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button1.Location = new System.Drawing.Point(338, 31);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 5;
            this.button1.Text = "Calcular";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.BCalc_Click);
            // 
            // panelApp
            // 
            this.panelApp.BackColor = System.Drawing.SystemColors.Window;
            this.panelApp.Location = new System.Drawing.Point(129, 60);
            this.panelApp.Name = "panelApp";
            this.panelApp.Size = new System.Drawing.Size(1323, 894);
            this.panelApp.TabIndex = 11;
            // 
            // MenuVertical
            // 
            this.MenuVertical.BackColor = System.Drawing.Color.Black;
            this.MenuVertical.Controls.Add(this.buttonFormHC);
            this.MenuVertical.Controls.Add(this.buttonFormGEC);
            this.MenuVertical.Controls.Add(this.logoTKE);
            this.MenuVertical.Controls.Add(this.buttonFormCaracComer);
            this.MenuVertical.Controls.Add(this.buttonFormCaracIng);
            this.MenuVertical.Controls.Add(this.panel1);
            this.MenuVertical.Dock = System.Windows.Forms.DockStyle.Left;
            this.MenuVertical.ForeColor = System.Drawing.Color.White;
            this.MenuVertical.Location = new System.Drawing.Point(0, 0);
            this.MenuVertical.Name = "MenuVertical";
            this.MenuVertical.Size = new System.Drawing.Size(129, 958);
            this.MenuVertical.TabIndex = 10;
            // 
            // buttonFormHC
            // 
            this.buttonFormHC.BackColor = System.Drawing.Color.Black;
            this.buttonFormHC.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFormHC.FlatAppearance.BorderSize = 0;
            this.buttonFormHC.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.buttonFormHC.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFormHC.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.buttonFormHC.ForeColor = System.Drawing.Color.White;
            this.buttonFormHC.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonFormHC.Location = new System.Drawing.Point(0, 234);
            this.buttonFormHC.Name = "buttonFormHC";
            this.buttonFormHC.Size = new System.Drawing.Size(130, 35);
            this.buttonFormHC.TabIndex = 18;
            this.buttonFormHC.Text = "HC";
            this.buttonFormHC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonFormHC.UseVisualStyleBackColor = false;
            // 
            // buttonFormGEC
            // 
            this.buttonFormGEC.BackColor = System.Drawing.Color.Black;
            this.buttonFormGEC.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFormGEC.FlatAppearance.BorderSize = 0;
            this.buttonFormGEC.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.buttonFormGEC.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFormGEC.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.buttonFormGEC.ForeColor = System.Drawing.Color.White;
            this.buttonFormGEC.Image = ((System.Drawing.Image)(resources.GetObject("buttonFormGEC.Image")));
            this.buttonFormGEC.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonFormGEC.Location = new System.Drawing.Point(0, 184);
            this.buttonFormGEC.Name = "buttonFormGEC";
            this.buttonFormGEC.Size = new System.Drawing.Size(130, 35);
            this.buttonFormGEC.TabIndex = 17;
            this.buttonFormGEC.Text = "GEC";
            this.buttonFormGEC.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonFormGEC.UseVisualStyleBackColor = false;
            this.buttonFormGEC.Click += new System.EventHandler(this.BGEC_Click);
            // 
            // logoTKE
            // 
            this.logoTKE.BackColor = System.Drawing.Color.White;
            this.logoTKE.Image = ((System.Drawing.Image)(resources.GetObject("logoTKE.Image")));
            this.logoTKE.Location = new System.Drawing.Point(0, 0);
            this.logoTKE.Name = "logoTKE";
            this.logoTKE.Size = new System.Drawing.Size(130, 59);
            this.logoTKE.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.logoTKE.TabIndex = 13;
            this.logoTKE.TabStop = false;
            // 
            // buttonFormCaracComer
            // 
            this.buttonFormCaracComer.BackColor = System.Drawing.Color.Black;
            this.buttonFormCaracComer.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFormCaracComer.FlatAppearance.BorderSize = 0;
            this.buttonFormCaracComer.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.buttonFormCaracComer.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFormCaracComer.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.buttonFormCaracComer.ForeColor = System.Drawing.Color.White;
            this.buttonFormCaracComer.Image = ((System.Drawing.Image)(resources.GetObject("buttonFormCaracComer.Image")));
            this.buttonFormCaracComer.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonFormCaracComer.Location = new System.Drawing.Point(0, 85);
            this.buttonFormCaracComer.Margin = new System.Windows.Forms.Padding(0);
            this.buttonFormCaracComer.Name = "buttonFormCaracComer";
            this.buttonFormCaracComer.Size = new System.Drawing.Size(129, 35);
            this.buttonFormCaracComer.TabIndex = 0;
            this.buttonFormCaracComer.Text = "      COMERCIAL";
            this.buttonFormCaracComer.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonFormCaracComer.UseVisualStyleBackColor = false;
            this.buttonFormCaracComer.Click += new System.EventHandler(this.BCaracComer_Click);
            // 
            // buttonFormCaracIng
            // 
            this.buttonFormCaracIng.BackColor = System.Drawing.Color.Black;
            this.buttonFormCaracIng.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonFormCaracIng.FlatAppearance.BorderSize = 0;
            this.buttonFormCaracIng.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.buttonFormCaracIng.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.buttonFormCaracIng.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Bold);
            this.buttonFormCaracIng.ForeColor = System.Drawing.Color.White;
            this.buttonFormCaracIng.Image = ((System.Drawing.Image)(resources.GetObject("buttonFormCaracIng.Image")));
            this.buttonFormCaracIng.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonFormCaracIng.Location = new System.Drawing.Point(0, 135);
            this.buttonFormCaracIng.Name = "buttonFormCaracIng";
            this.buttonFormCaracIng.Size = new System.Drawing.Size(130, 35);
            this.buttonFormCaracIng.TabIndex = 1;
            this.buttonFormCaracIng.Text = "INGENIERIA";
            this.buttonFormCaracIng.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonFormCaracIng.UseVisualStyleBackColor = false;
            this.buttonFormCaracIng.Click += new System.EventHandler(this.BCaracIngClick);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(171, 59);
            this.panel1.TabIndex = 16;
            // 
            // PanelTitulo
            // 
            this.PanelTitulo.BackColor = System.Drawing.Color.Black;
            this.PanelTitulo.Controls.Add(this.cBox_Obra);
            this.PanelTitulo.Controls.Add(this.ImportButton);
            this.PanelTitulo.Controls.Add(this.export_Button);
            this.PanelTitulo.Controls.Add(this.progressBar_Draw);
            this.PanelTitulo.Controls.Add(this.b_EntradaManual);
            this.PanelTitulo.Controls.Add(this.b_Draw);
            this.PanelTitulo.Controls.Add(this.tB_OE);
            this.PanelTitulo.Controls.Add(this.button1);
            this.PanelTitulo.Controls.Add(this.b_ReadCaract);
            this.PanelTitulo.Controls.Add(this.b_Cables);
            this.PanelTitulo.Dock = System.Windows.Forms.DockStyle.Top;
            this.PanelTitulo.Location = new System.Drawing.Point(129, 0);
            this.PanelTitulo.Name = "PanelTitulo";
            this.PanelTitulo.Size = new System.Drawing.Size(1314, 59);
            this.PanelTitulo.TabIndex = 12;
            // 
            // cBox_Obra
            // 
            this.cBox_Obra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cBox_Obra.FormattingEnabled = true;
            this.cBox_Obra.Location = new System.Drawing.Point(511, 3);
            this.cBox_Obra.Name = "cBox_Obra";
            this.cBox_Obra.Size = new System.Drawing.Size(157, 21);
            this.cBox_Obra.TabIndex = 11;
            // 
            // ImportButton
            // 
            this.ImportButton.BackColor = System.Drawing.SystemColors.HighlightText;
            this.ImportButton.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ImportButton.Cursor = System.Windows.Forms.Cursors.Hand;
            this.ImportButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.ImportButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.White;
            this.ImportButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(237)))), ((int)(((byte)(110)))), ((int)(((byte)(17)))));
            this.ImportButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ImportButton.Font = new System.Drawing.Font("Arial", 8F);
            this.ImportButton.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ImportButton.Location = new System.Drawing.Point(420, 2);
            this.ImportButton.Name = "ImportButton";
            this.ImportButton.Size = new System.Drawing.Size(75, 23);
            this.ImportButton.TabIndex = 9;
            this.ImportButton.Text = "Importar";
            this.ImportButton.UseVisualStyleBackColor = false;
            this.ImportButton.Click += new System.EventHandler(this.ImportButton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // tbStatus
            // 
            this.tbStatus.BackColor = System.Drawing.SystemColors.HighlightText;
            this.tbStatus.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tbStatus.Location = new System.Drawing.Point(129, 958);
            this.tbStatus.Name = "tbStatus";
            this.tbStatus.ReadOnly = true;
            this.tbStatus.Size = new System.Drawing.Size(1323, 13);
            this.tbStatus.TabIndex = 0;
            this.tbStatus.Text = "-";
            // 
            // Configurador
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.HighlightText;
            this.ClientSize = new System.Drawing.Size(1443, 958);
            this.Controls.Add(this.tbStatus);
            this.Controls.Add(this.PanelTitulo);
            this.Controls.Add(this.panelApp);
            this.Controls.Add(this.MenuVertical);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.MinimumSize = new System.Drawing.Size(1050, 664);
            this.Name = "Configurador";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Configurador";
            this.MenuVertical.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.logoTKE)).EndInit();
            this.PanelTitulo.ResumeLayout(false);
            this.PanelTitulo.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button export_Button;
        public System.Windows.Forms.TextBox tB_OE;
        private System.Windows.Forms.Button b_ReadCaract;
        private System.Windows.Forms.Button b_EntradaManual;
        private System.Windows.Forms.ProgressBar progressBar_Draw;
        private System.Windows.Forms.Button b_Draw;
        private System.Windows.Forms.Button b_Cables;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panelApp;
        private System.Windows.Forms.Button buttonFormCaracComer;
        private System.Windows.Forms.Button buttonFormCaracIng;
        private System.Windows.Forms.Panel MenuVertical;
        private System.Windows.Forms.Panel PanelTitulo;
        private System.Windows.Forms.PictureBox logoTKE;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button buttonFormGEC;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ComboBox cBox_Obra;
        private System.Windows.Forms.Button ImportButton;
        private System.Windows.Forms.Button buttonFormHC;
        public System.Windows.Forms.TextBox tbStatus;
    }
}