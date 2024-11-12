namespace EPLAN_API.Forms
{
    partial class NumeroEsquema
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
            this.label1 = new System.Windows.Forms.Label();
            this.tB_Drawing_Number = new System.Windows.Forms.TextBox();
            this.checkBox_NewDrawing = new System.Windows.Forms.CheckBox();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.buttonAcept = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(59, 58);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(136, 36);
            this.label1.TabIndex = 1;
            this.label1.Text = "0-55.1-3.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // tB_Drawing_Number
            // 
            this.tB_Drawing_Number.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tB_Drawing_Number.Location = new System.Drawing.Point(199, 54);
            this.tB_Drawing_Number.Margin = new System.Windows.Forms.Padding(4);
            this.tB_Drawing_Number.Name = "tB_Drawing_Number";
            this.tB_Drawing_Number.Size = new System.Drawing.Size(187, 41);
            this.tB_Drawing_Number.TabIndex = 2;
            // 
            // checkBox_NewDrawing
            // 
            this.checkBox_NewDrawing.AutoSize = true;
            this.checkBox_NewDrawing.Location = new System.Drawing.Point(65, 123);
            this.checkBox_NewDrawing.Margin = new System.Windows.Forms.Padding(4);
            this.checkBox_NewDrawing.Name = "checkBox_NewDrawing";
            this.checkBox_NewDrawing.Size = new System.Drawing.Size(182, 20);
            this.checkBox_NewDrawing.TabIndex = 3;
            this.checkBox_NewDrawing.Text = "Generar Nuevo Esquema";
            this.checkBox_NewDrawing.UseVisualStyleBackColor = true;
            this.checkBox_NewDrawing.CheckStateChanged += new System.EventHandler(this.check_NewDrawin_Clik);
            // 
            // button_Cancel
            // 
            this.button_Cancel.Location = new System.Drawing.Point(287, 168);
            this.button_Cancel.Margin = new System.Windows.Forms.Padding(4);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(100, 28);
            this.button_Cancel.TabIndex = 6;
            this.button_Cancel.Text = "Cancelar";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // buttonAcept
            // 
            this.buttonAcept.Location = new System.Drawing.Point(65, 168);
            this.buttonAcept.Margin = new System.Windows.Forms.Padding(4);
            this.buttonAcept.Name = "buttonAcept";
            this.buttonAcept.Size = new System.Drawing.Size(100, 28);
            this.buttonAcept.TabIndex = 5;
            this.buttonAcept.Text = "Aceptar";
            this.buttonAcept.UseVisualStyleBackColor = true;
            this.buttonAcept.Click += new System.EventHandler(this.buttonAcept_Click);
            // 
            // NumeroEsquema
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 209);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.buttonAcept);
            this.Controls.Add(this.checkBox_NewDrawing);
            this.Controls.Add(this.tB_Drawing_Number);
            this.Controls.Add(this.label1);
            this.Name = "NumeroEsquema";
            this.Text = "NumeroEsquema";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tB_Drawing_Number;
        private System.Windows.Forms.CheckBox checkBox_NewDrawing;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.Button buttonAcept;
    }
}