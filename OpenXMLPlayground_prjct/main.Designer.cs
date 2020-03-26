namespace OpenXMLPlayground_prjct
{
    partial class Main
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCreateTestWordDocument = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCreateTestWordDocument
            // 
            this.btnCreateTestWordDocument.Location = new System.Drawing.Point(12, 12);
            this.btnCreateTestWordDocument.Name = "btnCreateTestWordDocument";
            this.btnCreateTestWordDocument.Size = new System.Drawing.Size(143, 44);
            this.btnCreateTestWordDocument.TabIndex = 0;
            this.btnCreateTestWordDocument.Text = "Create test document";
            this.btnCreateTestWordDocument.UseVisualStyleBackColor = true;
            this.btnCreateTestWordDocument.Click += new System.EventHandler(this.btnCreateTestWordDocument_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(12, 62);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(143, 41);
            this.btnExcel.TabIndex = 1;
            this.btnExcel.Text = "Create Excel test document";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(336, 115);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnCreateTestWordDocument);
            this.Name = "Main";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCreateTestWordDocument;
        private System.Windows.Forms.Button btnExcel;
    }
}

