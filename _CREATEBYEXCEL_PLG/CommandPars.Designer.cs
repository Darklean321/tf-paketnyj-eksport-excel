namespace CREATEBYEXCEL_PLG {
    partial class ComParams {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
        if(disposing && (components != null)) {
        components.Dispose();
        }
        base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ComParams));
            this.pEXSTEP = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.pEXDOCs = new System.Windows.Forms.CheckBox();
            this.pEXPDF = new System.Windows.Forms.CheckBox();
            this.bCancel = new System.Windows.Forms.Button();
            this.pEXDXF = new System.Windows.Forms.CheckBox();
            this.bOK = new System.Windows.Forms.Button();
            this.pEXEXCEL = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pEXSTEP
            // 
            resources.ApplyResources(this.pEXSTEP, "pEXSTEP");
            this.pEXSTEP.Name = "pEXSTEP";
            this.pEXSTEP.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.pEXEXCEL);
            this.groupBox1.Controls.Add(this.pEXDOCs);
            this.groupBox1.Controls.Add(this.pEXPDF);
            this.groupBox1.Controls.Add(this.bCancel);
            this.groupBox1.Controls.Add(this.pEXDXF);
            this.groupBox1.Controls.Add(this.bOK);
            this.groupBox1.Controls.Add(this.pEXSTEP);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // pEXDOCs
            // 
            resources.ApplyResources(this.pEXDOCs, "pEXDOCs");
            this.pEXDOCs.Name = "pEXDOCs";
            this.pEXDOCs.UseVisualStyleBackColor = true;
            // 
            // pEXPDF
            // 
            resources.ApplyResources(this.pEXPDF, "pEXPDF");
            this.pEXPDF.Name = "pEXPDF";
            this.pEXPDF.UseVisualStyleBackColor = true;
            // 
            // bCancel
            // 
            this.bCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.bCancel, "bCancel");
            this.bCancel.Name = "bCancel";
            this.bCancel.UseVisualStyleBackColor = true;
            this.bCancel.Click += new System.EventHandler(this.bCancel_Click);
            // 
            // pEXDXF
            // 
            resources.ApplyResources(this.pEXDXF, "pEXDXF");
            this.pEXDXF.Name = "pEXDXF";
            this.pEXDXF.UseVisualStyleBackColor = true;
            // 
            // bOK
            // 
            this.bOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            resources.ApplyResources(this.bOK, "bOK");
            this.bOK.Name = "bOK";
            this.bOK.UseVisualStyleBackColor = true;
            // 
            // pEXEXCEL
            // 
            resources.ApplyResources(this.pEXEXCEL, "pEXEXCEL");
            this.pEXEXCEL.Name = "pEXEXCEL";
            this.pEXEXCEL.UseVisualStyleBackColor = true;
            // 
            // ComParams
            // 
            this.AcceptButton = this.bOK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.bCancel;
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ComParams";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button bOK;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.CheckBox pEXDXF;
        private System.Windows.Forms.CheckBox pEXSTEP;
        private System.Windows.Forms.CheckBox pEXPDF;
        private System.Windows.Forms.CheckBox pEXDOCs;
        private System.Windows.Forms.CheckBox pEXEXCEL;
    }
}