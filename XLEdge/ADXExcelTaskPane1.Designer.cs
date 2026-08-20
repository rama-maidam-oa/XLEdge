namespace XLEdge
{
    partial class ADXExcelTaskPane1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
  
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
  
        #region Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXExcelTaskPane1));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.WebCtrl = new Microsoft.Web.WebView2.WinForms.WebView2();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.WebCtrl)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.WebCtrl, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(688, 916);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // WebCtrl
            // 
            this.WebCtrl.AllowExternalDrop = true;
            this.WebCtrl.CreationProperties = null;
            this.WebCtrl.DefaultBackgroundColor = System.Drawing.Color.White;
            this.WebCtrl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.WebCtrl.Location = new System.Drawing.Point(3, 3);
            this.WebCtrl.Name = "WebCtrl";
            this.WebCtrl.Size = new System.Drawing.Size(682, 910);
            this.WebCtrl.TabIndex = 0;
            this.WebCtrl.ZoomFactor = 1D;
            // 
            // ADXExcelTaskPane1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(688, 916);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXExcelTaskPane1";
            this.Text = "";
            this.ADXBeforeTaskPaneShow += new AddinExpress.XL.ADXBeforeTaskPaneShowEventHandler(this.ADXExcelTaskPane1_ADXBeforeTaskPaneShow);
            this.ADXAfterTaskPaneShow += new AddinExpress.XL.ADXAfterTaskPaneShowEventHandler(this.ADXExcelTaskPane1_ADXAfterTaskPaneShow);
            this.ADXCloseButtonClick += new AddinExpress.XL.ADXCloseButtonClickEventHandler(this.ADXExcelTaskPane1_ADXCloseButtonClick);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.WebCtrl)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        public Microsoft.Web.WebView2.WinForms.WebView2 WebCtrl;
    }
}
