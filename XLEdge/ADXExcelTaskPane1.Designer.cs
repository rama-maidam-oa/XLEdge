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
            this.SuspendLayout();
            // 
            // ADXExcelTaskPane1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(600, 547);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXExcelTaskPane1";
            this.Text = "";
            this.ADXBeforeTaskPaneShow += new AddinExpress.XL.ADXBeforeTaskPaneShowEventHandler(this.ADXExcelTaskPane1_ADXBeforeTaskPaneShow);
            this.ADXAfterTaskPaneShow += new AddinExpress.XL.ADXAfterTaskPaneShowEventHandler(this.ADXExcelTaskPane1_ADXAfterTaskPaneShow);
            this.ADXCloseButtonClick += new AddinExpress.XL.ADXCloseButtonClickEventHandler(this.ADXExcelTaskPane1_ADXCloseButtonClick);
            this.ResumeLayout(false);

        }
        #endregion
    }
}
