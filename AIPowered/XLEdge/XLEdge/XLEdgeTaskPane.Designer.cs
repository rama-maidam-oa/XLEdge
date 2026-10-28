namespace XLEdge
{
    partial class XLEdgeTaskPane
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code
        private void InitializeComponent()
        {
            this.SuspendLayout();
            //
            // XLEdgeTaskPane
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(688, 916);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "XLEdgeTaskPane";
            this.Text = "";
            this.ADXBeforeTaskPaneShow += new AddinExpress.XL.ADXBeforeTaskPaneShowEventHandler(this.XLEdgeTaskPane_ADXBeforeTaskPaneShow);
            this.ADXAfterTaskPaneShow += new AddinExpress.XL.ADXAfterTaskPaneShowEventHandler(this.XLEdgeTaskPane_ADXAfterTaskPaneShow);
            this.ADXCloseButtonClick += new AddinExpress.XL.ADXCloseButtonClickEventHandler(this.XLEdgeTaskPane_ADXCloseButtonClick);
            this.ResumeLayout(false);
        }
        #endregion
    }
}
