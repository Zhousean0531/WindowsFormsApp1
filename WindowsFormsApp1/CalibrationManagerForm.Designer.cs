namespace WindowsFormsApp1
{
    partial class CalibrationManagerForm
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

        private void InitializeComponent()
        {
            this.CalibrationGrid = new System.Windows.Forms.DataGridView();
            this.LoadButton = new System.Windows.Forms.Button();
            this.SaveButton = new System.Windows.Forms.Button();

            ((System.ComponentModel.ISupportInitialize)(this.CalibrationGrid)).BeginInit();
            this.SuspendLayout();

            // 
            // CalibrationGrid
            // 
            this.CalibrationGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CalibrationGrid.Location = new System.Drawing.Point(12, 12);
            this.CalibrationGrid.Name = "CalibrationGrid";
            this.CalibrationGrid.RowTemplate.Height = 24;
            this.CalibrationGrid.Size = new System.Drawing.Size(560, 320);
            this.CalibrationGrid.TabIndex = 0;

            // 
            // LoadButton
            // 
            this.LoadButton.Location = new System.Drawing.Point(12, 350);
            this.LoadButton.Name = "LoadButton";
            this.LoadButton.Size = new System.Drawing.Size(100, 35);
            this.LoadButton.TabIndex = 1;
            this.LoadButton.Text = "重新讀取";
            this.LoadButton.UseVisualStyleBackColor = true;

            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(130, 350);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(100, 35);
            this.SaveButton.TabIndex = 2;
            this.SaveButton.Text = "儲存";
            this.SaveButton.UseVisualStyleBackColor = true;

            // 
            // CalibrationManagerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(584, 401);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.LoadButton);
            this.Controls.Add(this.CalibrationGrid);
            this.Name = "CalibrationManagerForm";
            this.Text = "儀器校正日期管理";

            ((System.ComponentModel.ISupportInitialize)(this.CalibrationGrid)).EndInit();
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.DataGridView CalibrationGrid;
        private System.Windows.Forms.Button LoadButton;
        private System.Windows.Forms.Button SaveButton;
    }
}