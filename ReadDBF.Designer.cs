namespace ReadDBF
{
    partial class ReadDBF
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.BtnOpen = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.CmboxSearch = new System.Windows.Forms.ComboBox();
            this.TxtboxSearch = new System.Windows.Forms.TextBox();
            this.BtnSearch = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.BtnSearch);
            this.panel1.Controls.Add(this.TxtboxSearch);
            this.panel1.Controls.Add(this.CmboxSearch);
            this.panel1.Controls.Add(this.BtnSave);
            this.panel1.Controls.Add(this.BtnOpen);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(800, 40);
            this.panel1.TabIndex = 0;
            // 
            // BtnOpen
            // 
            this.BtnOpen.Location = new System.Drawing.Point(30, 9);
            this.BtnOpen.Name = "BtnOpen";
            this.BtnOpen.Size = new System.Drawing.Size(111, 23);
            this.BtnOpen.TabIndex = 0;
            this.BtnOpen.Text = "打开DBF文件";
            this.BtnOpen.UseVisualStyleBackColor = true;
            // 
            // BtnSave
            // 
            this.BtnSave.Location = new System.Drawing.Point(164, 9);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(111, 23);
            this.BtnSave.TabIndex = 1;
            this.BtnSave.Text = "保存Excel文件";
            this.BtnSave.UseVisualStyleBackColor = true;
            // 
            // CmboxSearch
            // 
            this.CmboxSearch.FormattingEnabled = true;
            this.CmboxSearch.Location = new System.Drawing.Point(298, 10);
            this.CmboxSearch.Name = "CmboxSearch";
            this.CmboxSearch.Size = new System.Drawing.Size(169, 20);
            this.CmboxSearch.TabIndex = 2;
            // 
            // TxtboxSearch
            // 
            this.TxtboxSearch.Location = new System.Drawing.Point(490, 10);
            this.TxtboxSearch.Name = "TxtboxSearch";
            this.TxtboxSearch.Size = new System.Drawing.Size(169, 21);
            this.TxtboxSearch.TabIndex = 3;
            // 
            // BtnSearch
            // 
            this.BtnSearch.Location = new System.Drawing.Point(682, 9);
            this.BtnSearch.Name = "BtnSearch";
            this.BtnSearch.Size = new System.Drawing.Size(75, 23);
            this.BtnSearch.TabIndex = 4;
            this.BtnSearch.Text = "查询";
            this.BtnSearch.UseVisualStyleBackColor = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 40);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(800, 410);
            this.dataGridView1.TabIndex = 1;
            // 
            // ReadDBF
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel1);
            this.Name = "ReadDBF";
            this.Text = "ReadDBF";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button BtnSearch;
        private System.Windows.Forms.TextBox TxtboxSearch;
        private System.Windows.Forms.ComboBox CmboxSearch;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.Button BtnOpen;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

