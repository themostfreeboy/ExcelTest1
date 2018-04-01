namespace excel_test
{
    partial class Form1
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
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnTest1 = new System.Windows.Forms.Button();
            this.btnTest2 = new System.Windows.Forms.Button();
            this.txtTest2 = new System.Windows.Forms.TextBox();
            this.btnTest3 = new System.Windows.Forms.Button();
            this.dgvTest3 = new System.Windows.Forms.DataGridView();
            this.btnTest4 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTest3)).BeginInit();
            this.SuspendLayout();
            // 
            // btnTest1
            // 
            this.btnTest1.Location = new System.Drawing.Point(51, 26);
            this.btnTest1.Name = "btnTest1";
            this.btnTest1.Size = new System.Drawing.Size(75, 23);
            this.btnTest1.TabIndex = 0;
            this.btnTest1.Text = "测试程序1";
            this.btnTest1.UseVisualStyleBackColor = true;
            this.btnTest1.Click += new System.EventHandler(this.btnTest1_Click);
            // 
            // btnTest2
            // 
            this.btnTest2.Location = new System.Drawing.Point(51, 73);
            this.btnTest2.Name = "btnTest2";
            this.btnTest2.Size = new System.Drawing.Size(75, 23);
            this.btnTest2.TabIndex = 1;
            this.btnTest2.Text = "测试程序2";
            this.btnTest2.UseVisualStyleBackColor = true;
            this.btnTest2.Click += new System.EventHandler(this.btnTest2_Click);
            // 
            // txtTest2
            // 
            this.txtTest2.Location = new System.Drawing.Point(192, 74);
            this.txtTest2.Name = "txtTest2";
            this.txtTest2.Size = new System.Drawing.Size(207, 21);
            this.txtTest2.TabIndex = 2;
            // 
            // btnTest3
            // 
            this.btnTest3.Location = new System.Drawing.Point(51, 129);
            this.btnTest3.Name = "btnTest3";
            this.btnTest3.Size = new System.Drawing.Size(75, 23);
            this.btnTest3.TabIndex = 3;
            this.btnTest3.Text = "测试程序3";
            this.btnTest3.UseVisualStyleBackColor = true;
            this.btnTest3.Click += new System.EventHandler(this.btnTest3_Click);
            // 
            // dgvTest3
            // 
            this.dgvTest3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTest3.Location = new System.Drawing.Point(192, 115);
            this.dgvTest3.Name = "dgvTest3";
            this.dgvTest3.RowTemplate.Height = 23;
            this.dgvTest3.Size = new System.Drawing.Size(240, 150);
            this.dgvTest3.TabIndex = 4;
            // 
            // btnTest4
            // 
            this.btnTest4.Location = new System.Drawing.Point(51, 195);
            this.btnTest4.Name = "btnTest4";
            this.btnTest4.Size = new System.Drawing.Size(75, 23);
            this.btnTest4.TabIndex = 5;
            this.btnTest4.Text = "测试程序4";
            this.btnTest4.UseVisualStyleBackColor = true;
            this.btnTest4.Click += new System.EventHandler(this.btnTest4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(414, 256);
            this.Controls.Add(this.btnTest4);
            this.Controls.Add(this.dgvTest3);
            this.Controls.Add(this.btnTest3);
            this.Controls.Add(this.txtTest2);
            this.Controls.Add(this.btnTest2);
            this.Controls.Add(this.btnTest1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTest3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnTest1;
        private System.Windows.Forms.Button btnTest2;
        private System.Windows.Forms.TextBox txtTest2;
        private System.Windows.Forms.Button btnTest3;
        private System.Windows.Forms.DataGridView dgvTest3;
        private System.Windows.Forms.Button btnTest4;
    }
}

