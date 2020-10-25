namespace FluUserUploadData
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
            this.btnFluPatient = new System.Windows.Forms.Button();
            this.btnOutpatient = new System.Windows.Forms.Button();
            this.btnOutResult = new System.Windows.Forms.Button();
            this.btnDeath = new System.Windows.Forms.Button();
            this.btnDrug = new System.Windows.Forms.Button();
            this.btnLis = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpBegin = new System.Windows.Forms.DateTimePicker();
            this.dtpEnd = new System.Windows.Forms.DateTimePicker();
            this.btnSetting = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnFluPatient
            // 
            this.btnFluPatient.Location = new System.Drawing.Point(30, 113);
            this.btnFluPatient.Name = "btnFluPatient";
            this.btnFluPatient.Size = new System.Drawing.Size(210, 28);
            this.btnFluPatient.TabIndex = 0;
            this.btnFluPatient.Text = "门急诊和在院流感病例";
            this.btnFluPatient.UseVisualStyleBackColor = true;
            this.btnFluPatient.Click += new System.EventHandler(this.btnFluPatient_Click);
            // 
            // btnOutpatient
            // 
            this.btnOutpatient.Location = new System.Drawing.Point(30, 154);
            this.btnOutpatient.Name = "btnOutpatient";
            this.btnOutpatient.Size = new System.Drawing.Size(210, 28);
            this.btnOutpatient.TabIndex = 1;
            this.btnOutpatient.Text = "出院流感病例数据";
            this.btnOutpatient.UseVisualStyleBackColor = true;
            this.btnOutpatient.Click += new System.EventHandler(this.btnOutpatient_Click);
            // 
            // btnOutResult
            // 
            this.btnOutResult.Location = new System.Drawing.Point(30, 195);
            this.btnOutResult.Name = "btnOutResult";
            this.btnOutResult.Size = new System.Drawing.Size(210, 28);
            this.btnOutResult.TabIndex = 2;
            this.btnOutResult.Text = "出院小结数据";
            this.btnOutResult.UseVisualStyleBackColor = true;
            this.btnOutResult.Click += new System.EventHandler(this.btnOutResult_Click);
            // 
            // btnDeath
            // 
            this.btnDeath.Location = new System.Drawing.Point(30, 236);
            this.btnDeath.Name = "btnDeath";
            this.btnDeath.Size = new System.Drawing.Size(210, 28);
            this.btnDeath.TabIndex = 3;
            this.btnDeath.Text = "死亡记录数据";
            this.btnDeath.UseVisualStyleBackColor = true;
            this.btnDeath.Click += new System.EventHandler(this.btnDeath_Click);
            // 
            // btnDrug
            // 
            this.btnDrug.Location = new System.Drawing.Point(30, 277);
            this.btnDrug.Name = "btnDrug";
            this.btnDrug.Size = new System.Drawing.Size(210, 28);
            this.btnDrug.TabIndex = 4;
            this.btnDrug.Text = "用药记录数据";
            this.btnDrug.UseVisualStyleBackColor = true;
            this.btnDrug.Click += new System.EventHandler(this.btnDrug_Click);
            // 
            // btnLis
            // 
            this.btnLis.Location = new System.Drawing.Point(30, 318);
            this.btnLis.Name = "btnLis";
            this.btnLis.Size = new System.Drawing.Size(210, 28);
            this.btnLis.TabIndex = 5;
            this.btnLis.Text = "检验记录数据";
            this.btnLis.UseVisualStyleBackColor = true;
            this.btnLis.Click += new System.EventHandler(this.btnLis_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 7;
            this.label1.Text = "开始时间";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(229, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 9;
            this.label2.Text = "结束时间";
            // 
            // dtpBegin
            // 
            this.dtpBegin.Location = new System.Drawing.Point(87, 20);
            this.dtpBegin.Name = "dtpBegin";
            this.dtpBegin.Size = new System.Drawing.Size(123, 21);
            this.dtpBegin.TabIndex = 13;
            // 
            // dtpEnd
            // 
            this.dtpEnd.Location = new System.Drawing.Point(288, 20);
            this.dtpEnd.Name = "dtpEnd";
            this.dtpEnd.Size = new System.Drawing.Size(123, 21);
            this.dtpEnd.TabIndex = 14;
            // 
            // btnSetting
            // 
            this.btnSetting.Location = new System.Drawing.Point(30, 72);
            this.btnSetting.Name = "btnSetting";
            this.btnSetting.Size = new System.Drawing.Size(210, 28);
            this.btnSetting.TabIndex = 15;
            this.btnSetting.Text = "设置导出时间段";
            this.btnSetting.UseVisualStyleBackColor = true;
            this.btnSetting.Click += new System.EventHandler(this.btnSetting_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(464, 359);
            this.Controls.Add(this.btnSetting);
            this.Controls.Add(this.dtpEnd);
            this.Controls.Add(this.dtpBegin);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnLis);
            this.Controls.Add(this.btnDrug);
            this.Controls.Add(this.btnDeath);
            this.Controls.Add(this.btnOutResult);
            this.Controls.Add(this.btnOutpatient);
            this.Controls.Add(this.btnFluPatient);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "流感病例数据上传";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnFluPatient;
        private System.Windows.Forms.Button btnOutpatient;
        private System.Windows.Forms.Button btnOutResult;
        private System.Windows.Forms.Button btnDeath;
        private System.Windows.Forms.Button btnDrug;
        private System.Windows.Forms.Button btnLis;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpBegin;
        private System.Windows.Forms.DateTimePicker dtpEnd;
        private System.Windows.Forms.Button btnSetting;
    }
}

