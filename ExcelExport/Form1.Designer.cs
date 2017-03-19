namespace ExcelExport
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.txt_ImportFileName = new System.Windows.Forms.TextBox();
            this.txt_ExportFileName = new System.Windows.Forms.TextBox();
            this.btn_Import = new System.Windows.Forms.Button();
            this.btn_Export = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_Company = new System.Windows.Forms.TextBox();
            this.txt_TestMethod = new System.Windows.Forms.TextBox();
            this.txt_Position = new System.Windows.Forms.TextBox();
            this.txt_Manufacturer = new System.Windows.Forms.TextBox();
            this.date_Date = new System.Windows.Forms.DateTimePicker();
            this.txt_Model = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txt_TestedBy = new System.Windows.Forms.TextBox();
            this.btn_OK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txt_ImportFileName
            // 
            this.txt_ImportFileName.Location = new System.Drawing.Point(108, 21);
            this.txt_ImportFileName.Name = "txt_ImportFileName";
            this.txt_ImportFileName.ReadOnly = true;
            this.txt_ImportFileName.Size = new System.Drawing.Size(357, 21);
            this.txt_ImportFileName.TabIndex = 0;
            // 
            // txt_ExportFileName
            // 
            this.txt_ExportFileName.Location = new System.Drawing.Point(108, 53);
            this.txt_ExportFileName.Name = "txt_ExportFileName";
            this.txt_ExportFileName.ReadOnly = true;
            this.txt_ExportFileName.Size = new System.Drawing.Size(357, 21);
            this.txt_ExportFileName.TabIndex = 1;
            // 
            // btn_Import
            // 
            this.btn_Import.Location = new System.Drawing.Point(471, 19);
            this.btn_Import.Name = "btn_Import";
            this.btn_Import.Size = new System.Drawing.Size(75, 23);
            this.btn_Import.TabIndex = 2;
            this.btn_Import.Text = "選擇文件";
            this.btn_Import.UseVisualStyleBackColor = true;
            this.btn_Import.Click += new System.EventHandler(this.btn_Import_Click);
            // 
            // btn_Export
            // 
            this.btn_Export.Location = new System.Drawing.Point(471, 53);
            this.btn_Export.Name = "btn_Export";
            this.btn_Export.Size = new System.Drawing.Size(75, 23);
            this.btn_Export.TabIndex = 3;
            this.btn_Export.Text = "選擇目錄";
            this.btn_Export.UseVisualStyleBackColor = true;
            this.btn_Export.Click += new System.EventHandler(this.btn_Export_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 12);
            this.label1.TabIndex = 4;
            this.label1.Text = "導入文件路徑：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 12);
            this.label2.TabIndex = 5;
            this.label2.Text = "導出文件路徑：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 98);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "公司名稱：";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(305, 130);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 7;
            this.label4.Text = "Position：";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 193);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(47, 12);
            this.label5.TabIndex = 8;
            this.label5.Text = "Model：";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(305, 161);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 12);
            this.label6.TabIndex = 9;
            this.label6.Text = "Date：";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(12, 161);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(89, 12);
            this.label7.TabIndex = 10;
            this.label7.Text = "Manufacturer：";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(12, 130);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(83, 12);
            this.label8.TabIndex = 11;
            this.label8.Text = "Test Method：";
            // 
            // txt_Company
            // 
            this.txt_Company.Location = new System.Drawing.Point(107, 95);
            this.txt_Company.Name = "txt_Company";
            this.txt_Company.Size = new System.Drawing.Size(439, 21);
            this.txt_Company.TabIndex = 12;
            this.txt_Company.Text = "Alan Safety Company";
            // 
            // txt_TestMethod
            // 
            this.txt_TestMethod.Location = new System.Drawing.Point(107, 127);
            this.txt_TestMethod.Name = "txt_TestMethod";
            this.txt_TestMethod.Size = new System.Drawing.Size(177, 21);
            this.txt_TestMethod.TabIndex = 13;
            // 
            // txt_Position
            // 
            this.txt_Position.Location = new System.Drawing.Point(376, 127);
            this.txt_Position.Name = "txt_Position";
            this.txt_Position.Size = new System.Drawing.Size(170, 21);
            this.txt_Position.TabIndex = 14;
            // 
            // txt_Manufacturer
            // 
            this.txt_Manufacturer.Location = new System.Drawing.Point(108, 158);
            this.txt_Manufacturer.Name = "txt_Manufacturer";
            this.txt_Manufacturer.Size = new System.Drawing.Size(177, 21);
            this.txt_Manufacturer.TabIndex = 15;
            // 
            // date_Date
            // 
            this.date_Date.Location = new System.Drawing.Point(376, 158);
            this.date_Date.Name = "date_Date";
            this.date_Date.Size = new System.Drawing.Size(170, 21);
            this.date_Date.TabIndex = 16;
            // 
            // txt_Model
            // 
            this.txt_Model.Location = new System.Drawing.Point(108, 193);
            this.txt_Model.Name = "txt_Model";
            this.txt_Model.Size = new System.Drawing.Size(177, 21);
            this.txt_Model.TabIndex = 17;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(305, 196);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(71, 12);
            this.label9.TabIndex = 18;
            this.label9.Text = "Tested by：";
            // 
            // txt_TestedBy
            // 
            this.txt_TestedBy.Location = new System.Drawing.Point(376, 190);
            this.txt_TestedBy.Name = "txt_TestedBy";
            this.txt_TestedBy.Size = new System.Drawing.Size(170, 21);
            this.txt_TestedBy.TabIndex = 19;
            // 
            // btn_OK
            // 
            this.btn_OK.Location = new System.Drawing.Point(252, 238);
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.Size = new System.Drawing.Size(75, 23);
            this.btn_OK.TabIndex = 20;
            this.btn_OK.Text = "導出";
            this.btn_OK.UseVisualStyleBackColor = true;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 273);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.txt_TestedBy);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.txt_Model);
            this.Controls.Add(this.date_Date);
            this.Controls.Add(this.txt_Manufacturer);
            this.Controls.Add(this.txt_Position);
            this.Controls.Add(this.txt_TestMethod);
            this.Controls.Add(this.txt_Company);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_Export);
            this.Controls.Add(this.btn_Import);
            this.Controls.Add(this.txt_ExportFileName);
            this.Controls.Add(this.txt_ImportFileName);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Excel導入導出";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_ImportFileName;
        private System.Windows.Forms.TextBox txt_ExportFileName;
        private System.Windows.Forms.Button btn_Import;
        private System.Windows.Forms.Button btn_Export;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_Company;
        private System.Windows.Forms.TextBox txt_TestMethod;
        private System.Windows.Forms.TextBox txt_Position;
        private System.Windows.Forms.TextBox txt_Manufacturer;
        private System.Windows.Forms.DateTimePicker date_Date;
        private System.Windows.Forms.TextBox txt_Model;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txt_TestedBy;
        private System.Windows.Forms.Button btn_OK;
    }
}

