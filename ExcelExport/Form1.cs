using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelExport
{
    public partial class Form1 : Form
    {
        private ExcelOperation eo = new ExcelOperation();
        public Form1()
        {
            InitializeComponent();

            base.StartPosition = FormStartPosition.CenterScreen;
            this.date_Date.Value = DateTime.Now;
        }

        private void btn_Import_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Title = "請選擇Excel文件";
                ofd.Filter = "Excel(*.xls)|*.xlsx";
                ofd.Multiselect = false;
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    this.txt_ImportFileName.Text = ofd.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txt_ImportFileName.Text))
                {
                    MessageBox.Show("請先導入Excel文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                else
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Title = "請選擇保存路徑";
                    sfd.Filter = "Excel(*.xlsx)|*.xlsx";
                    sfd.AddExtension = true;
                    sfd.CheckPathExists = true;
                    sfd.DefaultExt = "xlsx";
                    if (sfd.ShowDialog(this) == DialogResult.OK)
                    {
                        this.txt_ExportFileName.Text = sfd.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txt_ExportFileName.Text))
                {
                    MessageBox.Show("請先選擇導出Excel文件的路徑！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
                Model model = this.eo.GetExcelDate<Model>(this.txt_ImportFileName.Text, "K");

                model.Company = this.txt_Company.Text;
                model.TestMethod = this.txt_TestMethod.Text;
                model.Position = this.txt_Position.Text;
                model.Manufacturer = this.txt_Manufacturer.Text;
                model.TestDate = this.date_Date.Value;
                model.ModelValue = txt_Model.Text;
                model.TestedBy = txt_TestedBy.Text;

                if (model.CanRun)
                {
                    string templateName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelTemplate\\Template.xlsx");
                    File.Copy(templateName, this.txt_ExportFileName.Text, true);
                    this.eo.WriteExcel(this.txt_ExportFileName.Text, model);
                    MessageBox.Show("保存成功！", "提示");
                }
                else
                {
                    MessageBox.Show("請先完整填寫報表資料！", "提示");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
