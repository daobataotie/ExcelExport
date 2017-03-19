using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    public class ExcelHelper
    {
        private dynamic xlApp = null;
        private dynamic workbook = null;
        private int processId = 0;

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);

        public ExcelHelper(string fileName)
        {
            Type t = Type.GetTypeFromProgID("EXCEL.Application");
            if (t == null)
            {
                throw new Exception("本機沒有安裝Excel");
            }
            this.xlApp = Activator.CreateInstance(t);

            GetWindowThreadProcessId(((IntPtr)this.xlApp.Hwnd), out this.processId);

            this.xlApp.Workbooks.Open(fileName);
            this.workbook = this.xlApp.ActiveWorkbook;
        }

        public void Close()
        {
            this.workbook.Close();
            if (this.processId != 0)
            {
                Process p = Process.GetProcessById(this.processId);
                p.Kill();
            }
        }

        public void Save()
        {
            this.workbook.Save();
        }

        public int GetColumnIndex(string endColumn)
        {
            string columnRef = "ABCDEFGHIGKLMNOPQRSTUVWXYZ";
            return columnRef.IndexOf(endColumn) + 1;
        }

        public object GetDate(string endColumn)
        {
            dynamic sheet = this.xlApp.ActiveSheet;

            int rowCount = sheet.UsedRange.Rows.Count;
            if (rowCount == 0)
            {
                throw new Exception("Excel數據錯誤");
            }
            dynamic range = sheet.Range["A1", string.Format("{0}{1}", endColumn, rowCount)];

            return range.Value2;
        }

        public void SetCellValue(int rowIndex, int colIndex, object value)
        {
            dynamic sheet = this.xlApp.ActiveSheet;
            dynamic cell = sheet.Cells[rowIndex, colIndex];

            cell.Value2 = value;
        }

        public void SetRangeFormat(string startIndex, string endIndex, object format)
        {
            dynamic sheet = this.xlApp.ActiveSheet;
            dynamic range = sheet.Range[startIndex, endIndex];

            range.NumberFormatLocal = format;
        }

    }
}
