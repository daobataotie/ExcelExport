using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    public class ExcelOperation
    {
        private double[] mean = new double[11];
        private double[] dev = new double[11];
        private double[] apv = new double[11];
        private double[] pnr = new double[9];
        private double[] C = new double[9];

        public T GetExcelDate<T>(string fileName, string endColumn) where T : new()
        {
            ExcelHelper excelHelper = null;
            T entity;
            try
            {
                excelHelper = new ExcelHelper(fileName);

                object[,] data = (object[,])excelHelper.GetDate(endColumn);
                List<Header> headerList = new List<Header>();
                entity = Activator.CreateInstance<T>();
                PropertyInfo[] properties = entity.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    PropertyInfo item = properties[i];
                    if (properties[i].CustomAttributes != null && properties[i].CustomAttributes.Count<CustomAttributeData>() != 0)
                    {
                        CustomAttributeData ca = item.CustomAttributes.First<CustomAttributeData>();
                        Header header = new Header();
                        header.PropertyInfo = item;
                        foreach (CustomAttributeNamedArgument na in ca.NamedArguments)
                        {
                            if (na.MemberName == "ColumnName")
                            {
                                header.ColumnName = na.TypedValue.Value.ToString();
                            }
                            if (na.MemberName == "RowName")
                            {
                                header.RowName = na.TypedValue.Value.ToString();
                            }
                        }
                        if (!headerList.Any(h => h.ColumnName == header.ColumnName && h.RowName == header.RowName))
                        {
                            headerList.Add(header);
                        }
                    }
                }
                for (int i = 1; i <= data.GetLength(0); i++)
                {
                    for (int j = 1; j <= data.GetLength(1); j++)
                    {
                        if (data[i, j] != null)
                        {
                            var header = headerList.Where(h => data[i, j].ToString().Contains(h.ColumnName)).ToList();
                            if (header != null && header.Count > 0)
                            {
                                header.ForEach(h => h.ColumnIndex = j);
                            }

                            var headerUpdate = headerList.Where(h => data[i, j].ToString().Contains(h.RowName) && h.ColumnIndex != 0).ToList();
                            if (headerUpdate != null && headerUpdate.Count > 0)
                            {
                                headerUpdate.ForEach(h =>
                                {
                                    h.PropertyInfo.SetValue(entity, Convert.ToDouble(data[i, h.ColumnIndex]?.ToString().Trim()));
                                });
                                break;
                            }
                        }
                    }
                }
                return entity;
            }

            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                bool flag8 = excelHelper != null;
                if (flag8)
                {
                    excelHelper.Close();
                }
            }
        }

        public void WriteExcel(string fileName, Model model)
        {
            ExcelHelper excelHelper = null;
            try
            {
                excelHelper = new ExcelHelper(fileName);
                this.CountArray(model);
                this.CountPnr();
                double hv = this.HValue();
                double mv = this.MValue();
                double lv = this.LValue();
                double snrv = this.SNRValue();
                double nrrv = this.NRRValue();
                double slcv = this.SLCValue();
                double cv = this.ClassValue();

                int rowIndex = 1;
                int colIndex = 1;
                excelHelper.SetCellValue(rowIndex, colIndex, model.Company);
                rowIndex = 3;
                excelHelper.SetCellValue(rowIndex, colIndex, model.TestMethod);
                excelHelper.SetCellValue(rowIndex, 7, model.Position);
                rowIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, model.Manufacturer);
                excelHelper.SetCellValue(rowIndex, 7, model.StrTestDate);
                rowIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, model.ModelValue);
                excelHelper.SetCellValue(rowIndex, 7, model.TestedBy);

                rowIndex = 10;
                colIndex = 2;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[1]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[2]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[3]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[4]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[5]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[6]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[7]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[8]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[9]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.mean[10]);

                rowIndex++;
                colIndex = 2;

                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[1]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[2]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[3]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[4]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[5]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[6]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[7]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[8]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[9]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.dev[10]);

                rowIndex++;
                colIndex = 2;

                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[1]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[2]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[3]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[4]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[5]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[6]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[7]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[8]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[9]);
                colIndex++;
                excelHelper.SetCellValue(rowIndex, colIndex, this.apv[10]);

                excelHelper.SetRangeFormat("B10", "K12", "0.0");

                excelHelper.SetCellValue(14, 2, hv);
                excelHelper.SetRangeFormat("B14", "B14", "##");
                excelHelper.SetCellValue(15, 2, mv);
                excelHelper.SetRangeFormat("B15", "B15", "##");
                excelHelper.SetCellValue(16, 2, lv);
                excelHelper.SetRangeFormat("B16", "B16", "##");
                excelHelper.SetCellValue(17, 2, snrv);
                excelHelper.SetRangeFormat("B17", "B17", "##.0");
                excelHelper.SetCellValue(18, 2, nrrv);
                excelHelper.SetRangeFormat("B18", "B18", "##.0");
                excelHelper.SetCellValue(19, 2, slcv);
                excelHelper.SetRangeFormat("B19", "B19", "##.0");
                excelHelper.SetCellValue(20, 2, cv);
                excelHelper.SetRangeFormat("B20", "B20", "##");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                excelHelper.Save();
                bool flag = excelHelper != null;
                if (flag)
                {
                    excelHelper.Close();
                }
            }
        }

        private void CountArray(Model model)
        {
            bool flag = model.Mean_63 == 0.0;
            if (flag)
            {
                model.Mean_63 = model.Mean_125;
                model.St_63 = model.St_125;
            }
            this.mean[1] = model.Mean_63;
            this.mean[2] = model.Mean_125;
            this.mean[3] = model.Mean_250;
            this.mean[4] = model.Mean_500;
            this.mean[5] = model.Mean_1000;
            this.mean[6] = model.Mean_2000;
            this.mean[7] = model.Mean_3150;
            this.mean[8] = model.Mean_4000;
            this.mean[9] = model.Mean_6300;
            this.mean[10] = model.Mean_8000;
            this.dev[1] = model.St_63;
            this.dev[2] = model.St_125;
            this.dev[3] = model.St_250;
            this.dev[4] = model.St_500;
            this.dev[5] = model.St_1000;
            this.dev[6] = model.St_2000;
            this.dev[7] = model.St_3150;
            this.dev[8] = model.St_4000;
            this.dev[9] = model.St_6300;
            this.dev[10] = model.St_8000;
            this.apv[1] = model.Apv_63;
            this.apv[2] = model.Apv_125;
            this.apv[3] = model.Apv_250;
            this.apv[4] = model.Apv_500;
            this.apv[5] = model.Apv_1000;
            this.apv[6] = model.Apv_2000;
            this.apv[7] = model.Apv_3150;
            this.apv[8] = model.Apv_4000;
            this.apv[9] = model.Apv_6300;
            this.apv[10] = model.Apv_8000;
            this.C[1] = -1.2;
            this.C[2] = -0.49;
            this.C[3] = 0.14;
            this.C[4] = 1.56;
            this.C[5] = -2.98;
            this.C[6] = -1.01;
            this.C[7] = 0.85;
            this.C[8] = 3.14;
        }

        private double Rounding(double value, int number)
        {
            return (double)Math.Round((decimal)value, number, MidpointRounding.AwayFromZero);
        }

        private double LogMethon(double value)
        {
            return Math.Log(value) / Math.Log(10.0);
        }

        private double HValue()
        {
            return 0.25 * (this.pnr[1] + this.pnr[2] + this.pnr[3] + this.pnr[4]) - 0.48 * (this.C[1] * this.pnr[1] + this.C[2] * this.pnr[2] + this.C[3] * this.pnr[3] + this.C[4] * this.pnr[4]);
        }

        private double MValue()
        {
            double mValue = 0.0;
            for (int i = 5; i <= 8; i++)
            {
                mValue += 0.25 * this.pnr[i] - 0.16 * this.C[i] * this.pnr[i];
            }
            return mValue;
        }

        private double LValue()
        {
            double lValue = 0.0;
            for (int i = 5; i <= 8; i++)
            {
                lValue += 0.25 * this.pnr[i] + 0.23 * this.C[i] * this.pnr[i];
            }
            return lValue;
        }

        private double SNRValue()
        {
            double snrSum = Math.Pow(10.0, 0.1 * (75.4 - this.apv[2])) + Math.Pow(10.0, 0.1 * (82.9 - this.apv[3])) + Math.Pow(10.0, 0.1 * (88.3 - this.apv[4])) + Math.Pow(10.0, 0.1 * (91.5 - this.apv[5])) + Math.Pow(10.0, 0.1 * (92.7 - this.apv[6])) + Math.Pow(10.0, 0.1 * (92.5 - this.apv[8])) + Math.Pow(10.0, 0.1 * (90.4 - this.apv[10]));
            return 100.0 - 10.0 * this.LogMethon(snrSum);
        }

        private double NRRValue()
        {
            double[] tem = new double[]
            {
        0.0,
        this.mean[1] - 2.0 * this.dev[1],
        this.mean[2] - 2.0 * this.dev[2],
        this.mean[3] - 2.0 * this.dev[3],
        this.mean[4] - 2.0 * this.dev[4],
        this.mean[5] - 2.0 * this.dev[5],
        this.mean[6] - 2.0 * this.dev[6],
        this.mean[7] - 2.0 * this.dev[7],
        this.mean[8] - 2.0 * this.dev[8],
        this.mean[9] - 2.0 * this.dev[9],
        this.mean[10] - 2.0 * this.dev[10]
            };
            double apu = (this.mean[7] + this.mean[8]) / 2.0 - (this.dev[7] + this.dev[8]);
            double apu2 = (this.mean[9] + this.mean[10]) / 2.0 - (this.dev[9] + this.dev[10]);
            double nrrSum = Math.Pow(10.0, 0.1 * (83.9 - tem[2])) + Math.Pow(10.0, 0.1 * (91.4 - tem[3])) + Math.Pow(10.0, 0.1 * (96.8 - tem[4])) + Math.Pow(10.0, 0.1 * (100.0 - tem[5])) + Math.Pow(10.0, 0.1 * (101.2 - tem[6])) + Math.Pow(10.0, 0.1 * (101.0 - apu)) + Math.Pow(10.0, 0.1 * (98.9 - apu2));
            return 107.9 - 10.0 * this.LogMethon(nrrSum) - 3.1;
        }

        private void CountPnr()
        {
            double[,] La = new double[9, 9];
            La[1, 1] = 51.4;
            La[1, 2] = 62.6;
            La[1, 3] = 70.8;
            La[1, 4] = 81.0;
            La[1, 5] = 90.4;
            La[1, 6] = 96.2;
            La[1, 7] = 94.7;
            La[1, 8] = 92.3;
            La[2, 1] = 59.5;
            La[2, 2] = 68.9;
            La[2, 3] = 78.3;
            La[2, 4] = 84.3;
            La[2, 5] = 92.8;
            La[2, 6] = 96.3;
            La[2, 7] = 94.0;
            La[2, 8] = 90.0;
            La[3, 1] = 59.8;
            La[3, 2] = 71.1;
            La[3, 3] = 80.8;
            La[3, 4] = 88.0;
            La[3, 5] = 95.0;
            La[3, 6] = 94.4;
            La[3, 7] = 94.1;
            La[3, 8] = 89.0;
            La[4, 1] = 65.4;
            La[4, 2] = 77.2;
            La[4, 3] = 84.5;
            La[4, 4] = 89.8;
            La[4, 5] = 95.5;
            La[4, 6] = 94.3;
            La[4, 7] = 92.5;
            La[4, 8] = 88.8;
            La[5, 1] = 65.3;
            La[5, 2] = 77.4;
            La[5, 3] = 86.5;
            La[5, 4] = 92.5;
            La[5, 5] = 96.4;
            La[5, 6] = 93.0;
            La[5, 7] = 90.4;
            La[5, 8] = 83.7;
            La[6, 1] = 70.7;
            La[6, 2] = 82.0;
            La[6, 3] = 89.3;
            La[6, 4] = 93.3;
            La[6, 5] = 95.6;
            La[6, 6] = 93.0;
            La[6, 7] = 90.1;
            La[6, 8] = 83.0;
            La[7, 1] = 75.6;
            La[7, 2] = 84.2;
            La[7, 3] = 90.1;
            La[7, 4] = 93.6;
            La[7, 5] = 96.2;
            La[7, 6] = 91.3;
            La[7, 7] = 87.9;
            La[7, 8] = 81.9;
            La[8, 1] = 77.6;
            La[8, 2] = 88.0;
            La[8, 3] = 93.4;
            La[8, 4] = 93.8;
            La[8, 5] = 94.2;
            La[8, 6] = 91.4;
            La[8, 7] = 87.9;
            La[8, 8] = 79.9;
            int[] Conver = new int[] { 0, 1, 2, 3, 4, 5, 6, 8, 10 };
            for (int i = 1; i <= 8; i++)
            {
                double pnrSum = 0.0;
                for (int j = 1; j <= 8; j++)
                {
                    pnrSum += Math.Pow(10.0, 0.1 * (La[i, j] - this.apv[Conver[j]]));
                }
                this.pnr[i] = 100.0 - 10.0 * this.LogMethon(pnrSum);
            }
        }

        private double SLCValue()
        {
            double slc = 71.0 - this.apv[2];
            double slc2 = 81.0 - this.apv[3];
            double slc3 = 89.0 - this.apv[4];
            double slc4 = 93.0 - this.apv[5];
            double slc5 = 95.0 - this.apv[6];
            double slc6 = 93.0 - this.apv[8];
            double slc7 = 86.0 - this.apv[10];
            double slcSum = Math.Pow(10.0, slc / 10.0) + Math.Pow(10.0, slc2 / 10.0) + Math.Pow(10.0, slc3 / 10.0) + Math.Pow(10.0, slc4 / 10.0) + Math.Pow(10.0, slc5 / 10.0) + Math.Pow(10.0, slc6 / 10.0) + Math.Pow(10.0, slc7 / 10.0);
            return 100.0 - 10.0 * this.LogMethon(slcSum);
        }

        private double ClassValue()
        {
            double value = this.Rounding(this.SLCValue(), 0);
            double classValue = 0.0;
            bool flag = value >= 10.0 && value <= 13.0;
            if (flag)
            {
                classValue = 1.0;
            }
            else
            {
                bool flag2 = value >= 14.0 && value <= 17.0;
                if (flag2)
                {
                    classValue = 2.0;
                }
                else
                {
                    bool flag3 = value >= 18.0 && value <= 21.0;
                    if (flag3)
                    {
                        classValue = 3.0;
                    }
                    else
                    {
                        bool flag4 = value >= 22.0 && value <= 24.0;
                        if (flag4)
                        {
                            classValue = 4.0;
                        }
                        else
                        {
                            bool flag5 = value >= 25.0;
                            if (flag5)
                            {
                                classValue = 5.0;
                            }
                        }
                    }
                }
            }
            return classValue;
        }
    }

    internal class Header
    {
        public string ColumnName
        {
            get;
            set;
        }

        public int ColumnIndex
        {
            get;
            set;
        }

        public string RowName
        {
            get;
            set;
        }

        public PropertyInfo PropertyInfo
        {
            get;
            set;
        }
    }
}
