using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport
{
    public class Model
    {
        [Mapping(ColumnName = "63Hz", RowName = "MEANS")]
        public double Mean_63
        {
            get;
            set;
        }

        [Mapping(ColumnName = "125Hz", RowName = "MEANS")]
        public double Mean_125
        {
            get;
            set;
        }

        [Mapping(ColumnName = "250Hz", RowName = "MEANS")]
        public double Mean_250
        {
            get;
            set;
        }

        [Mapping(ColumnName = "500Hz", RowName = "MEANS")]
        public double Mean_500
        {
            get;
            set;
        }

        [Mapping(ColumnName = "1000Hz", RowName = "MEANS")]
        public double Mean_1000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "2000Hz", RowName = "MEANS")]
        public double Mean_2000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "3150Hz", RowName = "MEANS")]
        public double Mean_3150
        {
            get;
            set;
        }

        [Mapping(ColumnName = "4000Hz", RowName = "MEANS")]
        public double Mean_4000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "6300Hz", RowName = "MEANS")]
        public double Mean_6300
        {
            get;
            set;
        }

        [Mapping(ColumnName = "8000Hz", RowName = "MEANS")]
        public double Mean_8000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "63Hz", RowName = "Standard deviation in dB")]
        public double St_63
        {
            get;
            set;
        }

        [Mapping(ColumnName = "125Hz", RowName = "Standard deviation in dB")]
        public double St_125
        {
            get;
            set;
        }

        [Mapping(ColumnName = "250Hz", RowName = "Standard deviation in dB")]
        public double St_250
        {
            get;
            set;
        }

        [Mapping(ColumnName = "500Hz", RowName = "Standard deviation in dB")]
        public double St_500
        {
            get;
            set;
        }

        [Mapping(ColumnName = "1000Hz", RowName = "Standard deviation in dB")]
        public double St_1000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "2000Hz", RowName = "Standard deviation in dB")]
        public double St_2000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "3150Hz", RowName = "Standard deviation in dB")]
        public double St_3150
        {
            get;
            set;
        }

        [Mapping(ColumnName = "4000Hz", RowName = "Standard deviation in dB")]
        public double St_4000
        {
            get;
            set;
        }

        [Mapping(ColumnName = "6300Hz", RowName = "Standard deviation in dB")]
        public double St_6300
        {
            get;
            set;
        }

        [Mapping(ColumnName = "8000Hz", RowName = "Standard deviation in dB")]
        public double St_8000
        {
            get;
            set;
        }

        public double Apv_63
        {
            get
            {
                bool flag = this.Mean_63 == 0.0;
                double result;
                if (flag)
                {
                    result = this.Mean_125 - this.St_125;
                }
                else
                {
                    result = this.Mean_63 - this.St_63;
                }
                return result;
            }
        }

        public double Apv_125
        {
            get
            {
                return this.Mean_125 - this.St_125;
            }
        }

        public double Apv_250
        {
            get
            {
                return this.Mean_250 - this.St_250;
            }
        }

        public double Apv_500
        {
            get
            {
                return this.Mean_500 - this.St_500;
            }
        }

        public double Apv_1000
        {
            get
            {
                return this.Mean_1000 - this.St_1000;
            }
        }

        public double Apv_2000
        {
            get
            {
                return this.Mean_2000 - this.St_2000;
            }
        }

        public double Apv_3150
        {
            get
            {
                return this.Mean_3150 - this.St_3150;
            }
        }

        public double Apv_4000
        {
            get
            {
                return this.Mean_4000 - this.St_4000;
            }
        }

        public double Apv_6300
        {
            get
            {
                return this.Mean_6300 - this.St_6300;
            }
        }

        public double Apv_8000
        {
            get
            {
                return this.Mean_8000 - this.St_8000;
            }
        }

        public string Company { get; set; }

        private string _testMethod;

        public string TestMethod
        {
            get
            {
                if (string.IsNullOrEmpty(_testMethod))
                    return null;
                return "Test Method:" + _testMethod;
            }
            set
            {
                _testMethod = value;
            }
        }


        private string _position;
        public string Position
        {
            get
            {
                if (string.IsNullOrEmpty(_position))
                    return null;
                return "Position:" + _position;
            }

            set
            {
                _position = value;
            }
        }

        private string _manufacturer;

        public string Manufacturer
        {
            get
            {
                if (string.IsNullOrEmpty(_manufacturer))
                    return null;
                return "Manufacturer:" + _manufacturer;
            }

            set
            {
                _manufacturer = value;
            }
        }

        private string _modelValue;

        public string ModelValue
        {
            get
            {
                if (string.IsNullOrEmpty(_modelValue))
                    return null;
                return "Model:" + _modelValue;
            }

            set
            {
                _modelValue = value;
            }
        }

        private string _testedBy;

        public string TestedBy
        {
            get
            {
                if (string.IsNullOrEmpty(_testedBy))
                    return null;
                return "Tested By:" + _testedBy;
            }

            set
            {
                _testedBy = value;
            }
        }

        public DateTime TestDate { get; set; }

        public string StrTestDate
        {
            get
            {
                return "Date:" + TestDate.ToString("yyyy/MM/dd");
            }
        }

        public bool CanRun
        {
            get
            {
                if (!string.IsNullOrEmpty(Company) && !string.IsNullOrEmpty(TestMethod) && !string.IsNullOrEmpty(Position) && !string.IsNullOrEmpty(Manufacturer) && !string.IsNullOrEmpty(ModelValue) && !string.IsNullOrEmpty(TestedBy))
                    return true;
                else
                    return false;
            }
        }
    }

    [AttributeUsage(AttributeTargets.Property)]
    internal class MappingAttribute : Attribute
    {
        public string ColumnName
        {
            get;
            set;
        }

        public string RowName
        {
            get;
            set;
        }
    }
}
