using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InIFileLib
{
    public interface IIniFileController
    {
        void IniWriteValue(string section, string key, string iValue);
        string IniReadValue(string section, string key);
        List<string> ReadKeys(string section);
        DataSet ExcelToToDataSet(string path);
        void WriteCSV(string fileName, DataTable dt);
    }
}
