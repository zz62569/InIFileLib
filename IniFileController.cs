using ExcelDataReader;
using System.Data;
using System.Diagnostics;
using System.Text;


namespace InIFileLib
{
    public class IniFileController : IIniFileController
    {
        #region 读写系统文件
        [System.Runtime.InteropServices.DllImport("kernel32", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [System.Runtime.InteropServices.DllImport("kernel32", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int GetPrivateProfileString(string section, string key, string defVal, StringBuilder retVal, int size, string filePath);
        [System.Runtime.InteropServices.DllImport("kernel32", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern uint GetPrivateProfileStringA(string section, string key,
                                                                string def, byte[] retVal, int size, string filePath);

        public string Path;
        /// <summary>
        /// INI文件的位置
        /// </summary>
        public IniFileController(string path)
        {
            Path = path;
        }
        /// <summary>
        /// 写INI文件
        /// </summary>
        /// <param name="section">段落</param>
        /// <param name="key">键</param>
        /// <param name="iValue">值</param>
        public void IniWriteValue(string section, string key, string iValue)
        {
            WritePrivateProfileString(section, key, iValue, this.Path);
        }
        /// <summary>
        /// 读取INI文件
        /// </summary>
        /// <param name="section">段落</param>
        /// <param name="key">键</param>
        /// <returns>返回的键值</returns>
        public string IniReadValue(string section, string key)
        {
            StringBuilder temp = new(255);
            _ = GetPrivateProfileString(section, key, "", temp, 255, this.Path);
            return temp.ToString();
        }

        /// <summary>
        /// 读取INI文件
        /// </summary>
        /// <param name="SectionName">段落</param>
        /// <returns>返回所有键值</returns>
        public List<string> ReadKeys(string section)
        {
            List<string> result = new();
            byte[] buf = new byte[65535];
            uint len = GetPrivateProfileStringA(section, string.Empty, string.Empty, buf, buf.Length, Path);
            int j = 0;
            for (int i = 0; i < len; i++)
                if (buf[i] == 0)
                {
                    result.Add(Encoding.Default.GetString(buf, j, i - j));
                    j = i + 1;
                }
            return result;
        }
        #endregion
        /// <summary>
        /// 获取文件后缀名，如.PDF
        /// </summary>
        /// <param name="fileName"是传入的文件名></param>
        /// <returns></returns>
        private string FileSuffixName(string fileName)
        {
            int suff = fileName.Length - 1;
            while (suff >= 0 && (fileName[suff] != '.'))
                suff--;
            string suffixName = fileName.Remove(0, suff);//获取文件后缀名，如.PDF
            return suffixName;
        }

        /// <summary>
        /// 
        /// </summary>
        public DataSet ExcelToToDataSet(string path)
        {
            DataSet result = new();
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                Trace.WriteLine(path);
                var _ = FileSuffixName(path);
                IExcelDataReader reader;
                if (_ == ".xls")
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                else if (_ == ".xlsx")
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                else if (_ == ".CSV")
                    reader = ExcelReaderFactory.CreateCsvReader(stream);
                else
                    return new DataSet();
                result = reader.AsDataSet();
                reader.Dispose();
            }
            return result;
        }
        /// <summary>
        /// 写入CSV
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="dt">要写入的datatable</param>
        public void WriteCSV(string fileName, DataTable dt)
        {
            FileStream fs;
            StreamWriter sw;
            string? data = null;
            //判断文件是否存在,存在就不再次写入列名
            if (!File.Exists(fileName))
            {
                fs = new (fileName, FileMode.Create, FileAccess.Write);
                sw = new (fs, Encoding.UTF8);
                //写出列名称
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data += dt.Columns[i].ColumnName.ToString();
                    if (i < dt.Columns.Count - 1)
                    {
                        data += ",";//中间用，隔开
                    }
                }
                sw.WriteLine(data);
            }
            else
            {
                fs = new (fileName, FileMode.Append, FileAccess.Write);
                sw = new (fs, Encoding.UTF8);
            }
            //写出各行数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                data = null;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    data += dt.Rows[i][j].ToString();
                    if (j < dt.Columns.Count - 1)
                    {
                        data += ",";//中间用，隔开
                    }
                }
                sw.WriteLine(data);
            }
            sw.Close();
            fs.Close();
        }
    }
}
