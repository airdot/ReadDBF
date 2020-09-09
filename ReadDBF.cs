using DbfDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ReadDBF
{
    public partial class ReadDBF : Form
    {
        DataTable dataTable = new DataTable();
        public ReadDBF()
        {
            InitializeComponent();
        }

        private DataTable GetDataTable(string dbfFilename)
        {
             
            DataTable dt = new DataTable();
            using (var dbfTable = new DbfTable(dbfFilename, Encoding.GetEncoding(936)))
            {
                var header = dbfTable.Header;
                var dbfRecord = new DbfRecord(dbfTable);
                var SkipDeleted = true;
                var versionDescription = header.VersionDescription;
                var hasMemo = dbfTable.Memo != null;
                var recordCount = header.RecordCount;
                List<DbfHeader> dbfHeaders = new List<DbfHeader>();
                foreach (var dbfColumn in dbfTable.Columns)
                {
                    var name = dbfColumn.Name;
                    var length = dbfColumn.Length;
                    DbfHeader dbfHeader = new DbfHeader();
                    dbfHeader.Name = name;
                    dbfHeader.Length = length;
                    dbfHeaders.Add(dbfHeader);
                }
                foreach (var dbfHeader in dbfHeaders)
                {
                    var name = dbfHeader.Name;
                    dt.Columns.Add(name);
                }
                int i = 0;
                while (dbfTable.Read(dbfRecord))
                {
                    int j = 0;
                    if (SkipDeleted && dbfRecord.IsDeleted) continue;
                    dt.Rows.Add();
                    foreach (var dbfValue in dbfRecord.Values)
                    {
                        var stringValue = dbfValue.ToString();
                        dt.Rows[i][j] = stringValue;
                        j++;
                    }
                    i++;
                }
            }
            
            return dt;
        }

        private DbfTable GetDbfTable(string dbfPath)
        {
            using (var dbfTable = new DbfTable(dbfPath, Encoding.GetEncoding(936)))
            {
                return dbfTable;
            }
        }

        private string GetDbfFile()
        {
            var dbfFileName = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = ".\\";
            openFileDialog.Filter = "DBF文件|*.DBF";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                dbfFileName = openFileDialog.FileName;
            }
            return dbfFileName;
        }

        private string SaveExcelFile()
        {
            var excelFileName = "";
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = ".";
            saveFileDialog.Filter = "Excel文件|*.xlsx|Excel2003文件|*.xls";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FilterIndex = 1;
            if(saveFileDialog.ShowDialog()== DialogResult.OK)
            {
                excelFileName = saveFileDialog.FileName;
            }
            return excelFileName;
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            //DataTable dataTable = new DataTable();
            var dbfFileName = GetDbfFile();
            if (string.IsNullOrEmpty(dbfFileName)) return;
            var dbfTable = GetDbfTable(dbfFileName);
            DataGridView dgv = dataGridView1;
            dataTable = GetDataTable(dbfFileName);
            dgv.DataSource = dataTable;
            ComboBox cmb = CmboxSearch;
            foreach(var col in dataTable.Columns)
            {                
                cmb.Items.Add(col.ToString());
            }
            
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (CmboxSearch.SelectedIndex == -1) return;
            string searchTxt = $"{CmboxSearch.SelectedItem} like '%{TxtboxSearch.Text}%'";
            Console.WriteLine(searchTxt);
            DataTable dt2 = dataTable.Clone();
            var rows = dataTable.Select(searchTxt);
            if (rows != null && rows.Length > 0)
            {
                foreach (DataRow row in rows)
                {
                    dt2.ImportRow(row);
                }
            }
            dataGridView1.DataSource = dt2;
        }
        /// <summary>
        /// 将DataTable(DataSet)导出到Execl文档
        /// </summary>
        /// <param name="dataTable">传入一个DataTable</param>
        /// <param name="Outpath">导出路径（可以不加扩展名，不加默认为.xls）</param>
        /// <returns>返回一个Bool类型的值，表示是否导出成功</returns>
        /// True表示导出成功，Flase表示导出失败
        public static bool DataTableToExcel(DataTable dataTable, string Outpath)
        {
            bool result = false;
            try
            {
                if (string.IsNullOrEmpty(Outpath))
                    throw new Exception("输入的路径异常");
                int sheetIndex = 0;
                //根据输出路径的扩展名判断workbook的实例类型
                IWorkbook workbook = null;
                string pathExtensionName = Outpath.Trim().Substring(Outpath.Length - 5);
                if (pathExtensionName.Contains(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else if (pathExtensionName.Contains(".xls"))
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    Outpath = Outpath.Trim() + ".xls";
                    workbook = new HSSFWorkbook();
                }
                //将DataSet导出为Excel
                
                    sheetIndex++;
                    if (dataTable != null && dataTable.Rows.Count > 0)
                    {
                        ISheet sheet = workbook.CreateSheet(string.IsNullOrEmpty(dataTable.TableName) ? ("sheet" + sheetIndex) : dataTable.TableName);//创建一个名称为Sheet0的表
                        int rowCount = dataTable.Rows.Count;//行数
                        int columnCount = dataTable.Columns.Count;//列数

                        //设置列头
                        IRow row = sheet.CreateRow(0);//excel第一行设为列头
                        for (int c = 0; c < columnCount; c++)
                        {
                            ICell cell = row.CreateCell(c);
                            cell.SetCellValue(dataTable.Columns[c].ColumnName);
                        }

                        //设置每行每列的单元格,
                        for (int i = 0; i < rowCount; i++)
                        {
                            row = sheet.CreateRow(i + 1);
                            for (int j = 0; j < columnCount; j++)
                            {
                                ICell cell = row.CreateCell(j);//excel第二行开始写入数据
                                cell.SetCellValue(dataTable.Rows[i][j].ToString());
                            }
                        }
                    }
                
                //向outPath输出数据
                using (FileStream fs = File.OpenWrite(Outpath))
                {
                    workbook.Write(fs);//向打开的这个xls文件中写入数据
                    result = true;
                }
                return result;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var excelFileName = SaveExcelFile();
            bool saveFile = DataTableToExcel(dataTable, excelFileName);
            if (saveFile == false) MessageBox.Show("导出失败！");
            else MessageBox.Show("导出成功！");
        }
    }

}
