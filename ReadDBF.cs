using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DbfDataReader;
using NPOI;

namespace ReadDBF
{
    public partial class ReadDBF : Form
    {
        public ReadDBF()
        {
            InitializeComponent();
        }

        private DataSet GetDataset(DbfTable dbfTable, List<DbfHeader> dbfHeaders)
        {
            DataSet da = new DataSet();
            DataTable dt = new DataTable();
            var dbfRecord = new DbfRecord(dbfTable);
            var SkipDeleted = true;
            foreach(var dbfHeader in dbfHeaders)
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
            da.Tables.Add(dt);
            return da;
        }

        private List<DbfHeader> GetDbfHeaders(DbfTable dbfTable)
        {
            List<DbfHeader> dbfHeaders = new List<DbfHeader>();
            DbfHeader dbfHeader = new DbfHeader();

            var header = dbfTable.Header;
            foreach (var dbfColumn in dbfTable.Columns)
            {
                var name = dbfColumn.Name;
                var length = dbfColumn.Length;

                dbfHeader.Name = name;
                dbfHeader.Length = length;

                dbfHeaders.Add(dbfHeader);
            }
            return dbfHeaders;
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
            var fileName = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
            openFileDialog.Filter = "DBF文件|*.DBF";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = openFileDialog.FileName;
            }
            return fileName;
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            var dbfFileName = GetDbfFile();
            if (string.IsNullOrEmpty(dbfFileName)) return;
            var dbfTable = GetDbfTable(dbfFileName);
            var dbfHeaders = GetDbfHeaders(dbfTable);
            DataGridView dgv = dataGridView1;
            foreach(var dbfHeader in dbfHeaders)
            {
                var name = dbfHeader.Name;
                var length = dbfHeader.Length;
                dgv.Columns.Add(name, name);
                dgv.Columns[name].Width = length;
            }
            dgv.DataSource = GetDataset(dbfTable,dbfHeaders);
        }
    }
}
