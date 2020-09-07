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


    }
}
