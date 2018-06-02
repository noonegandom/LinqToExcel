using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LinqToExcel;

namespace LinqToExcel
{
    public partial class Form1 : Form
    {
        public ExcelQueryFactory _Excel { get; set; }
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                _Excel = new ExcelQueryFactory(filePath);
                DoSomething();
            }
        }
        private IEnumerable<string> GetSheets()
        {
           return _Excel.GetWorksheetNames();
        }
        private List<SheetsModel> GetColumnNames(IEnumerable<string> sheets)
        {
            var model = new List<SheetsModel>();
            foreach (var sheet in sheets)
            {
               model.Add(
                   new SheetsModel {
                       SheetName = sheet,
                       ColumnNames = (_Excel.GetColumnNames(sheet)).ToList()
                   });
            }
            return model;
        }
        public void DoSomething()
        {
            var sheets = GetSheets();
            var SheetsWithColumns = GetColumnNames(sheets);
            ShowNodeTree(SheetsWithColumns);
        }
        public void ShowNodeTree(List<SheetsModel> SheetsWithColumns)
        {
            foreach (var sheet in SheetsWithColumns)
            {
                treeView1.BeginUpdate();
                var parentNode = treeView1.Nodes.Add(sheet.SheetName);
                treeView1.EndUpdate();
                treeView1.BeginUpdate();
                foreach (var column in sheet.ColumnNames)
                {
                    parentNode.Nodes.Add(column);
                }
                treeView1.EndUpdate();
            }
        }
        public class SheetsModel
        {
            public string SheetName { get; set; }
            public List<string> ColumnNames { get; set; }
        }
    }
}
