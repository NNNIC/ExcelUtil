using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test7
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string m_sample_xls_file = Environment.GetEnvironmentVariable("USERPROFILE") +  @"\a.xls";

        ExcelUtil.BookCtr  m_bc;
        ExcelUtil.SheetCtr m_sheet;

        private void button1_Click(object sender, EventArgs e)
        {
            if (m_bc!=null)
            {
                m_bc.Dispose();
                m_bc = null;
            }

            m_bc = ExcelUtil.OpenBook(m_sample_xls_file);
            m_bc.SetVisible(true);
            m_sheet = m_bc.GetActiveSheet();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var cols = 10;
            var rows = 2;
            var objs = m_sheet.GetValues(rows,cols);
            for(var x = 1; x<=rows; x++) for(var y = 1; y<=cols; y++)
            {
                objs[x,y] = x.ToString() + "|" + y.ToString();
            }
           m_sheet.SetValues(objs);

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (m_bc!=null)
            {
                m_bc.Dispose();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var objs = m_sheet.GetValues();
            var str = string.Empty;
            for(var x = 1; x<=objs.GetLength(0) ; x++) for(var y = 1; y<=objs.GetLength(1); y++)
            {
                if (str!=string.Empty) str += ",";
                var o = objs[x,y];
                if (o!=null)
                {
                    str += o.ToString();
                }
            }
            MessageBox.Show(str);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            m_bc.Write();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            m_bc = ExcelUtil.AttackBook();
            if (m_bc!=null)
            {
                m_sheet = m_bc.GetActiveSheet();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var objs = m_sheet.GetValues(1000,200);
            for(var x = 1; x<=objs.GetLength(0) ; x++) for(var y = 1; y<=objs.GetLength(1); y++)
            {
                objs[x,y] = x.ToString("0000") + "," + y.ToString("0000");
            }
            m_sheet.SetValues(objs);
        }
    }
}
