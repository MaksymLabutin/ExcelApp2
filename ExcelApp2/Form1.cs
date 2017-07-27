using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
namespace ExcelApp2
{
    public partial class Form1 : Form
    {




        public Form1()
        {
            InitializeComponent();
        }

        private void openExcelFile_Click(object sender, EventArgs e)
        {

            buildDiagram();

        }


        public void buildDiagram()
        {
            
            System.Array date, X;
            DateTime dateChart;
            int x;
            OpenExcelFile.getExcelFile(out  X, out  date);
            for(int i = 0; i <= X.Length; i++)
            {
                //dateChart = date[i,1];
                if (i == 0)
                    continue;
                this.diagram.Series["series1"].Points.AddXY(i, i);

            }
        }
    }
}
