using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace ExcelLinkage
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range target = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Worksheet sheet = (Excel.Worksheet)sh;

            Excel.Worksheet dataSheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[1]);



            Excel.Shape shArrow;
            long lngXStart, lngYStart;
            long lngXEnd, lngYEnd;

            Dictionary<string, int> dataToRow = new Dictionary<string, int>();
            Color[] colors = { Color.White ,Color.Aqua , Color.DarkOliveGreen, Color.Red, Color.Green, Color.DarkTurquoise, Color.Beige, Color.Crimson, Color.Tan, Color.MistyRose };
            Excel.Range myRange = dataSheet.get_Range("A:A", Type.Missing);
            foreach (Excel.Range c in myRange) {
                if (c.Value2 == null || c.Value2.ToString()=="")
                    break;
                if (c.Value2.ToString() != "ID")
                    dataToRow.Add(c.Value2.ToString(), c.Row);
                //MessageBox.Show(c.Value2.ToString());
            }

            int total_thread = 0;
            Dictionary<int, int> endNode = new Dictionary<int, int>();
            int parent_index = 0;

            dataSheet.Columns[12].ColumnWidth = 20;

            foreach (Excel.Range c in target.Cells) {
                    //MessageBox.Show(c.Value.ToString());
                    String[] str = c.Value.ToString().Split(',');

                    int row_parent = dataToRow[str[0]];
                    int row_child = dataToRow[str[1]];
                if (endNode.ContainsKey(row_parent))
                {
                    parent_index = endNode[row_parent];
                    endNode.Remove(row_parent);
                    endNode.Add(row_child, parent_index);
                }
                else {
                    total_thread = total_thread + 1;
                    endNode.Add(row_child, total_thread);
                    parent_index = total_thread;
                }
                    // int row1 = Convert.ToInt32(str[0]);
                    // int row2 = Convert.ToInt32(str[1]);
                    //int parent_index = Convert.ToInt32(str[2]);

                    Excel.Range rngFrom = (Excel.Range)dataSheet.Cells[row_parent, 12];
                    Excel.Range rngTo = (Excel.Range)dataSheet.Cells[row_child, 12];

                  //lngXStart = (long)(rngFrom.Left + rngFrom.Width / 2);
                 // lngXEnd = (long)(rngTo.Left + rngTo.Width / 2);
                    lngYStart = (long)(rngFrom.Top + rngFrom.Height / 2);
                    lngYEnd = (long)(rngTo.Top + rngTo.Height / 2);

                    lngXStart = (long)(rngFrom.Left + (rngFrom.Width / 10)* parent_index);
                    lngXEnd = (long)(rngTo.Left + (rngTo.Width / 10)* parent_index);

                    shArrow = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[1].Shapes.addconnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorCurve, lngXStart, lngYStart, lngXEnd, lngYEnd);
                    //shArrow = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeCurvedDownArrow, lngXStart, lngYStart, 5 , 10);
                    shArrow.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle;
                    shArrow.Line.EndArrowheadWidth = Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWide;
                    shArrow.Line.Weight = 1;
                    shArrow.Line.BeginArrowheadStyle= Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOval;
                    shArrow.Line.BeginArrowheadWidth = Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWide;
                    shArrow.Line.ForeColor.RGB = colors[parent_index].ToArgb();
                    

                //shArrow.Line.ForeColor = Color.Red;

            }



            //var nonEmptyRanges = myRange.Cast<Excel.Range>().Where(r => !string.IsNullOrEmpty(r.Text));
            //nonEmptyRanges=nonEmptyRanges.Skip(1);

            /*foreach (Excel.Range c in target.Cells) {
                //MessageBox.Show(c.Value.ToString());
                String[] str = c.Value.ToString().Split(',');
                int row1 = Convert.ToInt32(str[0]);
                int row2 = Convert.ToInt32(str[1]);
                int parent_index = Convert.ToInt32(str[2]);

                Excel.Range rngFrom = (Excel.Range)sheet.Cells[row1, 4];
                Excel.Range rngTo = (Excel.Range)sheet.Cells[row2, 4];

              //lngXStart = (long)(rngFrom.Left + rngFrom.Width / 2);
             // lngXEnd = (long)(rngTo.Left + rngTo.Width / 2);
                lngYStart = (long)(rngFrom.Top + rngFrom.Height / 2);
                lngYEnd = (long)(rngTo.Top + rngTo.Height / 2);

                lngXStart = (long)(rngFrom.Left + (rngFrom.Width / 10)* parent_index);
                lngXEnd = (long)(rngTo.Left + (rngTo.Width / 10)* parent_index);

                shArrow = Globals.ThisAddIn.Application.ActiveSheet.Shapes.addconnector(Microsoft.Office.Core.MsoConnectorType.msoConnectorCurve, lngXStart, lngYStart, lngXEnd, lngYEnd);
                //shArrow = Globals.ThisAddIn.Application.ActiveSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeCurvedDownArrow, lngXStart, lngYStart, 5 , 10);
                shArrow.Line.EndArrowheadStyle = Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadTriangle;


                //shArrow.Line.ForeColor = Color.Red;

            }*/
        }
    }
}
