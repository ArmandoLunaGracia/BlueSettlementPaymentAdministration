using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;       // EXCEL APPLICATION.


namespace MexInvoiceDownPayment
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dataGridView1.RowStateChanged += new System.Windows.Forms.DataGridViewRowStateChangedEventHandler(dataGridView1_RowStateChanged);
        }

        // CREATE EXCEL OBJECTS.
        Excel.Application xlApp = new Excel.Application();
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;

        string sFileName= System.IO.Path.Combine(@"C:\PDFExtractor\MexInvoice", "Settlements.xlsx"); 

        
        

        // IMPORT DATA FROM EXCEL AND POPULATE THE GRID.
        private void Excel2Grid(string sFile)
        {

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(sFile);               // WORKBOOK TO OPEN THE EXCEL FILE.
            xlWorkSheet = xlWorkBook.Worksheets["Details"];          // THE SHEET WITH THE DATA.

            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            int iRow, iCol;

            // FIRST, CREATE THE DataGridView COLUMN HEADERS.
            for (iCol = 1; iCol <= xlWorkSheet.Columns.Count; iCol++)
            {
                if (xlWorkSheet.Cells[1, iCol].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                    col.HeaderText = xlWorkSheet.Cells[1, iCol].value;
                    int colIndex = dataGridView1.Columns.Add(col);        // ADD A NEW COLUMN.
                }
            }

            //ADD A ROWINDEX COLUMN FOR EXCEL SAVING

            DataGridViewTextBoxColumn col1 = new DataGridViewTextBoxColumn();
            col1.HeaderText = "XlsRowIndex";
            dataGridView1.Columns.Add(col1);

            // ADD A BUTTON AT THE LAST COLUMN IN EVERY ROW.
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            btn.HeaderText = "";
            btn.Text = "Save Data";
            btn.Name = "btSave";
            btn.UseColumnTextForButtonValue = true;
            dataGridView1.Columns.Add(btn);
            double totalowed = 0;
            // ADD ROWS TO THE GRID USING EXCEL DATA.
            for (iRow = 2; iCol <= xlWorkSheet.Rows.Count; iRow++)
            {
                if (xlWorkSheet.Cells[iRow, 1].value == null)
                {
                    break;      // BREAK LOOP.
                }
                else
                {
                    // CREATE A STRING ARRAY USING THE VALUES IN EACH ROW OF THE SHEET.
                    if(Convert.ToDouble(xlWorkSheet.Cells[iRow, 8].value) > 0) 
                    {
                        string[] row = new string[] { xlWorkSheet.Cells[iRow, 1].value,
                        xlWorkSheet.Cells[iRow, 2].value.ToString(),
                        xlWorkSheet.Cells[iRow, 3].value.ToString(),
                        xlWorkSheet.Cells[iRow, 4].value is null ? "": xlWorkSheet.Cells[iRow, 4].value.ToString(),
                        xlWorkSheet.Cells[iRow, 5].value is null ? "": xlWorkSheet.Cells[iRow, 5].value.ToString(),
                        xlWorkSheet.Cells[iRow, 6].value is null ? "": xlWorkSheet.Cells[iRow, 6].value.ToString(),
                        xlWorkSheet.Cells[iRow, 7].value is null ? "": xlWorkSheet.Cells[iRow, 7].value.ToString(),
                        xlWorkSheet.Cells[iRow, 8].value is null ? "": xlWorkSheet.Cells[iRow, 8].value.ToString(),Convert.ToString(iRow)};
                        // ADD A NEW ROW TO THE GRID USING THE ARRAY DATA.
                        dataGridView1.Rows.Add(row);
                        totalowed = totalowed + Convert.ToDouble(xlWorkSheet.Cells[iRow, 8].value);
                    }
                    


                }
            }

            xlWorkBook.Close();
            xlApp.Quit();

            // CLEAN UP.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            label6.Text = Convert.ToString(totalowed);
        }

        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                    SendKeys.Send("{TAB}");     // MOVE NEXT CELL WHEN YOU PRESS ENTER KEY.
                return true;
            }
            else
            {
                return base.ProcessCmdKey(ref msg, keyData);
            }
        }

       

        // CHANGE THE COLOR OF VALUES IN THE FIRST COLUMN. MAKE THE VALUES REALONLY (CANNOT CHANGE).
        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
            if (e.Row.Cells[0].Value != null)
            {
                e.Row.Cells[0].Style.ForeColor = Color.Gray;
                e.Row.Cells[0].ReadOnly = true;
                e.Row.Cells[1].Style.ForeColor = Color.Gray;
                e.Row.Cells[1].ReadOnly = true;
                e.Row.Cells[2].Style.ForeColor = Color.Gray;
                e.Row.Cells[2].ReadOnly = true;
                e.Row.Cells[3].Style.ForeColor = Color.Gray;
                e.Row.Cells[3].ReadOnly = true;
                e.Row.Cells[7].Style.ForeColor = Color.Gray;
                e.Row.Cells[7].ReadOnly = true;
                e.Row.Cells[5].Style.ForeColor = Color.Gray;
                e.Row.Cells[5].ReadOnly = true;
                e.Row.Cells[6].Style.ForeColor = Color.Gray;
                e.Row.Cells[6].ReadOnly = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Enabled = false;
            Excel2Grid(sFileName);
        }
        public static bool IsNumeric(string s)
        {
            foreach (char c in s)
            {
                if (!char.IsDigit(c) && c != '.' && c != '-')
                {
                    return false;
                }
            }

            return true;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // EVERY ROW HAS A BUTTON AT THE LAST COLUMN. 
            // SAVE THE DATA IN EXCEL AFTER CLICKING THE BUTTON.

            var ourGrid = (DataGridView)sender;
            if (ourGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
            {
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(sFileName);   // WORKBOOK TO OPEN THE EXCEL FILE.
                xlWorkSheet = xlWorkBook.Worksheets["Details"];  // THE SHEET WITH THE DATA.

                // CHECK IF THE FIRST COLUMN IS ReadOnly. 
                // THIS IS TO ENSURE THAT YOU MODIFY EXISTING DATA IN EXCEL.

                if (dataGridView1.Rows[e.RowIndex].Cells[0].ReadOnly == true)
                {
                    string sXL = xlWorkSheet.Cells[e.RowIndex + 2, 1].value;
                    string sGrid = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();

                    if (sXL == sGrid)
                    {
                        if (IsNumeric(Convert.ToString(dataGridView1.Rows[e.RowIndex].Cells[4].Value)) ==false) 
                        {
                            MessageBox.Show("Please input numeric value for Downpayment amount");
                            
                        }
                        else
                        {
                            if (Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[4].Value) <= Convert.ToDouble(textBox2.Text)) 
                            {
                                double sumdown = 0;
                                double totalowed = 0;
                                foreach (DataGridViewRow dr in dataGridView1.Rows)
                                {
                                    if (Convert.ToString(dr.Cells[4].Value) == string.Empty) 
                                    {
                                        sumdown = sumdown + 0;
                                        totalowed = totalowed + 0;
                                    }
                                    else
                                    {
                                        sumdown = sumdown + Convert.ToDouble(dr.Cells[4].Value);
                                        totalowed = totalowed + Convert.ToDouble(dr.Cells[7].Value);
                                    }
                                    
                                }
                                if (sumdown > Convert.ToDouble(textBox2.Text))
                                {
                                    MessageBox.Show("Sum of downpayment qty overpass total amount, please verify!!");
                                    sumdown = 0;
                                    totalowed = 0;
                                    foreach (DataGridViewRow dr in dataGridView1.Rows)
                                    {
                                        if (dr.Index!=e.RowIndex)
                                        {
                                            if (Convert.ToString(dr.Cells[4].Value) == string.Empty)
                                            {
                                                sumdown = sumdown + 0;
                                                totalowed = totalowed + 0;
                                            }
                                            else
                                            {
                                                sumdown = sumdown + Convert.ToDouble(dr.Cells[4].Value);
                                                totalowed = totalowed + Convert.ToDouble(dr.Cells[7].Value);
                                            }
                                        }


                                    }                                    
                                    label5.Text = Convert.ToString(Convert.ToDouble(textBox2.Text));
                                    label5.Text = Convert.ToString(Convert.ToDouble(label5.Text) - sumdown);
                                    label6.Text = Convert.ToString(Convert.ToDouble(label6.Text) - totalowed);
                                }
                                else
                                {
                                    dataGridView1.Rows[e.RowIndex].Cells[7].Value = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[3].Value) - Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[4].Value);
                                    dataGridView1.Rows[e.RowIndex].Cells[5].Value = dateTimePicker1.Value;
                                    dataGridView1.Rows[e.RowIndex].Cells[6].Value = textBox1.Text;

                                    
                                    xlWorkSheet.Cells[dataGridView1.Rows[e.RowIndex].Cells[8].Value, 5].value = dataGridView1.Rows[e.RowIndex].Cells[4].Value;
                                    xlWorkSheet.Cells[dataGridView1.Rows[e.RowIndex].Cells[8].Value, 6].value = dataGridView1.Rows[e.RowIndex].Cells[5].Value;  // THIRD COLUMN.
                                    xlWorkSheet.Cells[dataGridView1.Rows[e.RowIndex].Cells[8].Value, 7].value = dataGridView1.Rows[e.RowIndex].Cells[6].Value;  // SECOND COLUMN.
                                    xlWorkSheet.Cells[dataGridView1.Rows[e.RowIndex].Cells[8].Value, 8].value = dataGridView1.Rows[e.RowIndex].Cells[7].Value;  // THIRD COLUMN.    

                                    //xlWorkSheet.Cells[e.RowIndex + 2, 5].value = dataGridView1.Rows[e.RowIndex].Cells[4].Value;
                                    //xlWorkSheet.Cells[e.RowIndex + 2, 6].value = dataGridView1.Rows[e.RowIndex].Cells[5].Value;  // THIRD COLUMN.
                                    //xlWorkSheet.Cells[e.RowIndex + 2, 7].value = dataGridView1.Rows[e.RowIndex].Cells[6].Value;  // SECOND COLUMN.
                                    //xlWorkSheet.Cells[e.RowIndex + 2, 8].value = dataGridView1.Rows[e.RowIndex].Cells[7].Value;  // THIRD COLUMN.    
                                    label5.Text = Convert.ToString(Convert.ToDouble(textBox2.Text));
                                    label5.Text = Convert.ToString(Convert.ToDouble(label5.Text) - sumdown);
                                    totalowed = totalowed - Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells[7].Value);
                                    label6.Text = Convert.ToString(Convert.ToDouble(label6.Text) - totalowed);
                                }
                            }                            
                            
                        }

                    }
                    
                }
                xlWorkBook.Save();
                xlWorkBook.Close();
                xlApp.Quit();

                // CLEAN UP.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
            }
        }
        private void AutomaticRolled() 
        {
            try
            {
                double amountToBeRolled = Convert.ToDouble(textBox2.Text);
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlWorkBook = xlApp.Workbooks.Open(sFileName);   // WORKBOOK TO OPEN THE EXCEL FILE.
                xlWorkSheet = xlWorkBook.Worksheets["Details"];  // THE SHEET WITH THE DATA.

                foreach (DataGridViewRow dr in dataGridView1.Rows) 
                {
                    if (Convert.ToDouble(dr.Cells[7].Value) > 0 && amountToBeRolled>0) 
                    {
                        if (Convert.ToString(dr.Cells[4].Value) == string.Empty) 
                        {
                            if (amountToBeRolled >= Convert.ToDouble(dr.Cells[7].Value)) 
                            {
                                dr.Cells[4].Value = dr.Cells[7].Value;
                                dr.Cells[7].Value = "0";
                                dr.Cells[5].Value = dateTimePicker1.Value;
                                dr.Cells[6].Value = textBox1.Text;
                                amountToBeRolled = amountToBeRolled - Convert.ToDouble(dr.Cells[4].Value);
                            }
                            else 
                            {
                                dr.Cells[4].Value = amountToBeRolled;
                                dr.Cells[7].Value = Convert.ToDouble(dr.Cells[7].Value) - amountToBeRolled;
                                dr.Cells[5].Value = dateTimePicker1.Value;
                                dr.Cells[6].Value = textBox1.Text;
                                amountToBeRolled = amountToBeRolled - amountToBeRolled;
                            }


                        }
                        else 
                        {
                            if (amountToBeRolled >= Convert.ToDouble(dr.Cells[7].Value))
                            {
                                dr.Cells[4].Value = dr.Cells[4].Value+"|"+ dr.Cells[7].Value;                                
                                dr.Cells[5].Value = dr.Cells[5].Value + "|" + dateTimePicker1.Value;
                                dr.Cells[6].Value = dr.Cells[6].Value + "|" + textBox1.Text;
                                amountToBeRolled = amountToBeRolled - Convert.ToDouble(dr.Cells[7].Value);
                                dr.Cells[7].Value = "0";
                            }
                            else
                            {
                                dr.Cells[4].Value = dr.Cells[4].Value + "|" + amountToBeRolled;
                                dr.Cells[7].Value = Convert.ToDouble(dr.Cells[7].Value) - amountToBeRolled;
                                dr.Cells[5].Value = dr.Cells[5].Value + "|" + dateTimePicker1.Value;
                                dr.Cells[6].Value = dr.Cells[6].Value + "|" + textBox1.Text;
                                amountToBeRolled = amountToBeRolled - amountToBeRolled;
                            }
                        }
                        xlWorkSheet.Cells[dr.Cells[8].Value, 5].value = dr.Cells[4].Value;
                        xlWorkSheet.Cells[dr.Cells[8].Value, 6].value = dr.Cells[5].Value;
                        xlWorkSheet.Cells[dr.Cells[8].Value, 7].value = dr.Cells[6].Value;
                        xlWorkSheet.Cells[dr.Cells[8].Value, 8].value = dr.Cells[7].Value;


                    }

                }
                label6.Text = Convert.ToString(Convert.ToDouble(label6.Text) - Convert.ToDouble(textBox2.Text));
                label5.Text = "0.00";
                xlWorkBook.Save();
                xlWorkBook.Close();
                xlApp.Quit();

                // CLEAN UP.
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);

            }
            catch (Exception e)
            {

                throw new Exception(e.Message.ToString());
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text!=string.Empty && textBox2.Text != string.Empty) 
            {
                if (Convert.ToDouble(textBox2.Text) >Convert.ToDouble(label6.Text)) 
                {
                    MessageBox.Show("Downpayment amount can not be greater than owed amount!!");
                }
                else
                {
                    if (checkBox1.Checked == false) 
                    {
                        dataGridView1.Enabled = true;
                        button1.Enabled = false;
                        textBox1.Enabled = false;
                        textBox2.Enabled = false;
                        dateTimePicker1.Enabled = false;
                        label5.Text = textBox2.Text;
                    }
                    else 
                    {
                        button1.Enabled = false;
                        textBox1.Enabled = false;
                        textBox2.Enabled = false;
                        dateTimePicker1.Enabled = false;
                        AutomaticRolled();
                    }
                    
                }
                

            }
            else
            {
                MessageBox.Show("please fill Downpayment ID and amount");
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            textBox2.Text = string.Empty;
            dataGridView1.Enabled = false;
            button1.Enabled = true;
            Excel2Grid(sFileName);
            
            textBox1.Enabled = true;
            textBox2.Enabled = true;
            dateTimePicker1.Enabled = true;
        }



    }
}
