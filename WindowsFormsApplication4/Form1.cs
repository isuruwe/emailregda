using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Drawing.Printing;
using Microsoft.Reporting.WinForms;
using System.Configuration;

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        int printRow = 0;
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();
      
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'ERMSDataSet.ERMSData' table. You can move, or remove it, as needed.
            //this.ERMSDataTableAdapter.Fill(this.ERMSDataSet.ERMSData);
            this.dateTimePicker1.Value = DateTime.Now.Date;
            this.dateTimePicker2.Value = DateTime.Now.Date;
            dt.Columns.Add("ID");
            dt.Columns.Add("Sender");
            dt.Columns.Add("MessageSubject");
            dt.Columns.Add("Timestamp");
           
            //this.reportViewer1.RefreshReport();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            String fdate = dateTimePicker1.Value.Date.Date.ToShortDateString().ToString();
            String sdate = dateTimePicker1.Value.ToString();
            String adate = fdate+ " 23:59:59 PM";
            String emnai = ConfigurationManager.AppSettings["setting1"];
            int cnt=1;
            string un = @"slaf\mailadmin";
            System.Security.SecureString pw = new System.Security.SecureString();
            string password = "mail@123";
            string databaseName = "dit";
            string exchangeServerName = "http://EXCHANGE-HA-02.slaf.int/powershell";
            //win-6k7ld7joseu
            string microsoftSchemaName = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
            foreach (char ch in password)
            {
                pw.AppendChar(ch);
            }
            PSCredential cred = new PSCredential(un, pw);
            DataTable dstIntData = new DataTable();
            if (dstIntData!=null)
            {
                dstIntData.Clear();
            }
            if (ds != null)
            {
                ds.Tables.Clear();
            }
            if (dt != null)
            {
                dt.Rows.Clear();
            }
            progressBar1.Visible = true;
            progressBar1.Style = ProgressBarStyle.Marquee;
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(new Uri(exchangeServerName), microsoftSchemaName, cred);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;
            // dataGridView1.Rows.Clear();
            String cmbtest = comboBox1.SelectedItem.ToString();

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                using (PowerShell powershell = PowerShell.Create())
                {
                    powershell.AddCommand("Get-MessageTrackingLog");
                    powershell.AddParameter("eventid", cmbtest);
                    powershell.AddParameter("Start", sdate);
                    powershell.AddParameter("End", adate);
                    powershell.AddParameter("Recipient", emnai);
                    powershell.AddCommand("select-object");
                    var props = new string[] { "Sender", "MessageSubject", "TimeStamp" };
                    powershell.AddParameter("property", props);
                    runspace.Open();
                    powershell.Runspace = runspace;
                   
                    Collection<PSObject> results = powershell.Invoke();
                   
                    
                    foreach (PSObject result in results)
                    {

                      

                        DataRow dr = dt.NewRow();
                        dr["ID"] = cnt.ToString();
                        foreach (PSPropertyInfo prop in result.Properties)
                            {
                                var count = prop.Name;
                                var name = prop.Value;
                           
                            //MessageBox.Show(count.ToString());
                            //In other words generate output as you desire.
                            // dt.Rows.Add(count, name);
                            if (count == "Sender")
                            {
                                 //Create New Row
                                dr["Sender"] = name;
                                
                            }
                            if (count == "MessageSubject")
                            {
                                //Create New Row
                                dr["MessageSubject"] = name;
                                
                            }
                            
                            if (count == "Timestamp")
                            {
                                //Create New Row
                                dr["Timestamp"] = name;
                               
                              
                            }
                           
                            // Set Column Value
                           
                        }
                        dt.Rows.Add(dr);
                        cnt++;





                    }

                   
                    ds.Tables.Add(dt);

                    dstIntData = ds.Tables[0];
                    //dstIntData = oBOCBF_TXN_DETAIL.SelectForSavingsReport(null, this.dateTimePicker1.Text);
                    reportViewer1.LocalReport.DataSources.Clear();
                    ReportDataSource IntData_rsource = new ReportDataSource("DataSet1", dstIntData);
                    reportViewer1.LocalReport.DataSources.Add(IntData_rsource);

                    this.reportViewer1.RefreshReport();
                    progressBar1.Visible = false;


                    //for (int count = 0; count < ds.Tables[0].Rows.Count; count++)
                    //{
                    //    dataGridView1.Rows.Add();
                    //    dataGridView1.Rows[count].Cells["Sender"].Value = ds.Tables[0].Rows[count]["Sender"].ToString();
                    //    dataGridView1.Rows[count].Cells["MessageSubject"].Value = ds.Tables[0].Rows[count]["MessageSubject"].ToString();
                    //    dataGridView1.Rows[count].Cells["Timestamp"].Value = ds.Tables[0].Rows[count]["Timestamp"].ToString();
                       

                    //}
                }
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
           // printDocument1.Print();
            PrintDocument document = new PrintDocument();
            document.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);

            PrintPreviewDialog ppDialog = new PrintPreviewDialog();
            ppDialog.Document = document;
            ppDialog.Show();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            // Bitmap bm =new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
            //dataGridView1.DrawToBitmap(bm, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));
            //e.Graphics.DrawImage(bm, 0, 0);

            PrintDocument document = (PrintDocument)sender;
            Graphics g = e.Graphics;

            Brush brush = new SolidBrush(Color.Black);
            Pen pen = new Pen(brush);
            Font font = new Font("Arial", 10, FontStyle.Bold);
            Font fonte = new Font("Arial", 15, FontStyle.Bold);
            int x = 0, y = 0, width = 300, height = 30;

            SizeF sizeeee = g.MeasureString("TIME :: ", fonte);
            float xPaddingeee = (width - sizeeee.Width) / 2;
            g.DrawString("TIME :: ", fonte, brush, x + xPaddingeee, y + 5);
            x += width;


            SizeF sizee = g.MeasureString(DateTime.Now.ToString(), fonte);
            float xPaddinge = (width - sizee.Width) / 2;

            g.DrawString(DateTime.Now.ToString(), fonte, brush, x + xPaddinge, y + 5);
            x += width;

            for (int kk = 0; kk < 2; kk++)
            {
                SizeF sizeee = g.MeasureString("", font);
                float xPaddingee = (width - sizee.Width) / 2;

                g.DrawString("", font, brush, x + xPaddingee, y + 5);
                x += width;
            }
            x = 0;
            y += 60;


            foreach (DataColumn column in dt.Columns)
            {
                g.DrawRectangle(pen, x, y, width, height);
                SizeF size = g.MeasureString(column.ColumnName, fonte);
                float xPadding = (width - size.Width) / 2;

                g.DrawString(column.ColumnName, fonte, brush, x + xPadding, y + 5);
                x += width;
            }


            x = 0;
            y += 30;
            int columnCount = dt.Columns.Count;

            foreach (DataRow row in dt.Rows)
            {
                for (int i = 0; i < columnCount; i++)
                {
                    g.DrawRectangle(pen, x, y, width, height);
                    SizeF size = g.MeasureString(row[i].ToString(), font);
                    float xPadding = (width - size.Width) / 2;

                    g.DrawString(row[i].ToString(), font, brush, x + xPadding, y + 5);
                    x += width;
                }
                x = 0;
                y += 30;
            }


        }
    }
}
