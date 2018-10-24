using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public int FldClient { get; set; }

        public int FldMatter { get; set; }


        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();
        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

            for (int i = 0; i < (checkedListBox1.Items.Count); i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);
            string SQL = "";
            DataSet batches;
            string items = "";
            if (checkedListBox1.Items.Count > 0) //did they select at least one checkbox?
            {
                int total = checkedListBox1.Items.Count;
                for (int i = 0; i < (checkedListBox1.Items.Count); i++)
                {
                    if (checkedListBox1.GetItemChecked(i))
                    {
                        switch (i)
                        {
                            case 0: //cash receipt
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct crbbatchnbr from CashReceiptsBatch where crbreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["crbbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=5300 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from CashReceiptsBatch   where crbreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL); 
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 1: //Check
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct cbbatchnbr  from CheckBatch where cbreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["cbbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=7300 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from CheckBatch where cbreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 2: //Credit Memo
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct cmbbatchnbr from CreditMemoBatch where cmbreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["cmbbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=5200 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from CreditMemoBatch where cmbreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 3: //Expense
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct ebbatchnbr from ExpenseBatch where ebreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["ebbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=5000 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from ExpenseBatch where ebreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 4: //Journal Entry
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct jebbatchnbr  from JEBatch   where jebreccount=0 and jebbatchnbr not in (select arpjebatchnbr from arpostbatch)");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["jebbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=4700 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from JEBatch   where jebreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 5: //Manual Bill
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct mbbbatchnbr from ManualBillBatch where mbbreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["mbbbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=5100 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from ManualBillBatch where mbbreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 6: //Time Batch
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct tbbatchnbr from TimeBatch where tbreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["tbbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');
                                    SQL = "delete from TimeBatchImportError where TBIEBatchNbr in (" + items + ")";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from TimeBatchDetail where TBDBatch in (" + items + ")";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from documenttree where dtdocclass=4900 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from TimeBatch where tbreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 7: //Trust Adjustment
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct tabbatchnbr from TrAdjBatch where tabreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["tabbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=7500 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from TrAdjBatch where tabreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;
                            case 8: //Voucher
                                items = "";
                                batches = _jurisUtility.RecordsetFromSQL("select distinct vbbatchnbr from VoucherBatch where vbreccount=0");
                                if (batches.Tables[0].Rows.Count != 0)
                                {
                                    foreach (DataRow dr in batches.Tables[0].Rows)
                                        items = items + dr["vbbatchnbr"].ToString() + ",";

                                    items = items.TrimEnd(',');

                                    SQL = "delete from documenttree where dtdocclass=7200 and dtkeyL in (" + items + ") and dtdoctype = 'R'";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                    SQL = "delete from VoucherBatch where vbreccount=0";
                                    _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                                }
                                UpdateStatus("Updating database...", i, total);
                                batches.Clear();
                                break;

                        }
                    }
                }

                UpdateStatus("Updating database...", total, total);
            }
            else
                MessageBox.Show("At least one checkbox needs to be selected");

        }
        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            if (areAnyCheckBoxesChecked())
                DoDaFix();
            else
                MessageBox.Show("Please select at least one checkbox", "Selection error", MessageBoxButtons.OK, MessageBoxIcon.Hand);




        }

        private void buttonReport_Click(object sender, EventArgs e)
        {
           // if (string.IsNullOrEmpty(toAtty) || string.IsNullOrEmpty(fromAtty))
          //      MessageBox.Show("Please select from both Timekeeper drop downs", "Selection Error");
          //  else
          //  {
                //generates output of the report for before and after the change will be made to client
            if (areAnyCheckBoxesChecked())
            {
                string SQLTkpr = getReportSQL();
                    DataSet myRSTkpr = _jurisUtility.RecordsetFromSQL(SQLTkpr);

                    ReportDisplay rpds = new ReportDisplay(myRSTkpr);
                    rpds.Show();
            }
            else
                MessageBox.Show("Please select at least one checkbox", "Selection error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
        }

        private string getReportSQL()
        {
            string reportSQL = "";
            for (int i = 0; i <= (checkedListBox1.Items.Count - 1); i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    switch (i)
                    {
                        case 0:
                            reportSQL = reportSQL + " select crbbatchnbr as BatchNbr, crbcomment as BatchComment, crbstatus as BatchStatus, crbdateentered as DateEntered, 'Cash Receipt Batch' as BatchType from CashReceiptsBatch where crbreccount=0 union all";
                            break;
                        case 1:
                            reportSQL = reportSQL + " select cbbatchnbr as BatchNbr, cbcomment as BatchComment, cbstatus as BatchStatus,cbdate as DateEntered, 'Check Batch' as BatchType from CheckBatch where cbreccount=0 union all";
                            break;
                        case 2:
                            reportSQL = reportSQL + " select cmbbatchnbr as BatchNbr, cmbcomment as BatchComment, cmbstatus as BatchStatus,cmbdateentered as DateEntered, 'Credit Memo Batch' as BatchType  from CreditMemoBatch where cmbreccount=0 union all";
                            break;
                        case 3:
                            reportSQL = reportSQL + " select ebbatchnbr as BatchNbr, ebcomment as BatchComment, ebstatus as BatchStatus, ebdateentered as DateEntered, 'Expense Batch' as BatchType  from ExpenseBatch where ebreccount=0 union all";
                            break;
                        case 4:
                            reportSQL = reportSQL + " select jebbatchnbr as BatchNbr, jebcomment as BatchComment, jebstatus as BatchStatus, jebentereddate  as DateEntered, 'Journal Entry Batch' as BatchType  from JEBatch where jebreccount=0 and jebbatchnbr not in (select arpjebatchnbr from arpostbatch) union all";
                            break;
                        case 5:
                            reportSQL = reportSQL + " select mbbbatchnbr as BatchNbr, mbbcomment as BatchComment, mbbstatus as BatchStatus, mbbdateentered  as DateEntered, 'Manual Bill Batch' as BatchType  from ManualBillBatch   where mbbreccount=0 union all";
                            break;
                        case 6:
                            reportSQL = reportSQL + " select tbbatchnbr as BatchNbr, tbcomment as BatchComment, tbstatus as BatchStatus, tbdateentered  as DateEntered, 'Time Batch' as BatchType  from TimeBatch   where tbreccount=0 union all";
                            break;
                        case 7:
                            reportSQL = reportSQL + " select tabbatchnbr as BatchNbr, tabcomment as BatchComment, tabstatus as BatchStatus, tabdate  as DateEntered, 'Trust Adjustment Batch' as BatchType  from TrAdjBatch   where tabreccount=0 union all";
                            break;
                        case 8:
                            reportSQL = reportSQL + " select vbbatchnbr as BatchNbr, vbcomment as BatchComment, vbstatus as BatchStatus, vbdate  as DateEntered, 'Voucher Batch' as BatchType  from VoucherBatch   where vbreccount=0 union all";
                            break;


                    }
                }

            }
            if (!string.IsNullOrEmpty(reportSQL))
                reportSQL = reportSQL.Substring(0, reportSQL.Length - 9);
            return reportSQL;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private bool areAnyCheckBoxesChecked()
        {
            if (checkedListBox1.CheckedIndices.Count > 0)
                return true;
            else
                return false;


        }


    }
}
