using System;
using System.Collections.Generic;
using System.Configuration;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Collections.Specialized;

namespace PayrollConverter
{
         public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // set the headers in grid 2 for the SOS Import Nominal Postings file
            dataGridView1.Columns.Add("BR", "Branch"); // col 0
            dataGridView1.Columns.Add("DEPT", "Dept"); // col 63 split 1
            dataGridView1.Columns.Add("PERIOD", "Period"); // col 5
            dataGridView1.Columns.Add("PRE_TAX", "Pre Tax Add/Ded");
            dataGridView1.Columns.Add("GU", "GUCosts");
            dataGridView1.Columns.Add("ABS_PAY", "Absence Pay");
            dataGridView1.Columns.Add("HOL_PAY", "Holiday Pay");
            dataGridView1.Columns.Add("PRE_PEN", "Pre TaxPension");
            dataGridView1.Columns.Add("TAX_PAY", "Taxable Pay");
            dataGridView1.Columns.Add("TAX", "Tax");
            dataGridView1.Columns.Add("NET_EE_NI", "Net Ee NI");
            dataGridView1.Columns.Add("POST_TAX", "Post TaxAdd/Ded");
            dataGridView1.Columns.Add("POST_PEN", "Post Tax Pension");
            dataGridView1.Columns.Add("AEO", "AEO");
            dataGridView1.Columns.Add("SLOAN", "StudentLoans");
            dataGridView1.Columns.Add("NET_PAY", "Net Pay");
            dataGridView1.Columns.Add("NET_ER_NI", "Net Er NIincl Class 1A");
            dataGridView1.Columns.Add("ER_PEN", "Er Pension");

            dataGridView1.Columns["BR"].Frozen = true;
            dataGridView1.Columns["DEPT"].Frozen = true;

            // set the headers in grid 2 for the SOS Import Nominal Postings file
            dataGridView2.Columns.Add("BR","BR");
            dataGridView2.Columns.Add("CODE", "CODE");
            dataGridView2.Columns.Add("NARRATIVE", "NARRATIVE");
            dataGridView2.Columns.Add("DR", "DR");
            dataGridView2.Columns.Add("CR","CR");
            dataGridView2.Columns.Add("VC","VC");
            dataGridView2.Columns.Add("VAT","VAT");
            dataGridView2.Columns.Add("FE","FE");
            dataGridView2.Columns.Add("MT","MT");
            dataGridView2.Columns.Add("REF","REF");

            // Get Values for Dictionary from Config File

            var nomPaymentCodes = ConfigurationManager.GetSection("NominalCodes/NominalPaymentCodes") as NameValueCollection;
            foreach (var key in nomPaymentCodes.AllKeys)
            {
                NominalPayment.nomCodes.Add(key, nomPaymentCodes[key]);
            }

            var nomReceiptCodes = ConfigurationManager.GetSection("NominalCodes/NominalReceiptCodes") as NameValueCollection;
            foreach(var key in nomReceiptCodes.AllKeys)
            {
                NominalReceipt.nomCodes.Add(key, nomReceiptCodes[key]);
            }

            var departmentCodes = ConfigurationManager.GetSection("DepartmentCodes") as NameValueCollection;
            foreach (var key in departmentCodes.AllKeys)
            {
                FEMT.deptCodes.Add(key, departmentCodes[key].Split(','));
            }

        }

        static class FEMT
        {
            public static Dictionary<string, string[]> deptCodes = new Dictionary<string, string[]>()
            {
                /* REMOVED - Replaced with call to App.Config - See Initization
                {"1", new string[] {"COMM","CORP"} },
                {"2", new string[] {"COMM","CLI1"} },
                {"4", new string[] {"COMM","CLI2"} },
                {"5", new string[] {"PROP","PRO1"} },
                {"6", new string[] {"PROP","RES1"} },
                {"7", new string[] {"PROP","PRO3"} },
                {"9", new string[] {"PRIV","PRI2"} },
                {"10", new string[] {"FIN","ZIFA"} },
                {"11", new string[] {"ZTRS","TRS"} },
                {"12", new string[] {"SUPP","ADMI"} },
                {"13", new string[] {"SUPP","MKTG"} },
                {"14", new string[] {"SUPP","IT"} },
                {"15", new string[] {"SUPP","ACCS"} },
                {"16", new string[] {"SUPP","PERS"} },
                {"17", new string[] {"PRIV","PRI1"} },
                {"19", new string[] {"PRIV","FAM"} },
                {"20", new string[] {"PRIV","INH"} },
                {"22", new string[] {"PRIV","FAM3"} },
                {"BLHR & Employment", new string[] {"ZHR","ZHR1"}},
                {"IP", new string[] {"COMM","ZIP"} }*/
            };
        }

        static class NominalPayment
        {
            public static Dictionary<string, string> nomCodes = new Dictionary<string, string>()
            {
                /* REMOVED - Replaced with call to App.Config - See Initization
                {"PRE_TAX", "200100"}, // Do not report on the Pre Tax Add/Ded column
                {"ABS", "100100"},
                {"TAX_PAY", "200100"},
                {"ER_PEN_NP", "200300"},
                {"NET_ER_NI_NP", "200200"}*/
            };
        }

        static class NominalReceipt
        {
            public static Dictionary<string, string> nomCodes = new Dictionary<string, string>()
            {
                /*  REMOVED - Replaced with call to App.Config - See Initization
                {"SLOAN", "800100"},
                {"NET_ER_NI", "800100"},
                {"NET_EE_NI", "800100"},
                {"TAX", "800100"},
                {"ER_PEN", "800300"}, // this is also 200100 so build a rule for this specific column
                {"POST_PEN", "800300"},
                {"NET_PAY", "800200"}*/
            };

        }

        static class ContraTotals
        {
            public static decimal NP_Contra = 0;
            public static decimal NR_Contra = 0;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = openFileDialog1.FileName;
                try
                {
                    // Get the data.
                    string[,] values = LoadCsv(file);
                    int num_rows = values.GetUpperBound(0) + 1;
                    int num_cols = values.GetUpperBound(1) + 1;
                    // Display the data to show we have it.

                    // Set the iterator, on col 63 "n Cost Ctr Total This Period"
                    string newDept = "";
                    int dGV1RowCount = 0;

                    // Add the data. (reset to 1 if data file contains headers)
                    for (int r = 0; r < num_rows; r++)
                    {

                        if(values[r, 63] == newDept) 
                        {
                            continue;
                        }
                       
                        dataGridView1.Rows.Add();

                        newDept = values[r, 63].ToString();
                        
                        /*for (int c = 0; c < num_cols; c++)
                        {
                            //dataGridView1.Rows[r - 1].Cells[c].Value = values[r, c];
                            dataGridView1.Rows[r].Cells[c].Value = values[r, c];
                        }*/
                        dataGridView1.Rows[dGV1RowCount].Cells[0].Value = values[r, 0];
                        int deptNameLength = values[r, 63].ToString().Length;
                        dataGridView1.Rows[dGV1RowCount].Cells[1].Value = values[r, 63].ToString().Substring(0,deptNameLength - 27);
                        dataGridView1.Rows[dGV1RowCount].Cells[2].Value = values[r, 5];

                        dataGridView1.Rows[dGV1RowCount].Cells[3].Value = values[r, 64];
                        dataGridView1.Rows[dGV1RowCount].Cells[4].Value = values[r, 65];
                        dataGridView1.Rows[dGV1RowCount].Cells[5].Value = values[r, 66];
                        dataGridView1.Rows[dGV1RowCount].Cells[6].Value = values[r, 67];
                        dataGridView1.Rows[dGV1RowCount].Cells[7].Value = values[r, 68];
                        dataGridView1.Rows[dGV1RowCount].Cells[8].Value = values[r, 69];
                        dataGridView1.Rows[dGV1RowCount].Cells[9].Value = values[r, 70];
                        dataGridView1.Rows[dGV1RowCount].Cells[10].Value = values[r, 71];
                        dataGridView1.Rows[dGV1RowCount].Cells[11].Value = values[r, 72];
                        dataGridView1.Rows[dGV1RowCount].Cells[12].Value = values[r, 73];
                        dataGridView1.Rows[dGV1RowCount].Cells[13].Value = values[r, 74];
                        dataGridView1.Rows[dGV1RowCount].Cells[14].Value = values[r, 75];
                        dataGridView1.Rows[dGV1RowCount].Cells[15].Value = values[r, 76];
                        dataGridView1.Rows[dGV1RowCount].Cells[16].Value = values[r, 77];
                        dataGridView1.Rows[dGV1RowCount].Cells[17].Value = values[r, 78];

                        dGV1RowCount = dGV1RowCount + 1; // increment view 1 row count

                    }

                }
                catch (IOException)
                {
                }

                // format and write the data to the 2nd data grid
                int rowIndex = 0; //index of the row
                foreach (DataGridViewRow dr in dataGridView1.Rows)
                {
                    if (dr.IsNewRow) // <-- Gets a value indicating whether the row is the row for new records
                        continue;

                    // common values all SOS Payroll Journal Lines
                    string sosBR = "001";
                    string sosNarrative = "PAYROLL JOURNAL";
                    string sosFE = "";
                    string sosMT = "";

                    if (FEMT.deptCodes.ContainsKey(dataGridView1.Rows[dr.Index].Cells["DEPT"].Value.ToString()))
                    {
                        sosFE = FEMT.deptCodes[dataGridView1.Rows[dr.Index].Cells["DEPT"].Value.ToString()][0];
                        sosMT = FEMT.deptCodes[dataGridView1.Rows[dr.Index].Cells["DEPT"].Value.ToString()][1];
                    }
                    else
                    {
                        sosFE = dataGridView1.Rows[dr.Index].Cells["DEPT"].Value.ToString();
                    }

                    // 21-Nov: spoke to Wendy and we do not need to import this column, only the TAX_PAY column
                    //rowIndex = dataGridView2.Rows.Add();                
                    //CreateExportGridRecord(dataGridView2, "PRE_TAX", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "GU", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "GU", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "ABS_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "ABS_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "HOL_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "HOL_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "PRE_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "PRE_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "TAX_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "TAX_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "TAX", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "TAX", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "NET_EE_NI", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "NET_EE_NI", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "POST_TAX", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "POST_TAX", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "POST_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "POST_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "AEO", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "AEO", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "SLOAN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "SLOAN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "NET_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "NET_PAY", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    // Employer Pension and NI payments appear twice! The first set is for Nominal Payments of ER related costs "true"
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "NET_ER_NI", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, true);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "NET_ER_NI", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, true);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "ER_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, true);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "ER_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, true);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "NET_ER_NI", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "NET_ER_NI", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord(dataGridView2, "ER_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);
                    rowIndex = dataGridView2.Rows.Add();
                    CreateExportGridRecord800350(dataGridView2, "ER_PEN", rowIndex, dataGridView1, dr.Index, sosFE, sosMT, sosNarrative, sosBR, false);

                }

            } 

        }

        private string[,] LoadCsv(string filename)
        {
            // Get the file's text.
            string whole_file = System.IO.File.ReadAllText(filename);

            // Split into lines.
            whole_file = whole_file.Replace('\n', '\r');
            string[] lines = whole_file.Split(new char[] { '\r' },
                StringSplitOptions.RemoveEmptyEntries);

            // Use regex split to irnore commas imbedded into double quotes

            // See how many rows and columns there are.
            int num_rows = lines.Length;
            //int num_cols = lines[0].Split(',').Length;
            int num_cols = Regex.Matches(lines[0], ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)").Count;

            // Allocate the data array.
            string[,] values = new string[num_rows, num_cols];

            // Load the array.
            for (int r = 0; r < num_rows; r++)
            {
                //string[] line_r = lines[r].Split(',');
                string[] line_r = Regex.Split(lines[r], ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                for (int c = 0; c < num_cols; c++)
                {
                    values[r, c] = line_r[c];
                }
            }

            // Return the values.
            return values;
        }

        private void CreateExportGridRecord(DataGridView DGV2, string reference, int rowIndex2, DataGridView DGV1, int rowIndex1, string FE, string MT, string NARRATIVE, string BR, bool NPFlag)
        {
            // Note, all ER values are added as a NR and PR

            decimal amount;
            string nominalCodeLookup;

            if (NPFlag == true) 
            {
                nominalCodeLookup = reference + "_NP";
            } else
            {
                nominalCodeLookup = reference;
            }

            DGV2.Rows[rowIndex2].Cells["BR"].Value = BR;
            DGV2.Rows[rowIndex2].Cells["NARRATIVE"].Value = NARRATIVE;

            DGV2.Rows[rowIndex2].Cells["REF"].Value = DGV1.Columns[reference].HeaderText;

            if (NominalPayment.nomCodes.ContainsKey(nominalCodeLookup)) 
            {
                DGV2.Rows[rowIndex2].Cells["CODE"].Value = NominalPayment.nomCodes[nominalCodeLookup].ToString();
                DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Green;
                DGV2.Rows[rowIndex2].Cells["FE"].Value = FE;
                DGV2.Rows[rowIndex2].Cells["MT"].Value = MT;

                // ToDo: KC this isn't quite right sure NPs can have DRs and CRs?

                    DGV2.Rows[rowIndex2].Cells["DR"].Value = DGV1.Rows[rowIndex1].Cells[reference].Value;
                
                // override color is value is ZERO eslse update the NP Total
                if (Decimal.TryParse(DGV1.Rows[rowIndex1].Cells[reference].Value.ToString(), out amount))
                {
                    if (amount == 0)
                    {
                        DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Gray;
                    } else 
                    {
                        ContraTotals.NP_Contra = ContraTotals.NP_Contra + amount;
                        NP_ContraTotal.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-GB"),"{0:C}", ContraTotals.NP_Contra);
                    }
                }

            } else
            {
                if (NominalReceipt.nomCodes.ContainsKey(nominalCodeLookup)) 
                {
                    DGV2.Rows[rowIndex2].Cells["CODE"].Value = NominalReceipt.nomCodes[nominalCodeLookup].ToString();
                    DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.PaleVioletRed;

                    DGV2.Rows[rowIndex2].Cells["CR"].Value = DGV1.Rows[rowIndex1].Cells[reference].Value;
                    // override color is value is ZERO.
                    if (Decimal.TryParse(DGV1.Rows[rowIndex1].Cells[reference].Value.ToString(), out amount))
                    {
                        if (amount == 0)
                        {
                            DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Gray;
                        } else 
                        {
                            ContraTotals.NR_Contra = ContraTotals.NR_Contra  + amount;
                            NR_ContraTotal.Text = string.Format(System.Globalization.CultureInfo.CreateSpecificCulture("en-GB"), "{0:C}", ContraTotals.NR_Contra);
                        }
                    }

                } else 
                {
                    DGV2.Rows[rowIndex2].Cells["CODE"].Value = "XXXXXX";
                    DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Gray;
                }
            }
        }

        private void CreateExportGridRecord800350(DataGridView DGV2, string reference, int rowIndex2, DataGridView DGV1, int rowIndex1, string FE, string MT, string NARRATIVE, string BR, bool NPFlag)
        {
            // Note, all ER values are added as a NR and PR
            // for every value posted to NP or NR we need to post a dummy posting to account O099
            // KC: 23/2/2020 - Wendy reports that we can not longer use O099 "theres an accounting reason" so switch to 800350

            decimal amount;
            string nominalCodeLookup;

            if (NPFlag == true)
            {
                nominalCodeLookup = reference + "_NP";
            }
            else
            {
                nominalCodeLookup = reference;
            }

            DGV2.Rows[rowIndex2].Cells["BR"].Value = BR;
            DGV2.Rows[rowIndex2].Cells["NARRATIVE"].Value = NARRATIVE;

            DGV2.Rows[rowIndex2].Cells["REF"].Value = DGV1.Columns[reference].HeaderText;

            if (NominalPayment.nomCodes.ContainsKey(nominalCodeLookup))
            {
                DGV2.Rows[rowIndex2].Cells["CODE"].Value = "800350";
                DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Green;
                DGV2.Rows[rowIndex2].Cells["FE"].Value = FE;
                DGV2.Rows[rowIndex2].Cells["MT"].Value = MT;

                // ToDo: KC this isn't quite right sure NPs can have DRs and CRs?

                DGV2.Rows[rowIndex2].Cells["CR"].Value = DGV1.Rows[rowIndex1].Cells[reference].Value;

                // override color is value is ZERO eslse update the NP Total
                if (Decimal.TryParse(DGV1.Rows[rowIndex1].Cells[reference].Value.ToString(), out amount))
                {
                    if (amount == 0)
                    {
                        DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Gray;
                    }
                }

            }
            else
            {
                if (NominalReceipt.nomCodes.ContainsKey(nominalCodeLookup))
                {
                    DGV2.Rows[rowIndex2].Cells["CODE"].Value = "800350";
                    DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.PaleVioletRed;

                    DGV2.Rows[rowIndex2].Cells["DR"].Value = DGV1.Rows[rowIndex1].Cells[reference].Value;
                    // override color is value is ZERO.
                    if (Decimal.TryParse(DGV1.Rows[rowIndex1].Cells[reference].Value.ToString(), out amount))
                    {
                        if (amount == 0)
                        {
                            DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Gray;
                        }
                    }

                }
                else
                {
                    DGV2.Rows[rowIndex2].Cells["CODE"].Value = "XXXXXX";
                    DGV2.Rows[rowIndex2].DefaultCellStyle.BackColor = Color.Gray;
                }
            }
        }

        private void SaveToCSV(DataGridView DGV)
        {
            string filename = "";
            SaveFileDialog sfd = new SaveFileDialog();
            //sfd.Filter = "CSV (*.csv)|*.csv";
            string todayDate = DateTime.Now.ToString("dd-MM-yyyy");
            sfd.FileName = "SOS-Nominal-Import-" + todayDate;
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Data will be exported and you will be notified when it is ready.");
                if (File.Exists(filename + "-NP.csv"))
                {
                    try
                    {
                        File.Delete(filename + "-NP.csv");
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("It wasn't possible to write the NP data to the disk." + ex.Message);
                    }
                }

                if (File.Exists(filename + "-NR.csv"))
                {
                    try
                    {
                        File.Delete(filename + "-NR.csv");
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show("It wasn't possible to write the NR data to the disk." + ex.Message);
                    }
                }

                int columnCount = DGV.ColumnCount;

                string[] outputNP = new string[DGV.RowCount];
                string[] outputNR = new string[DGV.RowCount];

                int x = 0; // total of all lines written to output
                for (int i = 0; i < DGV.RowCount; i++)
                {
                    // ignore grayed out lines in data grid 2
                    if (DGV.Rows[i].DefaultCellStyle.BackColor != Color.Gray && DGV.Rows[i].DefaultCellStyle.BackColor == Color.Green)
                    {
                        
                        for (int j = 0; j < columnCount; j++)
                        {
                            if (DGV.Rows[i].Cells[j].Value != null) {
                                outputNP[x] += DGV.Rows[i].Cells[j].Value.ToString() + ",";
                            } else
                            {
                                outputNP[x] += ",";
                            }
                        }

                        x++;
                    }
                }

                Array.Resize(ref outputNP, x);

                //sfd.FileName = "SOS-Nominal-Import-NP";

                System.IO.File.WriteAllLines(sfd.FileName+"-NP.csv", outputNP, System.Text.Encoding.UTF8);

                x = 0; // total of all lines written to output
                for (int i = 0; i < DGV.RowCount; i++)
                {
                    // ignore grayed out lines in data grid 2
                    if (DGV.Rows[i].DefaultCellStyle.BackColor != Color.Gray && DGV.Rows[i].DefaultCellStyle.BackColor == Color.PaleVioletRed)
                    {
  
                        for (int j = 0; j < columnCount; j++)
                        {
                            if (DGV.Rows[i].Cells[j].Value != null)
                            {
                                outputNR[x] += DGV.Rows[i].Cells[j].Value.ToString() + ",";
                            }
                            else
                            {
                                outputNR[x] += ",";
                            }
                        }

                        x++;
                    }
                }

                Array.Resize(ref outputNR, x);

                //sfd.FileName = "SOS-Nominal-Import-NR";

                System.IO.File.WriteAllLines(sfd.FileName+"-NR.csv", outputNR, System.Text.Encoding.UTF8);

                MessageBox.Show("Your files was generated and are ready for use.");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SaveToCSV(dataGridView2);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow dr in dataGridView2.Rows)
            {
                if (dr.DefaultCellStyle.BackColor == Color.Gray)
                { 
                    if (dr.Visible.Equals(true)) { dr.Visible = false; } else { dr.Visible = true; }
                }
 
            }
        }
    }
}
