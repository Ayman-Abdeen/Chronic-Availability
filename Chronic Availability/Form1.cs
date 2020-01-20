using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;


namespace Chronic_Availability
{
    public partial class chronic : Form
    {
        public chronic()
        {
            InitializeComponent();
        }

        DateTime startDate;
        DateTime endDate;

        DataTable dt_siteStatus = new DataTable();
        List<SiteStatus> li_siteStatus = new List<SiteStatus>();

        DataTable dt_Acc_plannedSite = new DataTable();
        DataTable dt_Acc_unplannedSite = new DataTable();

        List<Acc_planned> unplannedList = new List<Acc_planned>();
        List<Acc_unplanned> Acc_unplannedList = new List<Acc_unplanned>();

        DataTable dt_2G = new DataTable();
        DataTable dt_3G = new DataTable();


        DataTable dt_TT_Dashboard = new DataTable();
        List<TT_Dashboard> TT_DashboardList = new List<TT_Dashboard>();

        List<rowData> Li_2G = new List<rowData>();
        List<rowData> Li_chronicSite_2G = new List<rowData>();
        List<Excelsheet> Li_excel_2G = new List<Excelsheet>();
        List<string> chronicSiteName2G = new List<string>();
        string reason2G = "Unknown_Reason";

        List<rowData> Li_3G = new List<rowData>();
        List<rowData> Li_chronicSite_3G = new List<rowData>();
        List<Excelsheet> Li_excel_3G = new List<Excelsheet>();
        string reason3G = "Unknown_Reason";

        private void set_date_Click(object sender, EventArgs e)
        {
            startDate = dateTimePickerFrom.Value;
            endDate = dateTimePickerTo.Value;
            this.set_date.Enabled = false;
            this.dateTimePickerFrom.Enabled = false;
            this.dateTimePickerTo.Enabled = false;
        }

        private void siteStatus_Click(object sender, EventArgs e)
        {
            try
            {
                this.siteStatus.Enabled = false;
                String inputFile;
                OpenFileDialog op = new OpenFileDialog();
                op.Filter = "All File | *.*|Excel File |*.xlsx";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    inputFile = op.FileName;
                    OleDbConnection con = new OleDbConnection("provider=microsoft.ACE.OLEDB.12.0;Data Source=" + inputFile + ";Extended Properties='Excel 12.0 XML;HDR=YAS;IMEX=1;MAXSCANROWS=0'");
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "SELECT * from [Site Status$]";
                    con.Open();
                    dt_siteStatus.Load(cmd.ExecuteReader());
                    con.Close();
                    //////////

                    progressBar1.Maximum = dt_siteStatus.Rows.Count; // your loop max no. 
                    progressBar1.Value = 0;

                    for (int i = 0; i < dt_siteStatus.Rows.Count; i++)
                    {
                        SiteStatus ss = new SiteStatus();
                        if (dt_siteStatus.Rows[i][0] != null) { ss.NodeID = dt_siteStatus.Rows[i][0].ToString(); } else { ss.NodeID = ""; }
                        if (dt_siteStatus.Rows[i][1] != null) { ss.controller = dt_siteStatus.Rows[i][1].ToString(); } else { ss.controller = ""; }
                        if (dt_siteStatus.Rows[i][2] != null) { ss.Status = dt_siteStatus.Rows[i][2].ToString(); } else { ss.Status = ""; }
                        if (dt_siteStatus.Rows[i][3] != null) { ss.MarkedROT = dt_siteStatus.Rows[i][3].ToString(); } else { ss.MarkedROT = ""; }
                        if (dt_siteStatus.Rows[i][4] != null) { ss.subcategory = dt_siteStatus.Rows[i][4].ToString(); } else { ss.subcategory = ""; }
                        if (dt_siteStatus.Rows[i][5] != null) { ss.area = dt_siteStatus.Rows[i][5].ToString(); } else { ss.area = ""; }
                        if (dt_siteStatus.Rows[i][6] != null) { ss.Tier = dt_siteStatus.Rows[i][6].ToString(); } else { ss.Tier = ""; }
                        if (dt_siteStatus.Rows[i][7] != null) { ss.vendor = dt_siteStatus.Rows[i][7].ToString(); } else { ss.vendor = ""; }
                        if (dt_siteStatus.Rows[i][32] != null) { ss.ServiceType = dt_siteStatus.Rows[i][32].ToString(); } else { ss.ServiceType = ""; }
                        if (ss.ServiceType == "2G") { ss.Is2G = true; if (ss.NodeID.Right(7).Left(3) == "UPP") { ss.SiteID = "P" + ss.NodeID.Right(7).Right(4); } else { ss.SiteID = Extensions.GetNumbers(ss.NodeID) + "_" + ss.vendor; } } else { ss.Is2G = false; ss.SiteID = ss.NodeID; }
                        if (dt_siteStatus.Rows[i][23] != null) { ss.NumberOfActiveCells = dt_siteStatus.Rows[i][23].ToString(); } else { ss.NumberOfActiveCells = ""; }

                        li_siteStatus.Add(ss);

                        progressBar1.Value += 1;

                    }

                    MessageBox.Show("completed");
                    progressBar1.Value = 0;
                    dgvSS.DataSource = Extensions.ConvertToDataTable(li_siteStatus);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                progressBar1.Value = 0;
                this.siteStatus.Enabled = true;
            }
        }

        private void accOutage_Click(object sender, EventArgs e)
        {
            try
            {


                progressBar1.Maximum = 21; // your loop max no. 
                progressBar1.Value = 0;

                String inputFile;
                OpenFileDialog op = new OpenFileDialog();
                op.Filter = "All File | *.*|Excel File |*.xlsx";
                if (op.ShowDialog() == DialogResult.OK)
                {

                    inputFile = op.FileName;
                    this.accOutage.Enabled = false;
                    OleDbConnection con = new OleDbConnection("provider=microsoft.ACE.OLEDB.12.0;Data Source=" + inputFile + ";Extended Properties='Excel 12.0 XML;HDR=YAS;IMEX=1;MAXSCANROWS=0'");
                    OleDbCommand cmd = new OleDbCommand();
                    progressBar1.Value += 1; //////
                    cmd.Connection = con;
                    cmd.CommandText = "SELECT Site, Duration, Category, SubCategory, RootCause, Vendor from [Unplanned Site$]";
                    progressBar1.Value += 2; //////
                    con.Open();
                    dt_Acc_unplannedSite.Load(cmd.ExecuteReader());
                    progressBar1.Value += 3; //////
                    con.Close();
                    OleDbCommand cmd2 = new OleDbCommand();
                    progressBar1.Value += 4; //////
                    cmd2.Connection = con;
                    cmd2.CommandText = "SELECT Site, Duration, Category, SubCategory, RootCause, Vendor from [Planned Site$]";
                    progressBar1.Value += 5; //////
                    con.Open();
                    dt_Acc_plannedSite.Load(cmd2.ExecuteReader());
                    con.Close();
                    progressBar1.Value += 6; //////
                    ////////////
                    MessageBox.Show("completed");
                    progressBar1.Value = 0;
                    dgvAcc.DataSource = dt_Acc_unplannedSite;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                progressBar1.Value = 0;
                this.accOutage.Enabled = true;
            }
        }

        private void rowData_Click(object sender, EventArgs e)
        {
            try
            {

                progressBar1.Style = ProgressBarStyle.Marquee;
                progressBar1.MarqueeAnimationSpeed = 30;

                String inputFile;
                OpenFileDialog op = new OpenFileDialog();
                op.Filter = "All File | *.*|Excel File |*.xlsx";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    inputFile = op.FileName;
                    this.rowData.Enabled = false;
                    OleDbConnection con = new OleDbConnection("provider=microsoft.ACE.OLEDB.12.0;Data Source=" + inputFile + ";Extended Properties='Excel 12.0 XML;HDR=YAS;IMEX=1;MAXSCANROWS=0'");
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "SELECT * from [2G$]";
                    con.Open();
                    dt_2G.Load(cmd.ExecuteReader());
                    con.Close();
                    OleDbCommand cmd2 = new OleDbCommand();
                    cmd2.Connection = con;
                    cmd2.CommandText = "SELECT * from [3G$]";
                    con.Open();
                    dt_3G.Load(cmd2.ExecuteReader());
                    con.Close();

                    dgv_2G_Hua.DataSource = dt_2G;
                    dgv_3G_Hua.DataSource = dt_3G;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                progressBar1.Style = ProgressBarStyle.Continuous;
                progressBar1.MarqueeAnimationSpeed = 0;

                this.rowData.Enabled = true;
            }
        }

        private void TT_Dashboard_Click(object sender, EventArgs e)
        {

            try
            {

                this.TT_Dashboard.Enabled = false;
                String inputFile;
                OpenFileDialog op = new OpenFileDialog();
                op.Filter = "All File | *.*|Excel File |*.xlsx";
                if (op.ShowDialog() == DialogResult.OK)
                {
                    inputFile = op.FileName;
                    OleDbConnection con = new OleDbConnection("provider=microsoft.ACE.OLEDB.12.0;Data Source=" + inputFile + ";Extended Properties='Excel 12.0 XML;HDR=YAS;IMEX=1;MAXSCANROWS=0'");
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = con;
                    cmd.CommandText = "SELECT * from [Chronic Availability$]";
                    con.Open();
                    dt_TT_Dashboard.Load(cmd.ExecuteReader());
                    con.Close();
                }

                progressBar1.Maximum = dt_TT_Dashboard.Rows.Count; // your loop max no. 
                progressBar1.Value = 0;

                for (int i = 0; i < dt_TT_Dashboard.Rows.Count; i++)
                {
                    TT_Dashboard TT = new TT_Dashboard();
                    if (dt_TT_Dashboard.Rows[i][0] != null) { TT.ID = dt_TT_Dashboard.Rows[i][0].ToString(); } else { TT.ID = ""; }
                    if (dt_TT_Dashboard.Rows[i][1] != null) { TT.Number = dt_TT_Dashboard.Rows[i][1].ToString(); } else { TT.Number = ""; }
                    if (dt_TT_Dashboard.Rows[i][3] != null) { TT.Status = dt_TT_Dashboard.Rows[i][3].ToString(); } else { TT.Status = ""; }
                    if (dt_TT_Dashboard.Rows[i][27] != null) { TT.Solution = dt_TT_Dashboard.Rows[i][27].ToString(); } else { TT.Solution = ""; }
                    if (dt_TT_Dashboard.Rows[i][28] != null) { TT.Description = dt_TT_Dashboard.Rows[i][28].ToString(); } else { TT.Description = ""; }

                    TT_DashboardList.Add(TT);

                    progressBar1.Value += 1;

                }

                MessageBox.Show("completed");
                progressBar1.Value = 0;
                label1.Text = "✓✓";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                progressBar1.Value = 0;
                this.TT_Dashboard.Enabled = true;
            }
        }

        private void pross_2G_Hua_but_Click(object sender, EventArgs e)
        {
            try
            {
                this.pross_2G_Hua_but.Enabled = false;

                for (int i = 0; i < dt_2G.Rows.Count; i++)
                {
                    rowData RD = new rowData();
                    if (dt_2G.Rows[i][0] != null) { RD.Site = dt_2G.Rows[i][0].ToString(); } else { RD.Site = ""; }
                    if (dt_2G.Rows[i][1] != null) { RD.cell = dt_2G.Rows[i][1].ToString(); } else { RD.cell = ""; }
                    int N;
                    RD.D1 = Int32.TryParse(dt_2G.Rows[i][2].ToString(), out N) ? N : 0;
                    RD.D2 = Int32.TryParse(dt_2G.Rows[i][3].ToString(), out N) ? N : 0;
                    RD.D3 = Int32.TryParse(dt_2G.Rows[i][4].ToString(), out N) ? N : 0;
                    RD.D4 = Int32.TryParse(dt_2G.Rows[i][5].ToString(), out N) ? N : 0;
                    RD.D5 = Int32.TryParse(dt_2G.Rows[i][6].ToString(), out N) ? N : 0;
                    RD.D6 = Int32.TryParse(dt_2G.Rows[i][7].ToString(), out N) ? N : 0;
                    RD.D7 = Int32.TryParse(dt_2G.Rows[i][8].ToString(), out N) ? N : 0;
                    RD.D8 = Int32.TryParse(dt_2G.Rows[i][9].ToString(), out N) ? N : 0;
                    RD.D9 = Int32.TryParse(dt_2G.Rows[i][10].ToString(), out N) ? N : 0;
                    RD.D10 = Int32.TryParse(dt_2G.Rows[i][11].ToString(), out N) ? N : 0;
                    RD.D11 = Int32.TryParse(dt_2G.Rows[i][12].ToString(), out N) ? N : 0;
                    RD.D12 = Int32.TryParse(dt_2G.Rows[i][13].ToString(), out N) ? N : 0;
                    RD.D13 = Int32.TryParse(dt_2G.Rows[i][14].ToString(), out N) ? N : 0;
                    RD.D14 = Int32.TryParse(dt_2G.Rows[i][15].ToString(), out N) ? N : 0;

                    Li_2G.Add(RD);
                }

                for (int x = 0; x < Li_2G.Count(); x++)
                {
                    int rep = 0;
                    Double Day1 = Li_2G[x].D1;
                    if (Day1 >= 900 && Day1 < 86400)
                    {
                        rep++;
                    }
                    Double Day2 = Li_2G[x].D2;
                    if (Day2 >= 900 && Day2 < 86400)
                    {
                        rep++;
                    }
                    Double Day3 = Li_2G[x].D3;
                    if (Day3 >= 900 && Day3 < 86400)
                    {
                        rep++;
                    }
                    Double Day4 = Li_2G[x].D4;
                    if (Day4 >= 900 && Day4 < 86400)
                    {
                        rep++;
                    }
                    Double Day5 = Li_2G[x].D5;
                    if (Day5 >= 900 && Day5 < 86400)
                    {
                        rep++;
                    }
                    Double Day6 = Li_2G[x].D6;
                    if (Day6 >= 900 && Day6 < 86400)
                    {
                        rep++;
                    }
                    Double Day7 = Li_2G[x].D7;
                    if (Day7 >= 900 && Day7 < 86400)
                    {
                        rep++;
                    }
                    Double Day8 = Li_2G[x].D8;
                    if (Day8 >= 900 && Day8 < 86400)
                    {
                        rep++;
                    }
                    Double Day9 = Li_2G[x].D9;
                    if (Day9 >= 900 && Day9 < 86400)
                    {
                        rep++;
                    }
                    Double Day10 = Li_2G[x].D10;
                    if (Day10 >= 900 && Day10 < 86400)
                    {
                        rep++;
                    }
                    Double Day11 = Li_2G[x].D11;
                    if (Day11 >= 900 && Day11 < 86400)
                    {
                        rep++;
                    }
                    Double Day12 = Li_2G[x].D12;
                    if (Day12 >= 900 && Day12 < 86400)
                    {
                        rep++;
                    }
                    int CLD = 0;
                    Double Day13 = Li_2G[x].D13;
                    if (Day13 <= 450)
                    {
                        CLD++;
                    }
                    if (Day13 >= 900 && Day13 < 86400)
                    {
                        rep++;
                    }
                    Double Day14 = Li_2G[x].D14;
                    if (Day14 <= 450)
                    {
                        CLD++;
                    }
                    if (Day14 >= 900 && Day14 < 86400)
                    {
                        rep++;
                    }

                    Li_2G[x].Repeatition = rep;
                    Li_2G[x].ClearLastDays = CLD;

                    //Li_2G[x].Site = Li_2G[x].Site.Replace('G', 'H');

                    /////////////////////////////////////////////////////////

                    //if (Li_2G[x].Site.Left(1, x) == "P")
                    //{
                    //    Li_2G[x].Vendor = "Ericsson Upper";
                    //}
                    //else if (Li_2G[x].Site.Left(1) == "Z")
                    //{
                    //    Li_2G[x].Vendor = "ZTE";
                    //}
                    //else if (Li_2G[x].Site.Left(1) == "H")
                    //{
                    //    Li_2G[x].Vendor = "Huawei";
                    //}
                    //else
                    //{
                    //    Li_2G[x].Vendor = "Ericsson";
                    //}


                    //if (Li_2G[x].Vendor == "Ericsson" || Li_2G[x].Vendor == "ZTE")
                    //{
                    //    Li_2G[x].Site_Vendor = Extensions.GetNumbers(Li_2G[x].Site) + "_" + Li_2G[x].Vendor;
                    //}
                    //else
                    //{
                    //    if (Li_2G[x].Vendor == "Ericsson Upper")
                    //    {
                    //        Li_2G[x].Site_Vendor = Li_2G[x].Site;
                    //    }
                    //    else
                    //    {
                    //        Li_2G[x].Site_Vendor = Li_2G[x].Vendor;
                    //    }
                    //}

                    //if (Li_2G[x].Vendor == "Ericsson" || Li_2G[x].Vendor == "ZTE" || Li_2G[x].Vendor == "Ericsson Upper")
                    //{
                    //    var sitename = li_siteStatus.FirstOrDefault(o => o.SiteID == Li_2G[x].Site_Vendor);
                    //    if (sitename != null)
                    //    {
                    //        Li_2G[x].Site = sitename.NodeID;
                    //    }

                    //}
                    //////////////////////////////////////////////////////////

                    var item = li_siteStatus.FirstOrDefault(o => o.NodeID == Li_2G[x].Site);
                    if (item != null)
                    {
                        if ((item.Status.ToUpper() == "ON AIR") && (item.subcategory.ToUpper() == "EOT" || item.subcategory.ToUpper() == "FM INSPECTION"))
                        {
                            Li_2G[x].StatusMode = "Done";
                        }
                        else
                        {
                            Li_2G[x].StatusMode = "not Done";
                        }
                    }
                    else
                    {
                        Li_2G[x].StatusMode = "not find";
                    }

                }

                for (int y = 0; y < Li_2G.Count(); y++)
                {
                    if (Li_2G[y].Repeatition >= 5 && Li_2G[y].StatusMode == "Done" && (Li_2G[y].ClearLastDays != 2))
                    {
                        Li_chronicSite_2G.Add(Li_2G[y]);

                        Li_2G[y].chronic = "Yes";

                    }
                }



                DataView dv = Extensions.ConvertToDataTable(Li_chronicSite_2G).DefaultView;
                dv.Sort = "Site desc";
                DataTable sortedDT = dv.ToTable();

                foreach (rowData x in Li_chronicSite_2G)
                {
                    chronicSiteName2G.Add(x.Site);
                }

                List<string> distinctSiteName = chronicSiteName2G.Distinct().ToList();
                for (int i = 0; i < distinctSiteName.Count(); i++)
                {
                    distinctSiteName[i] = distinctSiteName[i].Right(7);
                }

                chronicSite2GHuaLB.DataSource = distinctSiteName;

                labelChronicSite2GHua.Text = distinctSiteName.Count().ToString();
                dgv_2G_Hua.DataSource = sortedDT;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                progressBar1.Value = 0;
                this.pross_2G_Hua_but.Enabled = true;
            }

        }

        private void chronicSite2GHuaLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                SiteName2GHuaTextBox.Text = "";
                SiteOrCell2GHuaTextBox.Text = "";
                Comment2GHuaTextBox.Text = "";
                Description2GHuaTextBox.Text = "";
                OldOrNew2GHuaTextBox.Text = "";
                IM2GHuaTextBox.Text = "";

                bool IsCell = false;
                bool IsCascaded = false;

                string curItem = chronicSite2GHuaLB.SelectedItem.ToString();
                MessageBox.Show(curItem);

                DataTable dt_dv = new DataTable();
                DataView dv = new DataView(dt_Acc_unplannedSite);
                dv.Sort = "Duration desc";
                dv.RowFilter = string.Format("Site LIKE '%{0}%'", curItem);
                dt_dv = dv.ToTable();
                if (dt_dv.Rows.Count > 0)
                {
                    reason2G = dt_dv.Rows[0][2].ToString() + "_" + dt_dv.Rows[0][3].ToString();
                    if ((dt_dv.Rows[0][3].ToString()).Contains("Dependency"))
                    {
                        IsCascaded = true;
                    }

                }
                dgv_2G_Hua_unplanned.DataSource = dt_dv;


                DataTable dt_dv1 = new DataTable();
                DataView dv1 = new DataView(dt_Acc_plannedSite);
                dv1.RowFilter = string.Format("Site LIKE '%{0}%'", curItem);
                dt_dv1 = dv1.ToTable();
                dgv_2G_Hua_planned.DataSource = dt_dv1;

                DataTable dt_dv2 = new DataTable();
                DataView dv2 = new DataView(dt_TT_Dashboard);
                dv2.RowFilter = string.Format("ID LIKE '%{0}%'", curItem.Right(7));
                dt_dv2 = dv2.ToTable();
                dgv_2G_Hua_TT.DataSource = dt_dv2;

                DataTable dt_dv3 = new DataTable();
                DataView dv3 = new DataView(Extensions.ConvertToDataTable(Li_2G));
                dv3.RowFilter = string.Format("Site LIKE '%{0}%'", curItem.Right(7));
                dt_dv3 = dv3.ToTable();
                int Cellnum = 0;
                List<string> Cellname = new List<string>();
                for (int i = 0; i < dt_dv3.Rows.Count; i++)
                {
                    if (dt_dv3.Rows[i][1].ToString() == "Yes")
                    {
                        Cellnum++;
                        Cellname.Add(dt_dv3.Rows[i][2].ToString());
                    }
                }
                if ((dt_dv3.Rows.Count) != Cellnum)
                {
                    IsCell = true;
                }
                dgv_2G_Hua.DataSource = dt_dv3;

                SiteName2GHuaTextBox.Text = curItem;
                if (IsCell)
                {

                    SiteOrCell2GHuaTextBox.Text = "Cell";
                    Comment2GHuaTextBox.Text = " Open a new T.T. on Problem Management (pending other) to follow up the case ";
                    Description2GHuaTextBox.Text = string.Join(" - ", Cellname) + " Were flapping during week " + (Extensions.GetIso8601WeekOfYear(startDate)).ToString() + " " + startDate.Date.ToString("yyyy") + " ( from " + startDate.Date.ToString("dddd, dd MMMM yyyy") + " till " + endDate.Date.ToString("dddd, dd MMMM yyyy") + " ), please support";
                }
                else
                {
                    SiteOrCell2GHuaTextBox.Text = "Site";
                    Comment2GHuaTextBox.Text = "Open a new T.T. on MSU to follow up the case";
                    reason2GTextBox.Text = reason2G;
                    Description2GHuaTextBox.Text = "Site was flapping during week " + (Extensions.GetIso8601WeekOfYear(startDate)).ToString() + " " + startDate.Date.ToString("yyyy") + " ( from " + startDate.Date.ToString("dddd, dd MMMM yyyy") + " till " + endDate.Date.ToString("dddd, dd MMMM yyyy") + " ) due to " + reason2G + " , please support ( and check battery status )";
                }

                if (IsCascaded)
                {
                    Comment2GHuaTextBox.Text = dt_dv.Rows[0][4].ToString();
                    Description2GHuaTextBox.Text = "";
                    OldOrNew2GHuaTextBox.Text = "no";
                    IM2GHuaTextBox.Text = "";
                }
                //foreach (TT_Dashboard isOld in TT_DashboardList)
                //{
                //    if (isOld.ID == curItem)
                //    {
                //        label_ISold.Text = "Old";
                //    }
                //    else
                //    {
                //        label_ISold.Text = "New";
                //    }
                //}

                var isOld = TT_DashboardList.FirstOrDefault(o => o.ID == curItem.Right(7));
                if (isOld != null)
                {
                    label_ISold_2GHua.Text = "Old";
                    OldOrNew2GHuaTextBox.Text = "Old";
                    IM2GHuaTextBox.Text = isOld.Number;
                }
                else
                {
                    label_ISold_2GHua.Text = "New";
                    OldOrNew2GHuaTextBox.Text = "New";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void Done2GHuabut_Click(object sender, EventArgs e)
        {
            try
            {
                Excelsheet item = new Excelsheet();
                item.Site = SiteName2GHuaTextBox.Text;
                item.SiteOrCell = SiteOrCell2GHuaTextBox.Text;
                item.OldOrNew = OldOrNew2GHuaTextBox.Text;
                item.Comment = Comment2GHuaTextBox.Text;
                item.Description = Description2GHuaTextBox.Text;
                item.IM = IM2GHuaTextBox.Text;
                Li_excel_2G.Add(item);
                dgv_Ex_sheet_2G_Hua.DataSource = Extensions.ConvertToDataTable(Li_excel_2G);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void Export_2G_Hua_but_Click(object sender, EventArgs e)
        {
            try
            {
                /////////////////////////////////
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dgv_Ex_sheet_2G_Hua.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dgv_Ex_sheet_2G_Hua.Columns[j].HeaderText;
                }

                StartRow++;

                //Write datagridview content
                for (i = 0; i < dgv_Ex_sheet_2G_Hua.Rows.Count; i++)
                {
                    for (j = 0; j < dgv_Ex_sheet_2G_Hua.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dgv_Ex_sheet_2G_Hua[j, i].Value == null ? "" : dgv_Ex_sheet_2G_Hua[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }

                var saveFileDialoge = new SaveFileDialog();
                saveFileDialoge.FileName = "OUTPUT 2G";
                saveFileDialoge.DefaultExt = ".xlsx";
                if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void pross_3G_Hua_but_Click(object sender, EventArgs e)
        {
            try
            {
                this.pross_3G_Hua_but.Enabled = false;

                for (int i = 0; i < dt_3G.Rows.Count; i++)
                {
                    rowData RD = new rowData();
                    if (dt_3G.Rows[i][0] != null) { RD.Site = dt_3G.Rows[i][0].ToString(); } else { RD.Site = ""; }
                    if (dt_3G.Rows[i][1] != null) { RD.cell = dt_3G.Rows[i][1].ToString(); } else { RD.cell = ""; }
                    int N;
                    RD.D1 = Int32.TryParse(dt_3G.Rows[i][2].ToString(), out N) ? N : 0;
                    RD.D2 = Int32.TryParse(dt_3G.Rows[i][3].ToString(), out N) ? N : 0;
                    RD.D3 = Int32.TryParse(dt_3G.Rows[i][4].ToString(), out N) ? N : 0;
                    RD.D4 = Int32.TryParse(dt_3G.Rows[i][5].ToString(), out N) ? N : 0;
                    RD.D5 = Int32.TryParse(dt_3G.Rows[i][6].ToString(), out N) ? N : 0;
                    RD.D6 = Int32.TryParse(dt_3G.Rows[i][7].ToString(), out N) ? N : 0;
                    RD.D7 = Int32.TryParse(dt_3G.Rows[i][8].ToString(), out N) ? N : 0;
                    RD.D8 = Int32.TryParse(dt_3G.Rows[i][9].ToString(), out N) ? N : 0;
                    RD.D9 = Int32.TryParse(dt_3G.Rows[i][10].ToString(), out N) ? N : 0;
                    RD.D10 = Int32.TryParse(dt_3G.Rows[i][11].ToString(), out N) ? N : 0;
                    RD.D11 = Int32.TryParse(dt_3G.Rows[i][12].ToString(), out N) ? N : 0;
                    RD.D12 = Int32.TryParse(dt_3G.Rows[i][13].ToString(), out N) ? N : 0;
                    RD.D13 = Int32.TryParse(dt_3G.Rows[i][14].ToString(), out N) ? N : 0;
                    RD.D14 = Int32.TryParse(dt_3G.Rows[i][15].ToString(), out N) ? N : 0;

                    Li_3G.Add(RD);
                }

                for (int x = 0; x < Li_3G.Count(); x++)
                {
                    int rep = 0;
                    Double Day1 = Li_3G[x].D1;
                    if (Day1 >= 900 && Day1 < 86400)
                    {
                        rep++;
                    }
                    Double Day2 = Li_3G[x].D2;
                    if (Day2 >= 900 && Day2 < 86400)
                    {
                        rep++;
                    }
                    Double Day3 = Li_3G[x].D3;
                    if (Day3 >= 900 && Day3 < 86400)
                    {
                        rep++;
                    }
                    Double Day4 = Li_3G[x].D4;
                    if (Day4 >= 900 && Day4 < 86400)
                    {
                        rep++;
                    }
                    Double Day5 = Li_3G[x].D5;
                    if (Day5 >= 900 && Day5 < 86400)
                    {
                        rep++;
                    }
                    Double Day6 = Li_3G[x].D6;
                    if (Day6 >= 900 && Day6 < 86400)
                    {
                        rep++;
                    }
                    Double Day7 = Li_3G[x].D7;
                    if (Day7 >= 900 && Day7 < 86400)
                    {
                        rep++;
                    }
                    Double Day8 = Li_3G[x].D8;
                    if (Day8 >= 900 && Day8 < 86400)
                    {
                        rep++;
                    }
                    Double Day9 = Li_3G[x].D9;
                    if (Day9 >= 900 && Day9 < 86400)
                    {
                        rep++;
                    }
                    Double Day10 = Li_3G[x].D10;
                    if (Day10 >= 900 && Day10 < 86400)
                    {
                        rep++;
                    }
                    Double Day11 = Li_3G[x].D11;
                    if (Day11 >= 900 && Day11 < 86400)
                    {
                        rep++;
                    }
                    Double Day12 = Li_3G[x].D12;
                    if (Day12 >= 900 && Day12 < 86400)
                    {
                        rep++;
                    }
                    int CLD = 0;
                    Double Day13 = Li_3G[x].D13;
                    if (Day13 <= 450)
                    {
                        CLD++;
                    }
                    if (Day13 >= 900 && Day13 < 86400)
                    {
                        rep++;
                    }
                    Double Day14 = Li_3G[x].D14;
                    if (Day14 <= 450)
                    {
                        CLD++;
                    }
                    if (Day14 >= 900 && Day14 < 86400)
                    {
                        rep++;
                    }



                    Li_3G[x].Repeatition = rep;

                    Li_3G[x].ClearLastDays = CLD;


                    //Li_3G[x].Site = Li_3G[x].Site.Replace("Z", string.Empty);

                    var item = li_siteStatus.FirstOrDefault(o => o.NodeID == Li_3G[x].Site);
                    if (item != null)
                    {
                        if ((item.Status.ToUpper() == "ON AIR") && (item.subcategory.ToUpper() == "EOT" || item.subcategory.ToUpper() == "FM INSPECTION"))
                        {
                            Li_3G[x].StatusMode = "Done";
                        }
                        else
                        {
                            Li_3G[x].StatusMode = "not Done";
                        }
                    }
                    else
                    {
                        Li_3G[x].StatusMode = "not find";
                    }

                }

                for (int y = 0; y < Li_3G.Count(); y++)
                {
                    if ((Li_3G[y].Repeatition >= 5) && (Li_3G[y].StatusMode == "Done") && (Li_3G[y].ClearLastDays != 2))
                    {

                        Li_chronicSite_3G.Add(Li_3G[y]);

                        Li_3G[y].chronic = "Yes";


                    }
                }



                DataView dv = Extensions.ConvertToDataTable(Li_chronicSite_3G).DefaultView;
                dv.Sort = "Site desc";
                DataTable sortedDT = dv.ToTable();

                List<string> chronicSiteName = new List<string>();
                foreach (rowData x in Li_chronicSite_3G)
                {
                    chronicSiteName.Add(x.Site);
                }

                List<string> distinctSiteName = chronicSiteName.Distinct().ToList();
                for (int i = 0; i < distinctSiteName.Count(); i++)
                {
                    distinctSiteName[i] = distinctSiteName[i].Right(7);
                }

                chronicSite3GHuaLB.DataSource = distinctSiteName;

                labelChronicSite3GHua.Text = distinctSiteName.Count().ToString();
                dgv_3G_Hua.DataSource = sortedDT;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                progressBar1.Value = 0;
                this.pross_3G_Hua_but.Enabled = true;
            }

        }

        private void chronicSite3GHuaLB_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {

                SiteName3GHuaTextBox.Text = "";
                SiteOrCell3GHuaTextBox.Text = "";
                Comment3GHuaTextBox.Text = "";
                Description3GHuaTextBox.Text = "";
                OldOrNew3GHuaTextBox.Text = "";
                IM3GHuaTextBox.Text = "";

                bool IsCascaded = false;
                bool IsCell = false;
                string curItem = chronicSite3GHuaLB.SelectedItem.ToString();
                MessageBox.Show(curItem);

                DataTable dt_dv = new DataTable();
                DataView dv = new DataView(dt_Acc_unplannedSite);
                dv.RowFilter = string.Format("Site LIKE '%{0}%'", curItem);
                dv.Sort = "Duration desc";
                dt_dv = dv.ToTable();
                if (dt_dv.Rows.Count > 0)
                {
                    reason3G = dt_dv.Rows[0][2].ToString() + "_" + dt_dv.Rows[0][3].ToString();
                    if ((dt_dv.Rows[0][3].ToString()).Contains("Dependency"))
                    {
                        IsCascaded = true;
                    }
                }
                dgv_3G_Hua_unplanned.DataSource = dt_dv;



                DataView dv1 = new DataView(dt_Acc_plannedSite);
                dv1.RowFilter = string.Format("Site LIKE '%{0}%'", curItem);
                dgv_3G_Hua_planned.DataSource = dv1;


                DataView dv2 = new DataView(dt_TT_Dashboard);
                dv2.RowFilter = string.Format("ID LIKE '%{0}%'", curItem.Right(7));
                dgv_3G_Hua_TT.DataSource = dv2;


                DataTable dt_dv3 = new DataTable();
                DataView dv3 = new DataView(Extensions.ConvertToDataTable(Li_3G));
                dv3.RowFilter = string.Format("Site LIKE '%{0}%'", curItem.Right(7));
                dt_dv3 = dv3.ToTable();
                int Cellnum = 0;
                List<string> Cellname = new List<string>();
                for (int i = 0; i < dt_dv3.Rows.Count; i++)
                {
                    if (dt_dv3.Rows[i][1].ToString() == "Yes")
                    {
                        Cellnum++;
                        Cellname.Add(dt_dv3.Rows[i][2].ToString());
                    }
                }
                if ((dt_dv3.Rows.Count) != Cellnum)
                {
                    IsCell = true;
                }
                dgv_3G_Hua.DataSource = dt_dv3;

                SiteName3GHuaTextBox.Text = "U"+curItem;
                if (IsCell)
                {

                    SiteOrCell3GHuaTextBox.Text = "Cell";
                    Comment3GHuaTextBox.Text = " Open a new T.T. on Problem Management (pending other) to follow up the case ";
                    Description3GHuaTextBox.Text = string.Join(" - ", Cellname) + " Were flapping during week " + (Extensions.GetIso8601WeekOfYear(startDate)).ToString() + " " + startDate.Date.ToString("yyyy") + " ( from " + startDate.Date.ToString("dddd, dd MMMM yyyy") + " till " + endDate.Date.ToString("dddd, dd MMMM yyyy") + " ), please support";
                }
                else
                {
                    SiteOrCell3GHuaTextBox.Text = "Site";
                    Comment3GHuaTextBox.Text = "Open a new T.T. on MSU to follow up the case";
                    //reason3GTextBox.Text = reason3G;
                    Description3GHuaTextBox.Text = "Site was flapping during week " + (Extensions.GetIso8601WeekOfYear(startDate)).ToString() + " " + startDate.Date.ToString("yyyy") + " ( from " + startDate.Date.ToString("dddd, dd MMMM yyyy") + " till " + endDate.Date.ToString("dddd, dd MMMM yyyy") + " ) due to " + reason3G + " , please support ( and check battery status )";
                }

                if (IsCascaded)
                {
                    Comment3GHuaTextBox.Text = dt_dv.Rows[0][4].ToString();
                    Description3GHuaTextBox.Text = "";
                    OldOrNew3GHuaTextBox.Text = "no";
                    IM3GHuaTextBox.Text = "";
                }

                var isOld = TT_DashboardList.FirstOrDefault(o => o.ID == curItem.Right(7));
                if (isOld != null)
                {
                    label_ISold_3GHua.Text = "Old";
                    OldOrNew3GHuaTextBox.Text = "Old";
                    IM3GHuaTextBox.Text = isOld.Number;

                }
                else
                {
                    label_ISold_3GHua.Text = "New";
                    OldOrNew3GHuaTextBox.Text = "New";
                }
                if (chronicSiteName2G.Count() >= 1)
                {
                    foreach (string v in chronicSiteName2G)
                    {
                        string g = v.Right(7);
                        if (g == curItem.Right(7))
                        {
                            label_Detected_3GHua.Text = "Detected in 2G";
                            Comment3GHuaTextBox.Text = "Detected in 2G";
                            Description3GHuaTextBox.Text = "";
                            OldOrNew3GHuaTextBox.Text = "no";
                            IM3GHuaTextBox.Text = "";
                            break;
                        }
                        else
                        {
                            label_Detected_3GHua.Text = "";
                        }
                    }
                }
                else
                {
                    label_Detected_3GHua.Text = "Can't Access 2G";
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Export_3G_Hua_but_Click(object sender, EventArgs e)
        {
            try
            {
                /////////////////////////////////
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                int StartCol = 1;
                int StartRow = 1;
                int j = 0, i = 0;

                //Write Headers
                for (j = 0; j < dgv_Ex_sheet_3G_Hua.Columns.Count; j++)
                {
                    Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow, StartCol + j];
                    myRange.Value2 = dgv_Ex_sheet_3G_Hua.Columns[j].HeaderText;
                }

                StartRow++;

                //Write datagridview content
                for (i = 0; i < dgv_Ex_sheet_3G_Hua.Rows.Count; i++)
                {
                    for (j = 0; j < dgv_Ex_sheet_3G_Hua.Columns.Count; j++)
                    {
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dgv_Ex_sheet_3G_Hua[j, i].Value == null ? "" : dgv_Ex_sheet_3G_Hua[j, i].Value;
                        }
                        catch
                        {
                            ;
                        }
                    }
                }

                var saveFileDialoge = new SaveFileDialog();
                saveFileDialoge.FileName = "OUTPUT 3G";
                saveFileDialoge.DefaultExt = ".xlsx";
                if (saveFileDialoge.ShowDialog() == DialogResult.OK)
                {
                    workbook.SaveAs(saveFileDialoge.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Done3GHuabut_Click(object sender, EventArgs e)
        {
            try
            {
                Excelsheet item = new Excelsheet();
                item.Site = SiteName3GHuaTextBox.Text;
                item.SiteOrCell = SiteOrCell3GHuaTextBox.Text;
                item.OldOrNew = OldOrNew3GHuaTextBox.Text;
                item.Comment = Comment3GHuaTextBox.Text;
                item.Description = Description3GHuaTextBox.Text;
                item.IM = IM3GHuaTextBox.Text;
                Li_excel_3G.Add(item);
                dgv_Ex_sheet_3G_Hua.DataSource = Extensions.ConvertToDataTable(Li_excel_3G);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void reason2GTextBox_TextChanged(object sender, EventArgs e)
        {
            reason2G = reason2GTextBox.Text;
            Description2GHuaTextBox.Text = "Site was flapping during week " + (Extensions.GetIso8601WeekOfYear(startDate)).ToString() + " " + startDate.Date.ToString("yyyy") + " ( from " + startDate.Date.ToString("dddd, dd MMMM yyyy") + " till " + endDate.Date.ToString("dddd, dd MMMM yyyy") + " ) due to " + reason2G + " , please support ( and check battery status )";
        }

        private void reason3GTextBox_TextChanged(object sender, EventArgs e)
        {
            reason3G = reason3GTextBox.Text; 
            Description3GHuaTextBox.Text = "Site was flapping during week " + (Extensions.GetIso8601WeekOfYear(startDate)).ToString() + " " + startDate.Date.ToString("yyyy") + " ( from " + startDate.Date.ToString("dddd, dd MMMM yyyy") + " till " + endDate.Date.ToString("dddd, dd MMMM yyyy") + " ) due to " + reason3G + " , please support ( and check battery status )";
        }
    }
}