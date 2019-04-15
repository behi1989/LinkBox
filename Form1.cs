using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using TMS.Class;

namespace NSP_LinkBox
{
    public partial class frmMain : Form
    {
        DataGridViewPrinter dgvPrint;

        System.Globalization.PersianCalendar pc = new System.Globalization.PersianCalendar();
        System.Globalization.HijriCalendar hc = new System.Globalization.HijriCalendar();

        //******************************************

        MaftooxCalendar.MaftooxPersianCalendar.TimeWork prdTime = new MaftooxCalendar.MaftooxPersianCalendar.TimeWork();
        MaftooxCalendar.MaftooxPersianCalendar.DateWork prd = new MaftooxCalendar.MaftooxPersianCalendar.DateWork();

        //*******************************************

        public void showMaindgv()
        {
            
            string sqlstring;
            sqlstring = "SELECT * FROM tbllinkbox ORDER BY url_visited DESC";
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = database.select(sqlstring);
            dgvMainShow.DataSource = dt;
            dgvMainShow.Columns["radif"].Width = 50;
            dgvMainShow.Columns["url_description"].HeaderText = "توضیحات لینک";
            dgvMainShow.Columns["url_description"].Width = 220;
            dgvMainShow.Columns["url_location"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvMainShow.Columns["url_location"].HeaderText = "آدرس لینک";
            dgvMainShow.Columns["id"].Visible = false;
            dgvMainShow.Columns["url_group"].Visible = false;
            dgvMainShow.Columns["url_tags"].Visible = false;
            dgvMainShow.Columns["url_priority"].Visible = false;
            dgvMainShow.Columns["url_rate"].Visible = false;
            dgvMainShow.Columns["url_visited"].Visible = false;
            
        }

        public void showManagedgv()
        {
            string sqlstring;
            sqlstring = "SELECT * FROM tbllinkbox";
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = database.select(sqlstring);
            dgvManage.DataSource = dt;
            dgvManage.Columns["radifs"].Width = 50;
            dgvManage.Columns["url_description"].HeaderText = "توضیحات لینک";
            dgvManage.Columns["url_description"].Width = 210;
            dgvManage.Columns["url_location"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvManage.Columns["url_location"].HeaderText = "آدرس لینک";
            dgvManage.Columns["id"].Visible = false;
            dgvManage.Columns["url_group"].Visible = false;
            dgvManage.Columns["url_tags"].Visible = false;
            dgvManage.Columns["url_priority"].Visible = false;
            dgvManage.Columns["url_rate"].Visible = false;
            dgvManage.Columns["url_visited"].Visible = false;
        }
        public void pass_dgv()
        {
            string sqlstring;
            sqlstring = "SELECT * FROM tblpass";
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = database.select(sqlstring);
            dataGridView1.DataSource = dt;
        }

        public void TopShowLink()
        {
            string sqlstring;
            sqlstring = "SELECT * FROM tbllinkbox ORDER BY url_visited DESC";
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = database.select(sqlstring);
            dgvTopShowLink.DataSource = dt;
            dgvTopShowLink.Columns["rdf"].Width = 50;
            dgvTopShowLink.Columns["url_description"].HeaderText = "توضیحات لینک";
            dgvTopShowLink.Columns["url_description"].Width = 220;
            dgvTopShowLink.Columns["url_location"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dgvTopShowLink.Columns["url_location"].HeaderText = "آدرس لینک";
            dgvTopShowLink.Columns["id"].Visible = false;
            dgvTopShowLink.Columns["url_group"].Visible = false;
            dgvTopShowLink.Columns["url_tags"].Visible = false;
            dgvTopShowLink.Columns["url_priority"].Visible = false;
            dgvTopShowLink.Columns["url_rate"].Visible = false;
            dgvTopShowLink.Columns["url_visited"].Visible = false;
        }
        
        public void dgvShowPrints()
        {
            string sqlstring;
            sqlstring = "SELECT * FROM tbllinkbox";
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = database.select(sqlstring);
            dgvShowPrint.DataSource = dt;
            dgvShowPrint.Columns["raddif"].Width = 50;
            dgvShowPrint.Columns["url_description"].HeaderText = "توضیحات لینک";
            dgvShowPrint.Columns["url_description"].Width = 220;
            dgvShowPrint.Columns["url_location"].Width = 200;
            dgvShowPrint.Columns["url_location"].HeaderText = "آدرس لینک";
            dgvShowPrint.Columns["url_group"].Width = 150;
            dgvShowPrint.Columns["url_group"].HeaderText = "گروه";
            dgvShowPrint.Columns["id"].Visible = false;
            //dgvShowPrint.Columns["url_group"].Visible = false;
            dgvShowPrint.Columns["url_tags"].Visible = false;
            dgvShowPrint.Columns["url_priority"].Visible = false;
            dgvShowPrint.Columns["url_rate"].Visible = false;
            dgvShowPrint.Columns["url_visited"].Visible = false;
        }

        public void ResetItems()
        {
            txtId.Clear();
            txtDescription.Clear();
            txtDescription.Focus();
            txtLinkAddress.Text="http://www.";
            txtLinkKeywords.Clear();
            cmbLinkgroup.SelectedItem = null;
            cmbLinkpriority.SelectedItem = null;
            cmbLinkrate.SelectedItem = null;
        }
        private bool SetupPrinting()
        {
            PrintDialog pd = new PrintDialog();
            pd.AllowCurrentPage = false;
            pd.AllowSomePages = false;
            pd.AllowPrintToFile = false;
            pd.PrintToFile = false;
            pd.ShowHelp = false;
            pd.ShowNetwork = false;
            if (pd.ShowDialog() != DialogResult.OK)
                return false;
            printDoc.DocumentName = "لیست تمامی لینک ها";
            printDoc.PrinterSettings = pd.PrinterSettings;
            printDoc.DefaultPageSettings = pd.PrinterSettings.DefaultPageSettings;
            printDoc.DefaultPageSettings.Margins = new System.Drawing.Printing.Margins(40, 40, 40, 40);

            dgvPrint = new DataGridViewPrinter(dgvShowPrint, printDoc, true, true, "لیست تمامی لینک های ثبت شده", new System.Drawing.Font("B Yekan", 12, FontStyle.Bold, GraphicsUnit.Point), Color.Black, true);

            return true;
        }


        public frmMain()
        {
            InitializeComponent();
        }

        private void picShowMainPanel1_MouseHover(object sender, EventArgs e)
        {
            picShowMainPanel1.Visible = false;
            picShowMainPanel2.Visible = true;
        }

        private void picShowMainPanel2_MouseLeave(object sender, EventArgs e)
        {
            picShowMainPanel1.Visible = true;
            picShowMainPanel2.Visible = false;
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            try
            {
                switch (DateTime.Now.DayOfWeek)
                {
                    case DayOfWeek.Friday:
                        lblDateNow.Text = "جمعه";
                        break;
                    case DayOfWeek.Monday:
                        lblDateNow.Text = "دوشنبه";
                        break;
                    case DayOfWeek.Saturday:
                        lblDateNow.Text = "شنبه";
                        break;
                    case DayOfWeek.Sunday:
                        lblDateNow.Text = "یکشنبه";
                        break;
                    case DayOfWeek.Thursday:
                        lblDateNow.Text = "پنجشنبه";
                        break;
                    case DayOfWeek.Tuesday:
                        lblDateNow.Text = "سه شنبه";
                        break;
                    case DayOfWeek.Wednesday:
                        lblDateNow.Text = "چهارشنبه";
                        break;
                }
                timer1.Enabled = true;

                System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("fa-ir");
                InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);

                showMaindgv();

                lblCounts.Text = dgvMainShow.RowCount.ToString();
                notifyIcon.ShowBalloonTip(8000,"لینک باکس", "به برنامه لینک باکس خوش آمدید", ToolTipIcon.Info);

                pass_dgv();
                if (dataGridView1.RowCount >= 1)
                {
                    pnlLogin.Visible = true;
                    txtLogin.Focus();
                    lblLoginMessage.Text = "شما هیچ پیغامی ندارید.";
                }    
                else
                {
                    pnlLogin.Visible = false;
                    txtLogin.Focus();
                } 
            }
            catch { }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                prdTime.Upate();
                String stry = prd.GetNameMonth() + prd.GetNameDayInMonth();
                lblDateNow.Text = "  " + prd.GetNameDayInMonth() + "   " + prd.GetNumberDayInMonth().ToString() + "   " + prd.GetNameMonth() + "  سال  " + prd.GetNumberYear().ToString();
                lblTimeNow.Text = prdTime.GetNumberHour().ToString("0#") + ":" + prdTime.GetNumberMinute().ToString("0#") + ":" + prdTime.GetNumberSecond().ToString("0#");
            }
            catch { }
        }

        private void dgvMainShow_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvMainShow.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (txtSearch.Text != "")
                {
                    string SmartSearch;
                    SmartSearch = "SELECT * FROM tbllinkbox WHERE url_description LIKE N'%" + txtSearch.Text + "%'";
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt = database.select(SmartSearch);
                    dgvMainShow.DataSource = dt;
                    dgvMainShow.Columns["radif"].Width = 50;
                    dgvMainShow.Columns["url_description"].HeaderText = "توضیحات لینک";
                    dgvMainShow.Columns["url_description"].Width = 220;
                    dgvMainShow.Columns["url_location"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    dgvMainShow.Columns["url_location"].HeaderText = "آدرس لینک";
                    dgvMainShow.Columns["id"].Visible = false;
                    dgvMainShow.Columns["url_group"].Visible = false;
                    dgvMainShow.Columns["url_tags"].Visible = false;
                    dgvMainShow.Columns["url_priority"].Visible = false;
                    dgvMainShow.Columns["url_rate"].Visible = false;
                    dgvMainShow.Columns["url_visited"].Visible = false;
                    if (dgvMainShow.RowCount > 0)
                    {
                        lblSearchMessages.ForeColor = Color.YellowGreen;
                        lblSearchMessages.Text = "تعداد پیدا شده : "+dgvMainShow.RowCount;
                    }
                    else
                    {
                        lblSearchMessages.Text = "جستجو نتیجه ای نداشت!";
                        lblSearchMessages.ForeColor = Color.Coral;
                        
                    }
                }
                else
                {
                    lblSearchMessages.Text = "برای جستجو کلمه یا حرفی را وارد نمایید";
                    lblSearchMessages.ForeColor = Color.Coral;
                    showMaindgv();
                }
            }
            catch { }
        }

        private void txtSearch_Enter(object sender, EventArgs e)
        {
            txtSearch.Clear();
            txtSearch.ForeColor = Color.Black;
        }

        private void picShowMainPanel2_Click(object sender, EventArgs e)
        {
            pnlmenu.Visible = true;
            dgvShowPrints();
            lblCounts.Text = dgvShowPrint.RowCount.ToString();
        }

        private void pichidemenu1_MouseHover(object sender, EventArgs e)
        {
            pichidemenu1.Visible = false;
            pichidemenu2.Visible = true;
        }

        private void pichidemenu2_MouseLeave(object sender, EventArgs e)
        {
            pichidemenu2.Visible = false;
            pichidemenu1.Visible = true;
        }

        private void pichidemenu2_Click(object sender, EventArgs e)
        {
            pnlmenu.Visible = false;
            txtSearch.Text = "";
            showMaindgv();
        }

        private void pnlmenu_VisibleChanged(object sender, EventArgs e)
        {
            if (pnlmenu.Visible == true)
            {
                picShowMainPanel1.Visible = false;
                picShowMainPanel2.Visible = false;
                pichidemenu1.Visible = true;
            }
            else
            {
                pichidemenu1.Visible = false;
                pichidemenu2.Visible = false;
                picShowMainPanel1.Visible = true;
            }
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void picshowprint_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.None;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlShowPrint.Visible = true;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlSearch.Visible = false;
            pnlTopShow.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            dgvShowPrints();
            lblCounts.Text = dgvShowPrint.RowCount.ToString();
        }

        private void picshowprint_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("نمایش و پرینت لینک ها", picshowprint);
        }

        private void picInsertlink_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("درج لینک جدید", picInsertlink);
        }

        private void picManagelink_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("مدیریت لینک ها", picManagelink);
        }

        private void picAdvanceSearch_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("جستجوی پیشرفته", picAdvanceSearch);
        }

        private void picTopShowLink_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("نمایش لینک های پربازدید", picTopShowLink);
        }

        private void picExport_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("پشتیبان گیری و بازگردانی اطلاعات", picExport);
        }

        private void picSetting_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("تنضیمات برنامه", picSetting);
        }

        private void picAboutUs_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("درباره برنامه", picAboutUs);
        }

        private void picExit_MouseHover(object sender, EventArgs e)
        {
            toolTips.Show("خروج موقت", picExit);
        }

        private void picInsertlink_Click(object sender, EventArgs e)
        {
            lblMessage2.Text = "توجه ! شما هیچ پیغامی ندارید.";
            lblMessage2.ForeColor = Color.LightCoral;
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.None;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlinsert.Visible = true;
            pnlShowPrint.Visible = false;
            pnlManage.Visible = false;
            pnlSearch.Visible = false;
            pnlTopShow.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            ResetItems();
        }

        private void picManagelink_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.None;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlSearch.Visible = false;
            pnlExtract.Visible = false;
            pnlTopShow.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            pnlManage.Visible = true;
            showManagedgv();
        }

        private void picAdvanceSearch_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.None;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlSearch.Visible = true;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlTopShow.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            txtSearchs.Text = "";
            cmbSearchs.SelectedItem = null;
            lblSearch.Text = "توجه ! شما هیچ پیغامی ندارید.";
            lblSearch.ForeColor = Color.LightCoral;
            chbDescription.CheckState = CheckState.Checked;
            chbLocation.CheckState = CheckState.Unchecked;
            chbGroup.CheckState = CheckState.Unchecked;
            chbTags.CheckState = CheckState.Unchecked;
            txtSearchs.Focus();
            cmbSearchs.Enabled = false;
        }

        private void picTopShowLink_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.None;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlTopShow.Visible = true;
            pnlSearch.Visible = false;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            TopShowLink();
        }

        private void picExport_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.None;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlExtract.Visible = true;
            pnlTopShow.Visible = false;
            pnlSearch.Visible = false;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            txtExtractAddress.Clear();
            txtExtractNote.Clear();
            txtBackup.Clear();
            txtRestore.Clear();
        }

        private void picSetting_Click(object sender, EventArgs e)
        {
            pass_dgv();
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.None;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlAboutUs.Visible = false;
            pnlTopShow.Visible = false;
            pnlSearch.Visible = false;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlExtract.Visible = false;
            pnlSettings.Visible = true;
            txtPassword.Focus();
        }

        private void picAboutUs_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.None;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlAboutUs.Visible = true;
            pnlTopShow.Visible = false;
            pnlSearch.Visible = false;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlExtract.Visible = false;
            pnlSettings.Visible = false;
        }

        private void picExit_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.None;
            System.Windows.Forms.Application.Exit();
        }

        private void picPrints2_MouseLeave(object sender, EventArgs e)
        {
            picPrints1.Visible = true;
            picPrints2.Visible = false;
        }

        private void picPrints1_MouseHover(object sender, EventArgs e)
        {
            picPrints1.Visible = false;
            picPrints2.Visible = true;
        }

        private void dgvShowPrint_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvShowPrint.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
        }

        private void picSabt1_MouseHover(object sender, EventArgs e)
        {
            picSabt1.Visible = false;
            picSabt2.Visible = true;
        }

        private void picSabt2_MouseLeave(object sender, EventArgs e)
        {
            picSabt2.Visible = false;
            picSabt1.Visible = true;
        }

        private void picHazf1_MouseHover(object sender, EventArgs e)
        {
            picHazf1.Visible = false;
            picHazf2.Visible = true;
        }

        private void picHazf2_MouseLeave(object sender, EventArgs e)
        {
            picHazf2.Visible = false;
            picHazf1.Visible = true;
        }

        private void picVirayesh1_MouseHover(object sender, EventArgs e)
        {
            picVirayesh1.Visible = false;
            picVirayesh2.Visible = true;
        }

        private void picVirayesh2_MouseLeave(object sender, EventArgs e)
        {
            picVirayesh2.Visible = false;
            picVirayesh1.Visible = true;
        }

        private void picSabt2_Click(object sender, EventArgs e)
        {
            if(txtDescription.Text=="")
            {
                lblMessage2.Text = "لطفا توضیحات لینک را وارد نمایید.";
                lblMessage2.ForeColor = Color.LightCoral;
                txtDescription.Focus();
            }
            else if(txtLinkAddress.Text=="http://www.")
            {
                lblMessage2.Text = "لطفا آدرس لینک را به صورت کامل وارد نمایید.";
                lblMessage2.ForeColor = Color.LightCoral;
                txtLinkAddress.Focus();
            }
            else if (cmbLinkgroup.SelectedItem == null)
            {
                lblMessage2.Text = "لطفا گروه لینک را از منوی کشویی انتخاب نمایید.";
                lblMessage2.ForeColor = Color.LightCoral;
            }
            else
            {
                int visit = 1;
                //visit = int.Parse(txtVisit.Text) + 1;
                string sqlInsert = "INSERT INTO tbllinkbox (url_description,url_location,url_group,url_tags,url_priority,url_rate,url_visited) VALUES " + "(N'" + txtDescription.Text + "',N'" + txtLinkAddress.Text + "',N'" + cmbLinkgroup.SelectedItem + "',N'" + txtLinkKeywords.Text + "',N'" + cmbLinkpriority.SelectedItem + "',N'" + cmbLinkrate.SelectedItem + "',"+visit+")";
                database.DoIMD(sqlInsert);
                lblMessage2.Text = "درج اطلاعات با موفقیت انجام شد.";
                lblMessage2.ForeColor = Color.YellowGreen;
                ResetItems();
            }
        }

        private void dgvManage_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvManage.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
        }

        private void tsmIns_Click(object sender, EventArgs e)
        {
            lblMessage2.Text = "توجه ! شما هیچ پیغامی ندارید.";
            lblMessage2.ForeColor = Color.LightCoral;
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.None;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlinsert.Visible = true;
            pnlShowPrint.Visible = false;
            pnlManage.Visible = false;
            pnlSearch.Visible = false;
            pnlTopShow.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            picSabt1.Visible = true;
            picVirayesh1.Visible = false;
            picHazf1.Visible = false;
            ResetItems();
        }

        private void tsmEdit_Click(object sender, EventArgs e)
        {
            lblMessage2.Text = "توجه ! شما هیچ پیغامی ندارید.";
            lblMessage2.ForeColor = Color.LightCoral;
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.None;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlinsert.Visible = true;
            pnlShowPrint.Visible = false;
            pnlManage.Visible = false;
            pnlSearch.Visible = false;
            pnlTopShow.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            picSabt1.Visible = false;
            picVirayesh1.Visible = true;
            picHazf1.Visible = true;

            txtLinkAddress.Text="";
            txtId.Text = dgvManage.SelectedRows[0].Cells["id"].Value.ToString();
            txtDescription.Text = dgvManage.SelectedRows[0].Cells["url_description"].Value.ToString();
            txtLinkAddress.Text = dgvManage.SelectedRows[0].Cells["url_location"].Value.ToString();
            cmbLinkgroup.Text = dgvManage.SelectedRows[0].Cells["url_group"].Value.ToString();
            txtLinkKeywords.Text = dgvManage.SelectedRows[0].Cells["url_tags"].Value.ToString();
            cmbLinkpriority.Text = dgvManage.SelectedRows[0].Cells["url_priority"].Value.ToString();
            cmbLinkrate.Text = dgvManage.SelectedRows[0].Cells["url_rate"].Value.ToString();
        }

        private void tsmDel_Click(object sender, EventArgs e)
        {
            DialogResult dr = new DialogResult();
            dr = MessageBox.Show("حذف لینک !" + "\n\n" + "آیا از حذف لینک انتخاب شده  مطمئن هستید؟", "اخطار", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            if (dr == DialogResult.Yes)
            {
                string sqlDel = "DELETE FROM tbllinkbox WHERE id=" + dgvManage.SelectedRows[0].Cells["id"].Value.ToString();
                database.DoIMD(sqlDel);
                MessageBox.Show("لینک انتخاب شده با موفقیت حذف شد", "تایید حذف لینک", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                showManagedgv();
            }
        }

        private void tsmSearchs_Click(object sender, EventArgs e)
        {
            lblSearch.Text = "توجه ! شما هیچ پیغامی ندارید.";
            lblMessage2.ForeColor = Color.LightCoral;
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.None;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlinsert.Visible = false;
            pnlShowPrint.Visible = false;
            pnlManage.Visible = false;
            pnlTopShow.Visible = false;
            pnlExtract.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
            pnlSearch.Visible = true;
            txtSearchs.Focus();
        }

        private void picVirayesh2_Click(object sender, EventArgs e)
        {
            if (txtDescription.Text != "" && txtLinkAddress.Text != "" && cmbLinkgroup.SelectedItem != null && txtLinkKeywords.Text != "" && cmbLinkpriority.SelectedItem != null && cmbLinkrate.SelectedItem != null)
            {
                string sqlEdit = "UPDATE tbllinkbox SET url_description=N'" + txtDescription.Text + "',url_location=N'" + txtLinkAddress.Text + "',url_group=N'" + cmbLinkgroup.SelectedItem + "',url_tags=N'" + txtLinkKeywords.Text + "',url_priority=N'" + cmbLinkpriority.SelectedItem + "',url_rate=N'" + cmbLinkrate.SelectedItem + "' WHERE id=" + txtId.Text + " ";
                database.DoIMD(sqlEdit);
                lblMessage2.Text = "ویرایش اطلاعات با موفقیت انجام شد.";
                lblMessage2.ForeColor = Color.YellowGreen;
                ResetItems();
            }
            else
            {
                lblMessage2.Text = "لطفا اطلاعات را به صورت کامل وارد نمایید!";
                lblMessage2.ForeColor = Color.LightCoral;
            }
        }

        private void picHazf2_Click(object sender, EventArgs e)
        {
            if (txtDescription.Text != "" && txtLinkAddress.Text != "" && cmbLinkgroup.SelectedItem != null && txtLinkKeywords.Text != "" && cmbLinkpriority.SelectedItem != null && cmbLinkrate.SelectedItem != null)
            {
                DialogResult dr = new DialogResult();
                dr = MessageBox.Show("حذف لینک !" + "\n\n" + "آیا از حذف لینک انتخاب شده  مطمئن هستید؟", "اخطار", MessageBoxButtons.YesNo, MessageBoxIcon.Question,MessageBoxDefaultButton.Button1,MessageBoxOptions.RightAlign);
                if (dr == DialogResult.Yes)
                {
                    string sqlDel = "DELETE FROM tbllinkbox WHERE id=" + txtId.Text;
                    database.DoIMD(sqlDel);
                    lblMessage2.Text = "لینک انتخاب شده حذف گردید.";
                    lblMessage2.ForeColor = Color.YellowGreen;
                    ResetItems();
                }
            }
            else
            {
                lblMessage2.Text = "لینک انتخاب نشده یا اطلاعات به صورت کامل وارد نشده است!";
                lblMessage2.ForeColor = Color.LightCoral;
            }
        }

        private void dgvManage_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
             DialogResult dr = new DialogResult();
                dr = MessageBox.Show("حذف لینک !" + "\n\n" + "آیا از حذف لینک انتخاب شده  مطمئن هستید؟", "اخطار", MessageBoxButtons.YesNo, MessageBoxIcon.Question,MessageBoxDefaultButton.Button1,MessageBoxOptions.RightAlign);
                if (dr == DialogResult.Yes)
                {
                    string sqlDel = "DELETE FROM tbllinkbox WHERE id=" + dgvManage.SelectedRows[0].Cells["id"].Value.ToString();
                    database.DoIMD(sqlDel);
                    MessageBox.Show("لینک انتخاب شده با موفقیت حذف شد","تایید حذف لینک",MessageBoxButtons.OK,MessageBoxIcon.Information,MessageBoxDefaultButton.Button1,MessageBoxOptions.RightAlign);
                }
                else
                {
                    e.Cancel = true;
                }
        }

        private void txtSearchs_TextChanged(object sender, EventArgs e)
        {
                if (chbDescription.CheckState==CheckState.Checked)
                {
                    if (txtSearchs.Text != "")
                    {
                        string sqlSearch1 = "SELECT url_description,url_location FROM tbllinkbox WHERE url_description LIKE N'%" + txtSearchs.Text + "%'";
                        System.Data.DataTable dtSearch = new System.Data.DataTable();
                        dtSearch = database.select(sqlSearch1);
                        dgvSearch.DataSource = dtSearch;
                        dgvSearch.Columns[1].HeaderText = "توضیحات لینک";
                        dgvSearch.Columns[2].HeaderText = "آدرس لینک";
                        dgvSearch.Columns[0].Width = 50;
                        dgvSearch.Columns[1].Width = 210;
                        dgvSearch.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        if (dgvSearch.RowCount > 0)
                        {
                            lblSearch.ForeColor = Color.YellowGreen;
                            lblSearch.Text = "تعداد پیدا شده : " + dgvSearch.RowCount;
                        }
                        else
                        {
                            lblSearch.Text = "جستجو نتیجه ای نداشت!";
                            lblSearch.ForeColor = Color.Coral;
                        }
                    }
                    else
                    {
                        lblSearch.Text = "لطفا برای جستجو کلمه یا حرفی را وارد نمایید.";
                        lblSearch.ForeColor = Color.LightCoral;
                        dgvSearch.DataSource = null;
                    }

                }
                else if(chbLocation.CheckState == CheckState.Checked)
                {
                    if (txtSearchs.Text != "")
                    {
                        string sqlSearch2 = "SELECT url_description,url_location FROM tbllinkbox WHERE url_location LIKE N'%" + txtSearchs.Text + "%'";
                        System.Data.DataTable dtSearch = new System.Data.DataTable();
                        dtSearch = database.select(sqlSearch2);
                        dgvSearch.DataSource = dtSearch;
                        dgvSearch.Columns[1].HeaderText = "توضیحات لینک";
                        dgvSearch.Columns[2].HeaderText = "آدرس لینک";
                        dgvSearch.Columns[0].Width = 50;
                        dgvSearch.Columns[1].Width = 210;
                        dgvSearch.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        if (dgvSearch.RowCount > 0)
                        {
                            lblSearch.ForeColor = Color.YellowGreen;
                            lblSearch.Text = "تعداد پیدا شده : " + dgvSearch.RowCount;
                        }
                        else
                        {
                            lblSearch.Text = "جستجو نتیجه ای نداشت!";
                            lblSearch.ForeColor = Color.Coral;
                        }
                    }
                    else
                    {
                        lblSearch.Text = "لطفا برای جستجو کلمه یا حرفی را وارد نمایید.";
                        lblSearch.ForeColor = Color.LightCoral;
                        dgvSearch.DataSource = null;
                    }
                }
                else if (chbTags.CheckState == CheckState.Checked)
                {
                    if (txtSearchs.Text != "")
                    {
                        string sqlSearch4 = "SELECT url_description,url_location FROM tbllinkbox WHERE url_tags LIKE N'%" + txtSearchs.Text + "%'";
                        System.Data.DataTable dtSearch = new System.Data.DataTable();
                        dtSearch = database.select(sqlSearch4);
                        dgvSearch.DataSource = dtSearch;
                        dgvSearch.Columns[1].HeaderText = "توضیحات لینک";
                        dgvSearch.Columns[2].HeaderText = "آدرس لینک";
                        dgvSearch.Columns[0].Width = 50;
                        dgvSearch.Columns[1].Width = 210;
                        dgvSearch.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        if (dgvSearch.RowCount > 0)
                        {
                            lblSearch.ForeColor = Color.YellowGreen;
                            lblSearch.Text = "تعداد پیدا شده : " + dgvSearch.RowCount;
                        }
                        else
                        {
                            lblSearch.Text = "جستجو نتیجه ای نداشت!";
                            lblSearch.ForeColor = Color.Coral;
                        }
                    }
                    else
                    {
                        lblSearch.Text = "لطفا برای جستجو کلمه یا حرفی را وارد نمایید.";
                        lblSearch.ForeColor = Color.LightCoral;
                        dgvSearch.DataSource = null;
                    }
                }
        }

        private void dgvSearch_CellFormatting_1(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvSearch.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
        }

        private void chbDescription_Click(object sender, EventArgs e)
        {
            chbDescription.CheckState = CheckState.Checked;
            chbLocation.CheckState = CheckState.Unchecked;
            chbGroup.CheckState = CheckState.Unchecked;
            chbTags.CheckState = CheckState.Unchecked;
            txtSearchs.Enabled = true;
            cmbSearchs.Enabled = false;
            cmbSearchs.SelectedItem = null;
            txtSearchs.Focus();
            System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("fa-ir");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);
        }

        private void chbLocation_Click(object sender, EventArgs e)
        {
            chbDescription.CheckState = CheckState.Unchecked;
            chbLocation.CheckState = CheckState.Checked;
            chbGroup.CheckState = CheckState.Unchecked;
            chbTags.CheckState = CheckState.Unchecked;
            txtSearchs.Enabled = true;
            cmbSearchs.Enabled = false;
            cmbSearchs.SelectedItem = null;
            txtSearchs.Focus();
            System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("EN");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);
        }

        private void chbGroup_Click(object sender, EventArgs e)
        {
            chbDescription.CheckState = CheckState.Unchecked;
            chbLocation.CheckState = CheckState.Unchecked;
            chbGroup.CheckState = CheckState.Checked;
            chbTags.CheckState = CheckState.Unchecked;
            txtSearchs.Enabled = false;
            cmbSearchs.Enabled = true;
            txtSearchs.Text = "";
        }

        private void chbTags_Click(object sender, EventArgs e)
        {
            chbDescription.CheckState = CheckState.Unchecked;
            chbLocation.CheckState = CheckState.Unchecked;
            chbGroup.CheckState = CheckState.Unchecked;
            chbTags.CheckState = CheckState.Checked;
            txtSearchs.Enabled = true;
            cmbSearchs.Enabled = false;
            cmbSearchs.SelectedItem = null;
            txtSearchs.Focus();
            System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("fa-ir");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);
        }

        private void cmbSearchs_TextChanged(object sender, EventArgs e)
        {
            if (cmbSearchs.SelectedItem != null)
            {
                string sqlSearch3 = "SELECT url_description,url_location FROM tbllinkbox WHERE url_group LIKE N'%" + cmbSearchs.Text + "%'";
                System.Data.DataTable dtSearch = new System.Data.DataTable();
                dtSearch = database.select(sqlSearch3);
                dgvSearch.DataSource = dtSearch;
                dgvSearch.Columns[1].HeaderText = "توضیحات لینک";
                dgvSearch.Columns[2].HeaderText = "آدرس لینک";
                dgvSearch.Columns[0].Width = 50;
                dgvSearch.Columns[1].Width = 210;
                dgvSearch.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                if (dgvSearch.RowCount > 0)
                {
                    lblSearch.ForeColor = Color.YellowGreen;
                    lblSearch.Text = "تعداد پیدا شده : " + dgvSearch.RowCount;
                }
                else
                {
                    lblSearch.Text = "جستجو نتیجه ای نداشت!";
                    lblSearch.ForeColor = Color.Coral;
                }
            }
            else
            {
                lblSearch.Text = "توجه ! شما هیچ پیغامی ندارید.";
                lblSearch.ForeColor = Color.LightCoral;
                dgvSearch.DataSource = null;
            }
        }

        private void txtLinkAddress_Enter(object sender, EventArgs e)
        {
            System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("EN");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);
        }

        private void txtLinkAddress_Leave(object sender, EventArgs e)
        {
            System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("fa-ir");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);
        }

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
          /*  if(e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Visible = false;
                this.notifyIcon.ShowBalloonTip(8000,"لینک باکس", "برنامه به تسکبار انتقال یافت", ToolTipIcon.Info);
            } */    
            System.Globalization.CultureInfo language = new System.Globalization.CultureInfo("EN");
            InputLanguage.CurrentInputLanguage = InputLanguage.FromCulture(language);
        }

        private void tscmsMainBrowser_Click(object sender, EventArgs e)
        {
            try
            {

                txtVisit.Clear();
                txtId.Clear();
                txtVisit.Text = dgvMainShow.SelectedRows[0].Cells["url_visited"].Value.ToString();
                txtId.Text = dgvMainShow.SelectedRows[0].Cells["id"].Value.ToString();
                int sr;
                sr = int.Parse(txtVisit.Text);
                sr += 1;
                string sqlIns = "UPDATE tbllinkbox SET url_visited=" + sr + " WHERE id=" + txtId.Text + "";
                database.DoIMD(sqlIns);
                if(PrBrowser == null)
                    System.Diagnostics.Process.Start("IExplore.exe", dgvMainShow.SelectedRows[0].Cells["url_location"].Value.ToString());
                else
                    System.Diagnostics.Process.Start(PrBrowser, dgvMainShow.SelectedRows[0].Cells["url_location"].Value.ToString());
            }
            catch { }
        }

        private void tscmsShowBrowser_Click(object sender, EventArgs e)
        {
            try
            {
                txtVisit.Clear();
                txtId.Clear();
                txtVisit.Text = dgvShowPrint.SelectedRows[0].Cells["url_visited"].Value.ToString();
                txtId.Text = dgvShowPrint.SelectedRows[0].Cells["id"].Value.ToString();
                int sr;
                sr = int.Parse(txtVisit.Text);
                sr += 1;
                string sqlIns = "UPDATE tbllinkbox SET url_visited=" + sr + " WHERE id=" + txtId.Text + "";
                database.DoIMD(sqlIns);
                if(PrBrowser == null)
                    System.Diagnostics.Process.Start("IExplore.exe",dgvShowPrint.SelectedRows[0].Cells["url_location"].Value.ToString());
                else
                    System.Diagnostics.Process.Start(PrBrowser, dgvShowPrint.SelectedRows[0].Cells["url_location"].Value.ToString());
            }
            catch { }
        }

        private void dgvTopShowLink_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dgvTopShowLink.Rows[e.RowIndex].Cells[0].Value = e.RowIndex + 1;
        }

        private void tsmTopShow_Click(object sender, EventArgs e)
        {
            try
            {
                txtVisit.Clear();
                txtId.Clear();
                txtVisit.Text = dgvTopShowLink.SelectedRows[0].Cells["url_visited"].Value.ToString();
                txtId.Text = dgvTopShowLink.SelectedRows[0].Cells["id"].Value.ToString();
                int sr;
                sr = int.Parse(txtVisit.Text);
                sr += 1;
                string sqlIns = "UPDATE tbllinkbox SET url_visited=" + sr + " WHERE id=" + txtId.Text + "";
                database.DoIMD(sqlIns);
                if(PrBrowser == null)
                    System.Diagnostics.Process.Start("IExplore.exe", dgvTopShowLink.SelectedRows[0].Cells["url_location"].Value.ToString());
                else
                    System.Diagnostics.Process.Start(PrBrowser, dgvTopShowLink.SelectedRows[0].Cells["url_location"].Value.ToString());
            }
            catch { }
        }

        private void tsmextract_Click(object sender, EventArgs e)
        {
            int k;
            FileStream fs = new FileStream(System.Windows.Forms.Application.StartupPath + @"\Export\Export.txt", FileMode.Create ,FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            k = dgvManage.Rows.Count;
            for(int i=0;i<=k-1;i++)
            {
                sw.WriteLine(dgvManage.Rows[i].Cells["url_description"].Value.ToString() + "\t" + dgvManage.Rows[i].Cells["url_location"].Value.ToString());
            }
            sw.Close();
            System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + @"\Export\Export.txt");
        }

        private void tscmsPrint_Click(object sender, EventArgs e)
        {
            if(SetupPrinting())
            {
                PrintPreviewDialog ppd = new PrintPreviewDialog();
                ppd.Document = printDoc;
                ppd.ShowDialog();
            }
        }

        private void tsmPrintt_Click(object sender, EventArgs e)
        {
            if (SetupPrinting())
            {
                PrintPreviewDialog ppd = new PrintPreviewDialog();
                ppd.Document = printDoc;
                ppd.ShowDialog();
            }
        }

        private void picPrints2_Click(object sender, EventArgs e)
        {
            if (SetupPrinting())
            {
                PrintPreviewDialog ppd = new PrintPreviewDialog();
                ppd.Document = printDoc;
                ppd.ShowDialog();
            }
        }

        private void printDoc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            bool more = dgvPrint.DrawDataGridView(e.Graphics);
            if (more)
                e.HasMorePages = true;
        }

        private void tsmBackupRestore_Click(object sender, EventArgs e)
        {
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.None;
            picSetting.BorderStyle = BorderStyle.FixedSingle;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlExtract.Visible = true;
            pnlTopShow.Visible = false;
            pnlSearch.Visible = false;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlAboutUs.Visible = false;
            pnlSettings.Visible = false;
        }

        private void notifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {

        }

        private void picExtra1_MouseHover(object sender, EventArgs e)
        {
            picExtra1.Visible = false;
            picExtra2.Visible = true;
        }

        private void picExtra2_MouseLeave(object sender, EventArgs e)
        {
            picExtra2.Visible = false;
            picExtra1.Visible = true;
        }

        private void btnOpenDialog_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents(*.xlsx)|*.xlsx";
            sfd.FileName = "Export";
            //sfd.ShowDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                txtExtractAddress.Text = sfd.FileName;
                picExtra1.Enabled = true;
            }
            else
            {
                picExtra1.Enabled = false;
                txtExtractAddress.Clear();
            }
        }

        private void picExtra2_Click(object sender, EventArgs e)
        {
            try
            {
                object miss = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = false;
                worksheet = (Worksheet)workbook.Sheets["Sheet1"];
                worksheet = (Worksheet)workbook.ActiveSheet;
                worksheet.Name = "Export";

                for (int i = 1; i < dgvShowPrint.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dgvShowPrint.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dgvShowPrint.Rows.Count; i++)
                {
                    for (int j = 0; j < dgvShowPrint.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dgvShowPrint.Rows[i].Cells[j].Value.ToString();
                    }
                }
                if (txtExtractAddress.Text != "")
                    workbook.SaveAs(txtExtractAddress.Text, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
   
                app.Quit();
                System.Diagnostics.Process.Start(txtExtractAddress.Text);
                txtExtractAddress.Clear();
            }
            catch { }
        }

        private void txtExtractAddress_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Documents(*.xlsx)|*.xlsx";
            sfd.FileName = "Export";
            //sfd.ShowDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                txtExtractAddress.Text = sfd.FileName;
                picExtra1.Enabled = true;
            }
            else
            {
                picExtra1.Enabled = false;
                txtExtractAddress.Clear();
            }
        }

        private void picExtractNote1_MouseHover(object sender, EventArgs e)
        {
            picExtractNote1.Visible = false;
            picExtractNote2.Visible = true;
        }

        private void picExtractNote2_MouseLeave(object sender, EventArgs e)
        {
            picExtractNote2.Visible = false;
            picExtractNote1.Visible = true;
        }

        private void picExtractNote2_Click(object sender, EventArgs e)
        {
            int k;
            FileStream fs = new FileStream(txtExtractNote.Text, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs);
            k = dgvShowPrint.Rows.Count;
            for (int i = 0; i <= k - 1; i++)
            {
                sw.WriteLine(dgvShowPrint.Rows[i].Cells["url_description"].Value.ToString() + "\t\t\t" + dgvShowPrint.Rows[i].Cells["url_location"].Value.ToString());
            }
            sw.Close();
            System.Diagnostics.Process.Start(txtExtractNote.Text);
            txtExtractNote.Clear();
        }

        private void btnOpenDialogNote_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text Documents(*.txt)|*.txt";
            sfd.FileName = "Export";
            //sfd.ShowDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                txtExtractNote.Text = sfd.FileName;
                picExtractNote1.Enabled = true;
            }
            else
            {
                picExtractNote1.Enabled = false;
                txtExtractNote.Clear();
            }
        }

        private void txtExtractNote_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Text Documents(*.txt)|*.txt";
            sfd.FileName = "Export";
            //sfd.ShowDialog();
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                txtExtractNote.Text = sfd.FileName;
                picExtractNote1.Enabled = true;
            }
            else
            {
                picExtractNote1.Enabled = false;
                txtExtractNote.Clear();
            }
        }

        private void picrestore1_MouseHover(object sender, EventArgs e)
        {
            picrestore1.Visible = false;
            picrestore2.Visible = true;
        }

        private void picrestore2_MouseLeave(object sender, EventArgs e)
        {
            picrestore2.Visible = false;
            picrestore1.Visible = true;
        }

        private void picBackup1_MouseHover(object sender, EventArgs e)
        {
            picBackup1.Visible = false;
            picBackup2.Visible = true;
        }

        private void picBackup2_MouseLeave(object sender, EventArgs e)
        {
            picBackup2.Visible = false;
            picBackup1.Visible = true;
        }

        private void picBackup2_Click(object sender, EventArgs e)
        {
                //string backupQuery = @"BACKUP DATABASE [Linkboxdb] TO DISK = N'" + txtBackup.Text + "' WITH NOFORMAT, NOINIT,  NAME = N'Linkbox-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10";
                //database.DoIMD(backupQuery); 
                //File.Copy(System.Windows.Forms.Application.StartupPath + @"\db\Linkboxdb.mdf", txtBackup.Text,true);
            try
            {
                if (File.Exists(txtBackup.Text))
                {
                    DialogResult dr = new DialogResult();
                    dr = MessageBox.Show("توجه فایلی با همین نام وجود دارد" + "\n\n" + "آیا تمایلی به گرفتن فایل پشتیبان دارید؟", "هشدار ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                    if (dr == DialogResult.Yes)
                    {
                        SqlConnection con = new SqlConnection();
                        con.ConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\db\Linkboxdb.mdf;Integrated Security=True;Connect Timeout=15";
                        SqlCommand sc = new SqlCommand();
                        sc.Connection = con;
                        SqlConnection.ClearAllPools();
                        con.Open();
                        string dbName = con.Database;
                        string backupQuery = @"BACKUP DATABASE [" + dbName + "] TO DISK = N'" + txtBackup.Text + "' WITH NO_COMPRESSION ,CONTINUE_AFTER_ERROR  ,FORMAT, INIT, NAME = N'Linkboxdb-Full Database Backup', SKIP,NOREWIND, NOUNLOAD, STATS = 10 ";
                        sc.CommandText = backupQuery;
                        sc.ExecuteNonQuery();
                        con.Close();
                        MessageBox.Show("پشتیبان گیری با موفقیت انجام گردید", "پشتیبان گیری");
                        txtBackup.Clear();
                    }
                }
            }
            catch { MessageBox.Show("پشتیبان گیری با خطا روبرو شده است", "اخطار"); }
             
        }

        private void picrestore2_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult dr = new DialogResult();
                dr = MessageBox.Show("توجه ممکن است با بازگردانی فایل،اطلاعاتی که پشتیبان گیری نشده است را از دست دهید" + "\n\n" + "آیا نسبت به بازگردانی فایل مطمئن هستید؟", "هشدار ", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                if (dr == DialogResult.Yes)
                {
                    SqlConnection con = new SqlConnection();
                    con.ConnectionString = @"Data Source=(LocalDB)\v11.0;AttachDbFilename=|DataDirectory|\db\Linkboxdb.mdf;Integrated Security=True;Connect Timeout=15";
                    SqlCommand sc = new SqlCommand();
                    sc.Connection = con;
                    SqlConnection.ClearAllPools();
                    con.Open();
                    string dbName = con.Database;
                    string RestoreQuery = @"USE [master]; RESTORE DATABASE [" + dbName + "] FROM  DISK = N'" + txtRestore.Text + "' WITH NOUNLOAD,  REPLACE,  STATS = 10";
                    sc.CommandText = RestoreQuery;
                    sc.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("بازگردانی با موفقیت انجام گردید", "بازگردانی");
                    txtRestore.Clear();
                }
            }
            catch { MessageBox.Show("بازگردانی با خطا روبرو شده است", "اخطار"); }
        }

        private void btnSaveDialogBackup_Click(object sender, EventArgs e)
        {
            PersianCalendar pc = new PersianCalendar();
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Backup Database(*.bac)|*.bac";
            sfd.FileName = "Linkbox_Backup_in_" + pc.GetYear(DateTime.Now).ToString() + "_" + pc.GetMonth(DateTime.Now).ToString() + "_" + pc.GetDayOfMonth(DateTime.Now).ToString();
            //sfd.ShowDialog();
            //sfd.FileName = "Linkboxdb.mdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                txtBackup.Text = sfd.FileName.ToString();
                picBackup1.Enabled = true;
            }
            else
            {
                picBackup1.Enabled = false;
                txtBackup.Clear();
            }
        }

        private void btnOpenDialogRestore_Click(object sender, EventArgs e)
        {
            PersianCalendar pc = new PersianCalendar();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Backup Database(*.bac)|*.bac";
            //ofd.FileName = "Linkbox_Backup_in_" + pc.GetYear(DateTime.Now).ToString() + "_" + pc.GetMonth(DateTime.Now).ToString() + "_" + pc.GetDayOfMonth(DateTime.Now).ToString();
            //sfd.ShowDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtRestore.Text = ofd.FileName.ToString();
                picrestore1.Enabled = true;
            }
            else
            {
                picrestore1.Enabled = false;
                txtRestore.Clear();
            }
        }

        private void label26_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("IExplore.exe", "www.NSP-Team.com");
            }
            catch { }
        }

        private void picSabtSet1_MouseHover(object sender, EventArgs e)
        {
            picSabtSet1.Visible = false;
            picSabtSet2.Visible = true;
        }

        private void picSabtSet2_MouseLeave(object sender, EventArgs e)
        {
            picSabtSet2.Visible = false;
            picSabtSet1.Visible = true;
        }

        private void picSabtPass1_MouseHover(object sender, EventArgs e)
        {
            picSabtPass1.Visible = false;
            picSabtPass2.Visible = true;
        }

        private void picSabtPass2_MouseLeave(object sender, EventArgs e)
        {
            picSabtPass2.Visible = false;
            picSabtPass1.Visible = true;
        }

        private void picEditPass1_MouseHover(object sender, EventArgs e)
        {
            picEditPass1.Visible = false;
            picEditPass2.Visible = true;
        }

        private void picEditPass2_MouseLeave(object sender, EventArgs e)
        {
            picEditPass2.Visible = false;
            picEditPass1.Visible = true;
        }

        private void picSabtPass2_Click(object sender, EventArgs e)
        {
            pass_dgv();
            int flag = 1;
                if (txtPassword.Text != "" && txtPasswordReapeat.Text != "")
                {
                    if (dataGridView1.RowCount >= 1)
                    {
                        MessageBox.Show("ثبت کلمه عبور با محدودیت ثبت مواجه گردیده است" + "\n\n" + "کلمه عبور برنامه فعال است", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                        txtPassword.Clear();
                        txtPasswordReapeat.Clear();
                        txtPassword.Focus();
                    }
                    else
                    {
                        if (txtPassword.Text == txtPasswordReapeat.Text)
                        {
                            string CreatePass = "INSERT INTO tblpass (apppassword,flag) VALUES " + " (N'" + txtPassword.Text + "'," + flag + ") ";
                            database.DoIMD(CreatePass);
                            MessageBox.Show("کلمه عبور ثبت گردید", "ثبت کلمه عبور ", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                            txtPassword.Clear();
                            txtPasswordReapeat.Clear();
                        }
                        else
                        {
                            MessageBox.Show("کلمه عبور و تکرار کلمه عبور یکی نیستند", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                            txtPassword.Focus();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("کلمه عبور یا تکرار کلمه عبور وارد نشده است" + "\n\n" + "لطفا کلمه عبور یا تکرار کلمه عبور را وارد نمایید", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                    txtPassword.Focus();
                }
        }
        public string sdx;
        private void picEditPass2_Click(object sender, EventArgs e)
        {
            pass_dgv();
            if (dataGridView1.RowCount > 0)
            {
                if (txtOldPassword.Text != "" && txtNewPasswordReapeat.Text != "" && txtNewPassword.Text != "")
                {
                    string existdata = dataGridView1.Rows[0].Cells["apppassword"].Value.ToString();
                    if (txtOldPassword.Text == existdata)
                    {
                        if (txtNewPassword.Text == txtNewPasswordReapeat.Text)
                        {
                            string EditPass = "UPDATE tblpass SET apppassword=N'" + txtNewPassword.Text + "' WHERE apppassword=" + existdata + " ";
                            database.DoIMD(EditPass);
                            MessageBox.Show("ویرایش کلمه عبور با موفقیت ثبت گردید", "ویرایش کلمه عبور ", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                            txtOldPassword.Clear();
                            txtNewPasswordReapeat.Clear();
                            txtNewPassword.Clear();
                            txtOldPassword.Focus();
                        }
                        else
                        {
                            MessageBox.Show("کلمه عبور و تکرار کلمه عبور یکی نیستند", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                            txtNewPassword.Focus();
                        }
                    }
                    else
                    {
                        MessageBox.Show("کلمه عبور فعلی اشتباه است" + "\n\n" + "لطفا کلمه عبور فعلی درست را وارد نمایید", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                        txtOldPassword.Focus();
                    }
                }
                else
                {
                    MessageBox.Show("کلمه عبور یا تکرار کلمه عبور وارد نشده است", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                    txtPassword.Focus();
                }
            }
            else
            {
                MessageBox.Show("در حال حاضر کلمه عبوری برای برنامه در نظر گرفته نشده است" + "\n\n" + "ابتدا کلمه عبور برنامه را فعال نمایید ", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                txtOldPassword.Clear();
                txtNewPasswordReapeat.Clear();
                txtNewPassword.Clear();
                txtOldPassword.Focus();
            }
        }
        public string PrBrowser = null;
        private void picSabtSet2_Click(object sender, EventArgs e)
        {
            try
            {
                if (rbtIe.Checked)
                {
                    PrBrowser = "IExplore.exe";
                    MessageBox.Show("مرورگر به روی اینترنت اکسپلورر تنظیم شد" + "\n\n" + "", "تنظیم مرورگر ", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                }
                else if (rbtChrome.Checked)
                {
                    PrBrowser = "chrome.exe";
                    MessageBox.Show("مرورگر به روی گوگل کروم تنظیم شد" + "\n\n" + "", "تنظیم مرورگر ", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                }
                else if (rbtFirefox.Checked)
                {
                    PrBrowser = "firefox.exe";
                    MessageBox.Show("مرورگر به روی فایرفاکس تنظیم شد" + "\n\n" + "", "تنظیم مرورگر ", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                }
                else if (rbtOpera.Checked)
                {
                    PrBrowser = "opera.exe";
                    MessageBox.Show("مرورگر به روی اوپرا تنظیم شد" + "\n\n" + "", "تنظیم مرورگر ", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                }

            }
            catch
            {
                MessageBox.Show("مرورگری انتخاب نشده است" + "\n\n" + "یا ممکن است مرورگری که انتخاب نموده اید نصب نباشد", "هشدار ", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            }
        }

        private void picCancel1_MouseHover(object sender, EventArgs e)
        {
            picCancel1.Visible = false;
            picCancel2.Visible = true;
        }

        private void picCancel2_MouseLeave(object sender, EventArgs e)
        {
            picCancel2.Visible = false;
            picCancel1.Visible = true;
        }

        private void picLogin1_MouseHover(object sender, EventArgs e)
        {
            picLogin1.Visible = false;
            picLogin2.Visible = true;
        }

        private void picLogin2_MouseLeave(object sender, EventArgs e)
        {
            picLogin2.Visible = false;
            picLogin1.Visible = true;
        }

        private void lblRecoverPass_MouseHover(object sender, EventArgs e)
        {
            lblRecoverPass.ForeColor = Color.SteelBlue;
        }

        private void lblRecoverPass_MouseLeave(object sender, EventArgs e)
        {
            lblRecoverPass.ForeColor = Color.White;
        }

        private void picLogin2_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtLogin.Text != "")
                {
                    pass_dgv();
                    string Login = dataGridView1.Rows[0].Cells["apppassword"].Value.ToString();
                    if (txtLogin.Text == Login)
                    {
                        pnlLogin.Visible = false;
                    }   
                    else
                    {
                        pnlLogin.Visible = true;
                        lblLoginMessage.Text = "کلمه عبور اشتباه می باشد !";
                    }
                }
                else
                {
                    lblLoginMessage.Text = "لطفا ابتدا کلمه عبور را وارد نمایید !";
                    txtLogin.Focus();
                }
            }
            catch { }
        }

        private void picCancel2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void txtLogin_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void tsmCloseApp_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void txtLogin_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                try
                {
                    if (txtLogin.Text != "")
                    {
                        pass_dgv();
                        string Login = dataGridView1.Rows[0].Cells["apppassword"].Value.ToString();
                        if (txtLogin.Text == Login)
                        {
                            pnlLogin.Visible = false;
                            txtLogin.Text = null;
                        }
                        else
                        {
                            txtLogin.Text = null;
                            lblLoginMessage.Text = "کلمه عبور اشتباه می باشد !";
                            pnlLogin.Visible = true;
                        }
                    }
                    else
                    {
                        lblLoginMessage.Text = "لطفا ابتدا کلمه عبور را وارد نمایید !";
                        txtLogin.Focus();
                        txtLogin.Text = null;
                    }
                }
                catch { }
            }
        }

        private void tsmInsLink_Click(object sender, EventArgs e)
        {
            if (!pnlLogin.Visible)
            {
                /*frmMain fm = new frmMain();
                fm.Visible = true;
                fm.WindowState = FormWindowState.Normal;*/
                pnlmenu.Visible = true;
                picShowMainPanel1.Visible = true;
                lblMessage2.Text = "توجه ! شما هیچ پیغامی ندارید.";
                lblMessage2.ForeColor = Color.LightCoral;
                picshowprint.BorderStyle = BorderStyle.FixedSingle;
                picInsertlink.BorderStyle = BorderStyle.None;
                picManagelink.BorderStyle = BorderStyle.FixedSingle;
                picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
                picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
                picExport.BorderStyle = BorderStyle.FixedSingle;
                picSetting.BorderStyle = BorderStyle.FixedSingle;
                picAboutUs.BorderStyle = BorderStyle.FixedSingle;
                picExit.BorderStyle = BorderStyle.FixedSingle;
                pnlinsert.Visible = true;
                pnlShowPrint.Visible = false;
                pnlManage.Visible = false;
                pnlSearch.Visible = false;
                pnlTopShow.Visible = false;
                pnlExtract.Visible = false;
                pnlAboutUs.Visible = false;
                pnlSettings.Visible = false;
                ResetItems();
            }
            else
            {
                txtLogin.Focus();
            }
        }

        private void tsmTopShown_Click(object sender, EventArgs e)
        {
            if(!pnlLogin.Visible)
            {
                pnlmenu.Visible = true;
                picShowMainPanel1.Visible = true;
                picshowprint.BorderStyle = BorderStyle.FixedSingle;
                picInsertlink.BorderStyle = BorderStyle.FixedSingle;
                picManagelink.BorderStyle = BorderStyle.FixedSingle;
                picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
                picTopShowLink.BorderStyle = BorderStyle.None;
                picExport.BorderStyle = BorderStyle.FixedSingle;
                picSetting.BorderStyle = BorderStyle.FixedSingle;
                picAboutUs.BorderStyle = BorderStyle.FixedSingle;
                picExit.BorderStyle = BorderStyle.FixedSingle;
                pnlTopShow.Visible = true;
                pnlSearch.Visible = false;
                pnlShowPrint.Visible = false;
                pnlinsert.Visible = false;
                pnlManage.Visible = false;
                pnlExtract.Visible = false;
                pnlAboutUs.Visible = false;
                pnlSettings.Visible = false;
                TopShowLink();
            }
            else
            {
                txtLogin.Focus();
            }
        }

        private void tsmBackupRestor_Click(object sender, EventArgs e)
        {
            if(!pnlLogin.Visible)
            {
                pnlmenu.Visible = true;
                picShowMainPanel1.Visible = true;
                picshowprint.BorderStyle = BorderStyle.FixedSingle;
                picInsertlink.BorderStyle = BorderStyle.FixedSingle;
                picManagelink.BorderStyle = BorderStyle.FixedSingle;
                picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
                picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
                picExport.BorderStyle = BorderStyle.None;
                picSetting.BorderStyle = BorderStyle.FixedSingle;
                picAboutUs.BorderStyle = BorderStyle.FixedSingle;
                picExit.BorderStyle = BorderStyle.FixedSingle;
                pnlExtract.Visible = true;
                pnlTopShow.Visible = false;
                pnlSearch.Visible = false;
                pnlShowPrint.Visible = false;
                pnlinsert.Visible = false;
                pnlManage.Visible = false;
                pnlAboutUs.Visible = false;
                pnlSettings.Visible = false;
                txtExtractAddress.Clear();
                txtExtractNote.Clear();
                txtBackup.Clear();
                txtRestore.Clear();
            }
            else
            {
                txtLogin.Focus();
            }
        }

        private void tsmSearchLinks_Click(object sender, EventArgs e)
        {
            if(!pnlLogin.Visible)
            {
                pnlmenu.Visible = true;
                picShowMainPanel1.Visible = true;
                picshowprint.BorderStyle = BorderStyle.FixedSingle;
                picInsertlink.BorderStyle = BorderStyle.FixedSingle;
                picManagelink.BorderStyle = BorderStyle.FixedSingle;
                picAdvanceSearch.BorderStyle = BorderStyle.None;
                picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
                picExport.BorderStyle = BorderStyle.FixedSingle;
                picSetting.BorderStyle = BorderStyle.FixedSingle;
                picAboutUs.BorderStyle = BorderStyle.FixedSingle;
                picExit.BorderStyle = BorderStyle.FixedSingle;
                pnlSearch.Visible = true;
                pnlShowPrint.Visible = false;
                pnlinsert.Visible = false;
                pnlManage.Visible = false;
                pnlTopShow.Visible = false;
                pnlExtract.Visible = false;
                pnlAboutUs.Visible = false;
                pnlSettings.Visible = false;
                txtSearchs.Text = "";
                cmbSearchs.SelectedItem = null;
                lblSearch.Text = "توجه ! شما هیچ پیغامی ندارید.";
                lblSearch.ForeColor = Color.LightCoral;
                chbDescription.CheckState = CheckState.Checked;
                chbLocation.CheckState = CheckState.Unchecked;
                chbGroup.CheckState = CheckState.Unchecked;
                chbTags.CheckState = CheckState.Unchecked;
                txtSearchs.Focus();
                cmbSearchs.Enabled = false;
            }
            else
            {
                txtLogin.Focus();
            }
        }

        private void tsmSettingLink_Click(object sender, EventArgs e)
        {
            if(!pnlLogin.Visible)
            {
                pnlmenu.Visible = true;
                picShowMainPanel1.Visible = true;
                pass_dgv();
                picshowprint.BorderStyle = BorderStyle.FixedSingle;
                picInsertlink.BorderStyle = BorderStyle.FixedSingle;
                picManagelink.BorderStyle = BorderStyle.FixedSingle;
                picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
                picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
                picExport.BorderStyle = BorderStyle.FixedSingle;
                picSetting.BorderStyle = BorderStyle.None;
                picAboutUs.BorderStyle = BorderStyle.FixedSingle;
                picExit.BorderStyle = BorderStyle.FixedSingle;
                pnlAboutUs.Visible = false;
                pnlTopShow.Visible = false;
                pnlSearch.Visible = false;
                pnlShowPrint.Visible = false;
                pnlinsert.Visible = false;
                pnlManage.Visible = false;
                pnlExtract.Visible = false;
                pnlSettings.Visible = true;
                txtPassword.Focus();
            }
            else
            {
                txtLogin.Focus();
            }
        }
        private void dgvMainShow_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tsmSettings_Click(object sender, EventArgs e)
        {
            pass_dgv();
            picshowprint.BorderStyle = BorderStyle.FixedSingle;
            picInsertlink.BorderStyle = BorderStyle.FixedSingle;
            picManagelink.BorderStyle = BorderStyle.FixedSingle;
            picAdvanceSearch.BorderStyle = BorderStyle.FixedSingle;
            picTopShowLink.BorderStyle = BorderStyle.FixedSingle;
            picExport.BorderStyle = BorderStyle.FixedSingle;
            picSetting.BorderStyle = BorderStyle.None;
            picAboutUs.BorderStyle = BorderStyle.FixedSingle;
            picExit.BorderStyle = BorderStyle.FixedSingle;
            pnlAboutUs.Visible = false;
            pnlTopShow.Visible = false;
            pnlSearch.Visible = false;
            pnlShowPrint.Visible = false;
            pnlinsert.Visible = false;
            pnlManage.Visible = false;
            pnlExtract.Visible = false;
            pnlSettings.Visible = true;
            txtPassword.Focus();
        }

        private void lblRecoverPass_Click(object sender, EventArgs e)
        {
            MessageBox.Show("بازیابی کلمه عبور فعال نمی باشد" + "\n\n" + "این گزینه در نسخه بعدی برنامه فعال خواهد شد لطفا منتظر نسخه بعدی باشید", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
            txtLogin.Focus();
        }

    }
}
