using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace CL_Cronos_Data
{
    public partial class Main_Form : Form
    {
        private string __root_url = "http://sn.gk001.gpk456.com";
        private string __url = "";
        private string __start_datetime_elapsed;
        private string __file_location = "\\\\192.168.10.22\\ssi-reporting";
        private string __brand_code = "CL";
        private string __brand_color_hex = "#1B60A8";
        private string __brand_color_rgb = "27, 96, 168";
        private string __app = "Cronos Data";
        private string __app_type = "2";
        private int __send = 0;
        private int __timer_count = 10;
        private int __page_size = 100000;
        private int __index_reg = 1;
        private int __index_dep = 1;
        private int __index_bon = 1;
        private int __index_bet = 1;
        private int __display_count = 0;
        private bool __is_close;
        private bool __is_login = false;
        private bool __is_start = false;
        private bool __is_autostart = true;
        private bool __detect_header = false;
        private bool __is_send = true;
        private JObject __jo;
        private JToken __jo_count;
        private JToken __conn_id = "";
        StringBuilder __DATA = new StringBuilder();
        List<String> __getdata_viplist = new List<String>();
        List<String> __getdata_affiliatelist = new List<String>();
        List<String> __getdata_paymentmethodlist = new List<String>();
        List<String> __getdata_bonuscode = new List<String>();
        List<String> __getdata_productcode = new List<String>();
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        Form __mainform_handler;


        // Drag Header to Move
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        // Form Shadow
        private bool m_aeroEnabled;
        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );
        [DllImport("dwmapi.dll")]
        public static extern int DwmExtendFrameIntoClientArea(IntPtr hWnd, ref MARGINS pMarInset);
        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
        [DllImport("dwmapi.dll")]
        public static extern int DwmIsCompositionEnabled(ref int pfEnabled);
        private const int CS_DROPSHADOW = 0x00020000;
        private const int WM_NCPAINT = 0x0085;
        private const int WM_ACTIVATEAPP = 0x001C;
        private const int WM_NCHITTEST = 0x84;
        private const int HTCLIENT = 0x1;
        private const int HTCAPTION = 0x2;
        private const int WS_MINIMIZEBOX = 0x20000;
        private const int CS_DBLCLKS = 0x8;
        public struct MARGINS
        {
            public int leftWidth;
            public int rightWidth;
            public int topHeight;
            public int bottomHeight;
        }
        protected override CreateParams CreateParams
        {
            get
            {
                m_aeroEnabled = CheckAeroEnabled();

                CreateParams cp = base.CreateParams;
                if (!m_aeroEnabled)
                    cp.ClassStyle |= CS_DROPSHADOW;

                cp.Style |= WS_MINIMIZEBOX;
                cp.ClassStyle |= CS_DBLCLKS;
                return cp;
            }
        }
        private bool CheckAeroEnabled()
        {
            if (Environment.OSVersion.Version.Major >= 6)
            {
                int enabled = 0;
                DwmIsCompositionEnabled(ref enabled);
                return (enabled == 1) ? true : false;
            }
            return false;
        }
        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_NCPAINT:
                    if (m_aeroEnabled)
                    {
                        var v = 2;
                        DwmSetWindowAttribute(Handle, 2, ref v, 4);
                        MARGINS margins = new MARGINS()
                        {
                            bottomHeight = 1,
                            leftWidth = 0,
                            rightWidth = 0,
                            topHeight = 0
                        };
                        DwmExtendFrameIntoClientArea(Handle, ref margins);

                    }
                    break;
                default:
                    break;
            }
            base.WndProc(ref m);

            if (m.Msg == WM_NCHITTEST && (int)m.Result == HTCLIENT)
                m.Result = (IntPtr)HTCAPTION;
        }
        // ----- Form Shadow

        public Main_Form()
        {
            InitializeComponent();

            Opacity = 0;
            timer.Interval = 20;
            timer.Tick += new EventHandler(FadeIn);
            timer.Start();
        }

        private void FadeIn(object sender, EventArgs e)
        {
            if (Opacity >= 1)
            {
                timer_landing.Start();
            }
            else
            {
                Opacity += 0.05;
            }
        }

        private void panel_header_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void label_title_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void panel_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void pictureBox_minimize_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void pictureBox_close_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Exit the program?", __brand_code + " Cronos Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dr == DialogResult.Yes)
            {
                __is_close = false;
                Environment.Exit(0);
            }
        }

        private void panel_cl_Paint(object sender, PaintEventArgs e)
        {
            Rectangle rect = panel_cl.ClientRectangle;
            rect.Width--;
            rect.Height--;
            e.Graphics.DrawRectangle(Pens.LightGray, rect);
        }

        private bool __detect_navigate = false;
        private void timer_landing_Tick(object sender, EventArgs e)
        {
            if (!__detect_navigate)
            {
                webBrowser.Navigate(__root_url + "/Account/Login");
            }
            __detect_navigate = true;
            // comment
            panel_landing.Visible = false;
            label_title.Visible = true;
            panel.Visible = true;
            pictureBox_minimize.Visible = true;
            pictureBox_close.Visible = true;
            label_version.Visible = true;
            label_status.Visible = true;
            label_status_1.Visible = true;
            label_cycle_in.Visible = true;
            label_cycle_in_1.Visible = true;
            button1.Visible = true;
            timer_landing.Stop();
        }

        private void Main_Form_Load(object sender, EventArgs e)
        {            
            comboBox.SelectedIndex = 0;
            comboBox_list.SelectedIndex = 0;
            dateTimePicker_start.Format = DateTimePickerFormat.Custom;
            dateTimePicker_start.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            dateTimePicker_end.Format = DateTimePickerFormat.Custom;
            dateTimePicker_end.CustomFormat = "yyyy-MM-dd HH:mm:ss";
        }

        private void Main_Form_Shown(object sender, EventArgs e)
        {
            ___GETDATA_VIPLIST();
            ___GETDATA_AFFIALIATELIST();
            ___GETDATA_PAYMENTMETHODLIST();
            ___GETDATA_BONUSCODE();
            ___GETDATA_PRODUCTCODE();
        }

        // WebBrowser
        private async void WebBrowser_DocumentCompletedAsync(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (webBrowser.ReadyState == WebBrowserReadyState.Complete)
            {
                if (e.Url == webBrowser.Url)
                {
                    try
                    {
                        if (webBrowser.Url.ToString().Equals(__root_url + "/Account/Login"))
                        {
                            if (__is_login)
                            {
                                pictureBox_loader.Visible = false;
                                label_page_count.Visible = false;
                                label_total_records.Visible = false;
                                button_start.Visible = false;
                                // comment
                                __mainform_handler = Application.OpenForms[0];
                                __mainform_handler.Size = new Size(569, 514);
                                panel_loader.Visible = false;
                                label_navigate_up.Enabled = false;

                                // comment
                                //SendITSupport("The application have been logout, please re-login again.");
                                SendMyBot("The application have been logout, please re-login again.");
                                __send = 0;
                            }

                            __is_login = false;
                            timer.Stop();
                            webBrowser.Document.Body.Style = "zoom:.8";
                            webBrowser.Visible = true;
                            webBrowser.WebBrowserShortcutsEnabled = true;
                            label_status.Text = "Logout";
                        }

                        if (webBrowser.Url.ToString().Equals("http://sn.gk001.gpk456.com/"))
                        {
                            pictureBox_loader.Visible = true;
                            label_page_count.Visible = true;
                            label_total_records.Visible = true;
                            button_start.Visible = true;
                            // comment
                            __mainform_handler = Application.OpenForms[0];
                            __mainform_handler.Size = new Size(569, 208);
                            panel_loader.Visible = true;
                            label_navigate_up.Enabled = false;

                            if (!__is_login)
                            {
                                __is_login = true;
                                webBrowser.Visible = false;
                                pictureBox_loader.Visible = true;
                            }

                            if (!__is_start)
                            {
                                if (Properties.Settings.Default.______start_detect == "0")
                                {
                                    button_start.Enabled = false;
                                    panel_filter.Enabled = false;
                                    label_status.Text = "Waiting";
                                }
                                // registration
                                else if (Properties.Settings.Default.______start_detect == "1")
                                {
                                    comboBox_list.SelectedIndex = 0;
                                    button_start.PerformClick();
                                }
                                // payment
                                else if (Properties.Settings.Default.______start_detect == "2")
                                {
                                    comboBox_list.SelectedIndex = 1;
                                    button_start.PerformClick();
                                }
                                // bonus
                                else if (Properties.Settings.Default.______start_detect == "3")
                                {
                                    comboBox_list.SelectedIndex = 2;
                                    button_start.PerformClick();
                                }
                                // turnover
                                else if (Properties.Settings.Default.______start_detect == "4")
                                {
                                    comboBox_list.SelectedIndex = 3;
                                    button_start.PerformClick();
                                }
                                // bet
                                else if (Properties.Settings.Default.______start_detect == "5")
                                {
                                    comboBox_list.SelectedIndex = 4;
                                    button_start.PerformClick();
                                }
                            }
                            else
                            {
                                label_status.Text = "Waiting";
                            }
                        }
                    }
                    catch (Exception err)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                }
            }
        }

        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox.SelectedIndex == 0)
            {
                // Yesterday
                if (comboBox_list.SelectedIndex == 0)
                {
                    string start = DateTime.Now.ToString("2018-01-22 00:00:00");
                    DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                    dateTimePicker_start.Value = datetime_start;
                    dateTimePicker_start.Visible = false;

                    string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                    DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    dateTimePicker_end.Value = datetime_end;
                }
                else
                {
                    string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                    DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                    dateTimePicker_start.Value = datetime_start;

                    string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                    DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    dateTimePicker_end.Value = datetime_end;
                }
            }
            else if (comboBox.SelectedIndex == 1)
            {
                // Last Week
                DayOfWeek weekStart = DayOfWeek.Sunday;
                DateTime startingDate = DateTime.Today;

                while (startingDate.DayOfWeek != weekStart)
                {
                    startingDate = startingDate.AddDays(-1);
                }

                DateTime datetime_start = startingDate.AddDays(-7);
                dateTimePicker_start.Value = datetime_start;

                string last = startingDate.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(last, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end.Value = datetime_end;
                dateTimePicker_start.Visible = true;
            }
            else if (comboBox.SelectedIndex == 2)
            {
                // Last Month
                var today = DateTime.Today;
                var month = new DateTime(today.Year, today.Month, 1);
                var first = month.AddMonths(-1).ToString("yyyy-MM-dd 00:00:00");
                var last = month.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");

                DateTime datetime_start = DateTime.ParseExact(first, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_start.Value = datetime_start;

                DateTime datetime_end = DateTime.ParseExact(last, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end.Value = datetime_end;
                dateTimePicker_start.Visible = true;
            }
        }

        private void comboBox_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_list.SelectedIndex == 0)
            {
                // Registration
                string start = DateTime.Now.ToString("2018-01-22 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                dateTimePicker_start.Value = datetime_start;
                dateTimePicker_start.Visible = false;

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end.Value = datetime_end;
            }
            else
            {
                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                dateTimePicker_start.Value = datetime_start;
                dateTimePicker_start.Visible = true;

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                dateTimePicker_end.Value = datetime_end;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime today = DateTime.Now;
            DateTime date = today.AddDays(1);
            Properties.Settings.Default.______midnight_time = date.ToString("yyyy-MM-dd 00:30");
            Properties.Settings.Default.______start_detect = "0";
            Properties.Settings.Default.Save();
        }

        private void timer_cycle_in_Tick(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.______midnight_time != "")
            {
                string cyclein_parse = Properties.Settings.Default.______midnight_time;
                DateTime cyclein = DateTime.ParseExact(cyclein_parse, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                string start_parse = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                DateTime start = DateTime.ParseExact(start_parse, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                TimeSpan difference = cyclein - start;
                int hrs = difference.Hours;
                int mins = difference.Minutes;
                int secs = difference.Seconds;

                TimeSpan spinTime = new TimeSpan(hrs, mins, secs);

                TimeSpan delta = DateTime.Now - start;
                TimeSpan timeRemaining = spinTime - delta;

                if (timeRemaining.Hours != 0 && timeRemaining.Minutes != 0)
                {
                    label_cycle_in.Text = timeRemaining.Hours + " hr(s) " + timeRemaining.Minutes + " min(s)";
                }
                else if (timeRemaining.Hours == 0 && timeRemaining.Minutes == 0)
                {
                    label_cycle_in.Text = timeRemaining.Seconds + " sec(s)";
                }
                else if (timeRemaining.Hours == 0)
                {
                    label_cycle_in.Text = timeRemaining.Minutes + " min(s) " + timeRemaining.Seconds + " sec(s)";
                }
            }
            else
            {
                label_cycle_in.Text = "-";
            }
        }

        private void button_start_Click(object sender, EventArgs e)
        {
            __is_start = true;
            panel_filter.Enabled = false;
            label_status.Text = "Waiting";

            string start_datetime = dateTimePicker_start.Text;
            DateTime start = DateTime.Parse(start_datetime);

            string end_datetime = dateTimePicker_end.Text;
            DateTime end = DateTime.Parse(end_datetime);

            string result_start = start.ToString("yyyy-MM-dd");
            string result_end = end.ToString("yyyy-MM-dd");
            string result_start_time = start.ToString("HH:mm:ss");
            string result_end_time = end.ToString("HH:mm:ss");

            if (start <= end)
            {
                button_stop.Visible = true;
                button_start.Visible = false;
                __timer_count = 10;
                label_count.Text = __timer_count.ToString();
                __timer_count = 9;
                label_count.Visible = true;
                timer_start_button.Start();
            }
            else
            {
                MessageBox.Show("No data found.", __brand_code + " Cronos Data", MessageBoxButtons.OK, MessageBoxIcon.Information);
                panel_filter.Enabled = true;
            }
        }

        private void button_stop_Click(object sender, EventArgs e)
        {
            panel_filter.Enabled = true;
            button_stop.Visible = false;
            button_start.Visible = true;
            __timer_count = 10;
            label_count.Visible = false;
            timer_start_button.Stop();
            __is_autostart = false;
            label_status.Text = "Stop";
        }

        private void timer_elapsed_Tick(object sender, EventArgs e)
        {
            string start_datetime = __start_datetime_elapsed;
            DateTime start = DateTime.ParseExact(start_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            string finish_datetime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            DateTime finish = DateTime.ParseExact(finish_datetime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            TimeSpan span = finish.Subtract(start);

            if (span.Hours == 0 && span.Minutes == 0)
            {
                label_elapsed.Text = span.Seconds + " sec(s)";
            }
            else if (span.Hours != 0)
            {
                label_elapsed.Text = span.Hours + " hr(s) " + span.Minutes + " min(s) " + span.Seconds + " sec(s)";
            }
            else if (span.Minutes != 0)
            {
                label_elapsed.Text = span.Minutes + " min(s) " + span.Seconds + " sec(s)";
            }
            else
            {
                label_elapsed.Text = span.Seconds + " sec(s)";
            }
        }

        private async void timer_start_button_TickAsync(object sender, EventArgs e)
        {
            if (__is_login)
            {
                try
                {
                    label_count.Text = __timer_count--.ToString();
                    if (label_count.Text == "9")
                    {
                        label_status.Text = "Running";
                        panel_status.Visible = true;
                        label_start_datetime.Text = DateTime.Now.ToString("ddd, dd MMM HH:mm:ss");
                        __start_datetime_elapsed = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        timer_elapsed.Start();
                        button_stop.Visible = false;
                        label_count.Visible = false;
                        timer_start_button.Stop();

                        if (__conn_id.ToString() == "")
                        {
                            await ___GETCONAsync();
                        }

                        if (comboBox_list.SelectedIndex == 0)
                        {
                            // Registration
                            label_cl_status.Text = "status: doing calculation... --- MEMBER LIST";
                            await ___REGISTRATIONAsync(0);
                        }
                        else if (comboBox_list.SelectedIndex == 1)
                        {
                            // Payment
                            label_cl_status.Text = "status: doing calculation... --- DEPOSIT RECORD";
                            await ___PAYMENT_DEPOSITAsync();
                        }
                        else if (comboBox_list.SelectedIndex == 2)
                        {
                            // Bonus
                            label_cl_status.Text = "status: doing calculation... --- BONUS REPORT";
                            await ___BONUSAsync(0);
                        }
                        else if (comboBox_list.SelectedIndex == 3)
                        {
                            // Turnover Record
                            label_cl_status.Text = "status: doing calculation... --- TURNOVER RECORD";
                            await ___TURNOVERAsync();
                        }
                        else if (comboBox_list.SelectedIndex == 4)
                        {
                            // Bet Record Record
                            label_cl_status.Text = "status: doing calculation... --- BET RECORD";
                            await ___BETAsync(0);
                        }
                    }
                }
                catch (Exception err)
                {
                    // send telegram
                    MessageBox.Show(err.ToString());
                }
            }
        }

        private async Task ___GETCONAsync()
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                var reqparm = new NameValueCollection
                {
                    {"pageIndex", "1"},
                    {"connectionId", "9ca65a15-aa52-4767-b486-60800fb872db"},
                };

                string result = await wc.DownloadStringTaskAsync("http://sn.gk001.gpk456.com/signalr/negotiate");
                var deserialize_object = JsonConvert.DeserializeObject(result);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                __conn_id = _jo.SelectToken("$.ConnectionId");

                __send = 0;
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());
                        
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___GETCONAsync();
                    }
                }
            }
        }

        // REGISTRATION -----
        private async Task ___REGISTRATIONAsync(int index)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                var reqparm = new NameValueCollection
                {
                    {"pageIndex", index.ToString()},
                    {"connectionId", __conn_id.ToString()},
                };
                
                byte[] result = await wc.UploadValuesTaskAsync("http://sn.gk001.gpk456.com/Member/Search", "POST", reqparm);
                string responsebody = Encoding.UTF8.GetString(result).Replace("Date", "TestDate");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_count = _jo.SelectToken("$.PageData");
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- MEMBER LIST";

                // REGISTRATION PROCESS DATA
                char[] split = "*|*".ToCharArray();

                if (_jo_count.Count() > 0)
                {
                    for (int i = 0; i < _jo_count.Count(); i++)
                    {
                        Application.DoEvents();
                        __display_count++;
                        label_total_records.Text = __display_count.ToString("N0");

                        // ----- Username
                        JToken _username = _jo.SelectToken("$.PageData[" + i + "].Account").ToString();
                        // ----- Name
                        JToken _name = _jo.SelectToken("$.PageData[" + i + "].Name").ToString();
                        // ----- VIP
                        JToken _vip = _jo.SelectToken("$.PageData[" + i + "].MemberLevelSettingName").ToString();
                        for (int i_v = 0; i_v < __getdata_viplist.Count; i_v++)
                        {
                            string[] results = __getdata_viplist[i_v].Split(split);
                            if (results[0].Trim() == _vip.ToString())
                            {
                                _vip = results[3].Trim();
                                break;
                            }
                        }
                        if (_vip.ToString() == "")
                        {
                            // notify
                            _vip = "";
                        }
                        // ----- Registration Time
                        JToken _registration_time = _jo.SelectToken("$.PageData[" + i + "].JoinTime").ToString();
                        string _registration_date = "";
                        string _month = "";
                        _registration_time = _registration_time.ToString().Replace("/TestDate(", "");
                        _registration_time = _registration_time.ToString().Replace(")/", "");
                        DateTime _registration_time_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(_registration_time.ToString()) / 1000d)).ToLocalTime();
                        _registration_time = _registration_time_replace.ToString("yyyy/MM/dd HH:mm:ss");
                        _registration_date = _registration_time_replace.ToString("MM/dd/yyyy");
                        _month = _registration_time_replace.ToString("MM/1/yyyy");

                        // ----- Status
                        string _status = _jo.SelectToken("$.PageData[" + i + "].State").ToString();
                        if (_status == "0")
                        {
                            _status = "Inactive";
                        }
                        else
                        {
                            _status = "Active";
                        }
                        // -----
                        string _details = await ___REGISTRATION_DETAILSAsync(_username.ToString());
                        string[] _details_replace = _details.Split('|');
                        string _phone = "";
                        string _email = "";
                        string _wechat = "";
                        string _qq = "";
                        string _last_login_date = "";
                        string _affiliate = "";
                        int _count_details = 0;
                        foreach (string _detail in _details_replace)
                        {
                            _count_details++;

                            if (_count_details == 1)
                            {
                                _phone = _detail;
                            }
                            else if (_count_details == 2)
                            {
                                _email = _detail;
                            }
                            else if (_count_details == 3)
                            {
                                _wechat = _detail;
                            }
                            else if (_count_details == 4)
                            {
                                _qq = _detail;
                            }
                            else if (_count_details == 5)
                            {
                                _last_login_date = _detail;
                            }
                            else if (_count_details == 6)
                            {
                                _affiliate = _detail;
                            }
                        }
                        // ----- Last Deposit Date
                        // ----- First Deposit Data
                        string _fd_ld_date = await ___REGISTRATION_FIRSTLASTDEPOSITAsync(_username.ToString());
                        string[] _fd_ld_date_replace = _fd_ld_date.Split('|');
                        string _fd_date = "";
                        string _ld_date = "";
                        string _fd_date_month = "";
                        int _count_fd_ld = 0;
                        foreach (string _detail in _fd_ld_date_replace)
                        {
                            _count_fd_ld++;

                            if (_count_fd_ld == 1)
                            {
                                _fd_date = _detail;
                            }
                            else if (_count_fd_ld == 2)
                            {
                                _ld_date = _detail;
                            }
                            else if (_count_fd_ld == 3)
                            {
                                _fd_date_month = _detail;
                            }
                        }

                        if (__display_count == 1)
                        {
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", "Brand", "Username", "Name", "Registration Time", "Last Log in Date", "Last Deposit Date", "Status", "Phone", "Email", "Wechat", "QQ", "VIP Level", "Registration Date", "Month", "First Deposit Date", "First Deposit Month", "Affiliate");
                            __DATA.AppendLine(header);
                        }
                        var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16}", __brand_code, "\"" + _username + "\"", "\"" + _name + "\"", "\"" + _registration_time + "\"", "\"" + _last_login_date + "\"", "\"" + _ld_date + "\"", "\"" + _status + "\"", "\"" + _phone + "\"", "\"" + _email + "\"", "\"" + _wechat + "\"", "\"" + _qq + "\"", "\"" + _vip + "\"", "\"" + _registration_date + "\"", "\"" + _month + "\"", "\"" + _fd_date + "\"", "\"" + _fd_date_month + "\"", "\"" + _affiliate + "\"");
                        __DATA.AppendLine(data);
                    }

                    await ___REGISTRATIONAsync(__index_reg++);
                }
                else
                {
                    // REGISTRATION SAVING TO EXCEL
                    string _current_datetime = DateTime.Now.ToString("yyyy-MM-dd");

                    label_cl_status.ForeColor = Color.FromArgb(78, 122, 159);
                    label_cl_status.Text = "status: saving excel... --- MEMBER LIST";
                    label_page_count.Text = "1 of 1";

                    if (!Directory.Exists(__file_location + "\\Cronos Data"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data");
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime);
                    }

                    string _folder_path_result = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\" + __brand_code + " Registration.txt";
                    string _folder_path_result_xlsx = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\" + __brand_code + " Registration.xlsx";

                    if (File.Exists(_folder_path_result))
                    {
                        File.Delete(_folder_path_result);
                    }

                    if (File.Exists(_folder_path_result_xlsx))
                    {
                        File.Delete(_folder_path_result_xlsx);
                    }

                    __DATA.ToString().Reverse();
                    File.WriteAllText(_folder_path_result, __DATA.ToString(), Encoding.UTF8);

                    Excel.Application app = new Excel.Application();
                    Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet worksheet = wb.ActiveSheet;
                    worksheet.Activate();
                    worksheet.Application.ActiveWindow.SplitRow = 1;
                    worksheet.Application.ActiveWindow.FreezePanes = true;
                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                    firstRow.AutoFilter(1,
                                        Type.Missing,
                                        Excel.XlAutoFilterOperator.xlAnd,
                                        Type.Missing,
                                        true);
                    worksheet.Columns[4].NumberFormat = "MM/dd/yyyy HH:mm:ss";
                    Excel.Range usedRange = worksheet.UsedRange;
                    Excel.Range rows = usedRange.Rows;
                    int count = 0;
                    foreach (Excel.Range row in rows)
                    {
                        if (count == 0)
                        {
                            Excel.Range firstCell = row.Cells[1];

                            string firstCellValue = firstCell.Value as string;

                            if (!string.IsNullOrEmpty(firstCellValue))
                            {
                                row.Interior.Color = Color.FromArgb(27, 96, 168);
                                row.Font.Color = Color.FromArgb(255, 255, 255);
                            }

                            break;
                        }

                        count++;
                    }
                    int i_;
                    for (i_ = 1; i_ <= 21; i_++)
                    {
                        worksheet.Columns[i_].ColumnWidth = 20;
                    }
                    wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wb.Close();
                    app.Quit();
                    Marshal.ReleaseComObject(app);

                    if (File.Exists(_folder_path_result))
                    {
                        File.Delete(_folder_path_result);
                    }

                    __DATA.Clear();
                    
                    // REGISTRATION SEND TO DATABASE
                    // AUTO START

                    Properties.Settings.Default.______start_detect = "2";
                    Properties.Settings.Default.Save();

                    panel_status.Visible = false;
                    label_cl_status.Text = "-";
                    label_page_count.Text = "-";
                    label_total_records.Text = "-";
                    button_start.Visible = true;
                    __index_reg = 1;
                    if (__is_autostart)
                    {
                        comboBox_list.SelectedIndex = 1;
                        button_start.PerformClick();
                    }
                    else
                    {
                        panel_filter.Enabled = true;
                    }

                    __send = 0;
                }
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___REGISTRATIONAsync(__index_reg);
                    }
                }
            }
        }

        private async Task<string> ___REGISTRATION_DETAILSAsync(string username)
        {
            var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            wc.Headers["X-Requested-With"] = "XMLHttpRequest";

            var reqparm = new NameValueCollection
            {
                {"account", username},
            };

            byte[] result = await wc.UploadValuesTaskAsync("http://sn.gk001.gpk456.com/Member/GetDetail", "POST", reqparm);
            string responsebody = Encoding.UTF8.GetString(result).Replace("Date", "TestDate");
            var deserialize_object = JsonConvert.DeserializeObject(responsebody);
            JObject _jo = JObject.Parse(deserialize_object.ToString());
            // ----- Phone
            JToken _phone = _jo.SelectToken("$.Member.Mobile").ToString();
            // ----- Phone
            JToken _email = _jo.SelectToken("$.Member.Email").ToString();
            // ----- WeChat
            JToken _wechat = _jo.SelectToken("$.Member.IdNumber").ToString();
            // ----- QQ
            JToken _qq = _jo.SelectToken("$.Member.QQ").ToString();
            // ----- Last Login Date
            JToken _last_login_date_detect = _jo.SelectToken("$.Member.LatestLogin").ToString();
            JToken _last_login_date = "";
            if (_last_login_date_detect.ToString() != "")
            {
                _last_login_date = _jo.SelectToken("$.Member.LatestLogin.Time").ToString();
                _last_login_date = _last_login_date.ToString().Replace("/TestDate(", "");
                _last_login_date = _last_login_date.ToString().Replace(")/", "");
                DateTime _last_login_date_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(_last_login_date.ToString()) / 1000d)).ToLocalTime();
                _last_login_date = _last_login_date_replace.ToString("yyyy-MM-dd");
            }
            // ----- Affiliate
            JToken _jo_affiliate_count = _jo.SelectToken("$.Member.Parents");
            JToken _jo_affiliate = _jo.SelectToken("$.Member.Parents[" + (_jo_affiliate_count.Count()-1) + "].Account");
            char[] split = "*|*".ToCharArray();
            for (int i_a = 0; i_a < __getdata_affiliatelist.Count; i_a++)
            {
                string[] results = __getdata_affiliatelist[i_a].Split(split);
                if (results[0].Trim() == _jo_affiliate.ToString())
                {
                    _jo_affiliate = results[3].Trim();
                    break;
                }
            }
            if (_jo_affiliate.ToString() == "")
            {
                // notify
                _jo_affiliate = "";
            }

            return _phone + "|" + _email + "|" + _wechat + "|" + _qq + "|" + _last_login_date + "|" + _jo_affiliate;
        }

        private async Task<string> ___REGISTRATION_FIRSTLASTDEPOSITAsync(string username)
        {
            var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            wc.Headers["X-Requested-With"] = "XMLHttpRequest";

            var reqparm = new NameValueCollection
            {
                {"Account", username},
                {"AmountBegin", "0"},
                {"IsReal", "true"},
            };

            byte[] result = await wc.UploadValuesTaskAsync("http://sn.gk001.gpk456.com/MemberTransaction/Search", "POST", reqparm);
            string responsebody = Encoding.UTF8.GetString(result).Replace("Date", "TestDate");
            var deserialize_object = JsonConvert.DeserializeObject(responsebody);
            JObject _jo = JObject.Parse(deserialize_object.ToString());
            JToken _jo_count = _jo.SelectToken("$.PageData");
            JToken _fd_date = "";
            JToken _ld_date = "";
            string _month = "";
            if (_jo_count.Count() > 0)
            {
                _ld_date = _jo.SelectToken("$.PageData[0].Time").ToString();
                _ld_date = _ld_date.ToString().Replace("/TestDate(", "");
                _ld_date = _ld_date.ToString().Replace(")/", "");
                DateTime _ld_date_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(_ld_date.ToString()) / 1000d)).ToLocalTime();
                _ld_date = _ld_date_replace.ToString("MM/dd/yyyy");
                
                _fd_date = _jo.SelectToken("$.PageData[" + (_jo_count.Count()-1) + "].Time").ToString();
                _fd_date = _fd_date.ToString().Replace("/TestDate(", "");
                _fd_date = _fd_date.ToString().Replace(")/", "");
                DateTime _fd_date_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(_fd_date.ToString()) / 1000d)).ToLocalTime();
                _fd_date = _fd_date_replace.ToString("MM/dd/yyyy");
                _month = _fd_date_replace.ToString("MM/1/yyyy");
            }
            
            return _fd_date + "|" + _ld_date + "|" + _month;
        }

        // PAYMENT -----
        private async Task ___PAYMENT_DEPOSITAsync()
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                start = datetime_start.ToString("yyyy/MM/dd");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                end = datetime_end.ToString("yyyy/MM/dd");
                
                string responsebody = await wc.UploadStringTaskAsync("http://sn.gk001.gpk456.com/ThirdPartyPayment/LoadNew", "{\"count\":" + __page_size + ",\"minId\":null,\"query\":{\"search\":\"true\",\"ApplyDateBegin\":\"" + start + "\",\"ApplyDateEnd\":\"" + end + "\",\"States\":[3,4,5],\"IsCheckStates\":true,\"isDTPP\":true}}");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_count = _jo.SelectToken("$.Data");
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- DEPOSIT RECORD";

                // PAYMENT DEPOSIT PROCESS DATA
                char[] split = "*|*".ToCharArray();

                // comment comment
                //_jo_count.Count()

                for (int i = 0; i < 10; i++)
                {
                    Application.DoEvents();
                    __display_count++;
                    label_total_records.Text = __display_count.ToString("N0") + " of " + int.Parse(_jo_count.Count().ToString()).ToString("N0");

                    // ----- Username
                    JToken _username = _jo.SelectToken("$.Data[" + i + "].Account").ToString();
                    // ----- Date && Submitted Date
                    JToken _date = _jo.SelectToken("$.Data[" + i + "].Time").ToString().Replace("Date", "TestDate");
                    string _month = "";
                    string _submitted_time = "";
                    string _submitted_date_duration = "";
                    _date = _date.ToString().Replace("/TestDate(", "");
                    _date = _date.ToString().Replace(")/", "");
                    DateTime _date_replace = DateTime.ParseExact(_date.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                    _date = _date_replace.ToString("MM/dd/yyyy");
                    _submitted_time = _date_replace.ToString("hh:mm:ss tt");
                    _submitted_date_duration = _date_replace.ToString("yyyy-MM-dd HH:mm:ss");
                    _month = _date_replace.ToString("MMM-dd");
                    // ----- Transanction Number
                    JToken _transaction_num = _jo.SelectToken("$.Data[" + i + "].Id").ToString();
                    // ----- VIP
                    JToken _vip = _jo.SelectToken("$.Data[" + i + "].MemberLevelName").ToString();
                    for (int i_v = 0; i_v < __getdata_viplist.Count; i_v++)
                    {
                        string[] results = __getdata_viplist[i_v].Split(split);
                        if (results[0].Trim() == _vip.ToString())
                        {
                            _vip = results[3].Trim();
                            break;
                        }
                    }
                    if (_vip.ToString() == "")
                    {
                        // notify
                        _vip = "";
                    }
                    // ----- Amount
                    JToken _amount = _jo.SelectToken("$.Data[" + i + "].Amount").ToString();
                    // ----- Status
                    JToken _status = _jo.SelectToken("$.Data[" + i + "].State").ToString();
                    // ----- Updated Date
                    JToken _updated_time = _jo.SelectToken("$.Data[" + i + "].StateTime").ToString().Replace("Date", "TestDate");
                    string _updated_date_duration = "";
                    if (_updated_time.ToString() != "")
                    {
                        _updated_time = _updated_time.ToString().Replace("/TestDate(", "");
                        _updated_time = _updated_time.ToString().Replace(")/", "");
                        DateTime _updated_time_replace = DateTime.ParseExact(_updated_time.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                        _updated_time = _updated_time_replace.ToString("hh:mm:ss tt");
                        _updated_date_duration = _updated_time_replace.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        _updated_time = "";
                    }
                    // ----- Payment Method && PG Company && PG Type
                    JToken _payment_method = _jo.SelectToken("$.Data[" + i + "].SettingName").ToString();
                    string _pg_company = "";
                    string _pg_type = "";
                    for (int i_p = 0; i_p < __getdata_paymentmethodlist.Count; i_p++)
                    {
                        string[] results = __getdata_paymentmethodlist[i_p].Split(split);
                        if (results[0].Trim() == _payment_method.ToString())
                        {
                            _pg_company = results[3].Trim();
                            _pg_type = results[6].Trim();
                            break;
                        }
                    }
                    if (_pg_company.ToString() == "" || _pg_type.ToString() == "")
                    {
                        // notify
                    }
                    // ----- Duration Time
                    string _duration = "";
                    string _process_duration = "";
                    if (_updated_date_duration.ToString() != "")
                    {
                        DateTime start_date = DateTime.ParseExact(_submitted_date_duration.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        DateTime end_date = DateTime.ParseExact(_updated_date_duration.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        TimeSpan span = end_date - start_date;
                        
                        if (span.Hours.ToString().Length == 1)
                        {
                            _duration += "0" + span.Hours + ":";
                        }
                        else
                        {
                            _duration += span.Hours + ":";
                        }
                        if (span.Minutes.ToString().Length == 1)
                        {
                            _duration += "0" + span.Minutes + ":";
                        }
                        else
                        {
                            _duration += span.Minutes + ":";
                        }
                        if (span.Seconds.ToString().Length == 1)
                        {
                            _duration += "0" + span.Seconds;
                        }
                        else
                        {
                            _duration += span.Seconds;
                        }
                        
                        double totalMinutes = Math.Floor(span.TotalMinutes);
                        if (totalMinutes <= 5)
                        {
                            // 0-5
                            _process_duration = "0-5min";
                        }
                        else if (totalMinutes <= 10)
                        {
                            // 6-10
                            _process_duration = "6-10min";
                        }
                        else if (totalMinutes <= 15)
                        {
                            // 11-15
                            _process_duration = "11-15min";
                        }
                        else if (totalMinutes <= 20)
                        {
                            // 16-20
                            _process_duration = "16-20min";
                        }
                        else if (totalMinutes <= 25)
                        {
                            // 21-25
                            _process_duration = "21-25min";
                        }
                        else if (totalMinutes <= 30)
                        {
                            // 26-30
                            _process_duration = "26-30min";
                        }
                        else if (totalMinutes <= 60)
                        {
                            // 31-60
                            _process_duration = "31-60min";
                        }
                        else if (totalMinutes >= 61)
                        {
                            // >60
                            _process_duration = ">60min";
                        }
                    }
                    // ----- Last Deposit Date
                    // ----- First Deposit Data
                    string _fd_ld_date = await ___REGISTRATION_FIRSTLASTDEPOSITAsync(_username.ToString());
                    string[] _fd_ld_date_replace = _fd_ld_date.Split('|');
                    string _fd_date = "";
                    string _ld_date = "";
                    string _fd_date_month = "";
                    int _count_fd_ld = 0;
                    foreach (string _detail in _fd_ld_date_replace)
                    {
                        _count_fd_ld++;

                        if (_count_fd_ld == 1)
                        {
                            _fd_date = _detail;
                        }
                        else if (_count_fd_ld == 2)
                        {
                            _ld_date = _detail;
                        }
                        else if (_count_fd_ld == 3)
                        {
                            _fd_date_month = _detail;
                        }
                    }
                    // ----- New
                    string _new = "";
                    string _retained = "";
                    string _reactivated = "";
                    if (_status.ToString() == "Success" && !_username.ToString().ToLower().Contains("test"))
                    {
                        if (_fd_date != "" && _ld_date != "")
                        {
                            DateTime _fd_date_ = DateTime.ParseExact(_fd_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                            DateTime _ld_date_ = DateTime.ParseExact(_ld_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                            var _last2months = DateTime.Today.AddMonths(-2);
                            DateTime _last2months_ = DateTime.ParseExact(_last2months.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            if (_ld_date_ >= _last2months_)
                            {
                                _retained = "Retained";
                            }
                            else
                            {
                                _retained = "Not Retained";
                            }

                            string _month_ = DateTime.Now.Month.ToString();
                            string _year_ = DateTime.Now.Year.ToString();
                            string _year_month = _year_ + "-" + _month_;

                            // new
                            if (_fd_date_.ToString("yyyy-MM") == _year_month)
                            {
                                _new = "New";
                            }
                            else
                            {
                                _new = "Not New";
                            }

                            // reactivated
                            if (_retained == "Not Retained" && _new == "Not New")
                            {
                                _reactivated = "Reactivated";
                            }
                            else
                            {
                                _reactivated = "Not Reactivated";
                            }
                        }
                    }
                    else
                    {
                        _fd_date = "";
                    }
                       
                    if (__display_count == 1)
                    {
                        __detect_header = true;
                        var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", "Brand", "Month", "Date", "Transaction #", "Username", "VIP", "Submit Time", "Amount", "Status", "Update Time", "Payment Method", "PG Company", "PG Type", "Duration", "Process Duration", "Transaction Type", "Memo", "FD Date", "New", "Retained", "Reactivated");
                        __DATA.AppendLine(header);
                    }
                    var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", __brand_code, "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _transaction_num + "\"", "\"" + _username + "\"", "\"" + _vip + "\"", "\"" + _submitted_time + "\"", "\"" + _amount + "\"", "\"" + _status + "\"", "\"" + _updated_time + "\"", "\"" + _payment_method + "\"", "\"" + _pg_company + "\"", "\"" + _pg_type + "\"", "\"" + _duration + "\"", "\"" + _process_duration + "\"", "\"" + "Deposit" + "\"", "\"" + "" + "\"", "\"" + _fd_date + "\"", "\"" + _new + "\"", "\"" + _retained + "\"", "\"" + _reactivated + "\"");
                    __DATA.AppendLine(data);
                }

                __display_count = 0;

                await ___PAYMENT_DEPOSITMANUALAsync(0);

                __send = 0;
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___PAYMENT_DEPOSITAsync();
                    }
                }
            }
        }

        private async Task ___PAYMENT_DEPOSITMANUALAsync(int index)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                start = datetime_start.ToString("yyyy/MM/dd 00:00:00");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                end = datetime_end.ToString("yyyy/MM/dd 23:59:59");

                string responsebody = await wc.UploadStringTaskAsync("http://sn.gk001.gpk456.com/MemberTransaction/Search", "{\"TimeBegin\":\"" + start + "\",\"TimeEnd\":\"" + end + "\",\"IsReal\":true,\"Types\":[\"Manual\"],\"pageIndex\":" + index + "}");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_count = _jo.SelectToken("$.PageData");
                JToken _jo_count_ = _jo.SelectToken("$.Total");
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- M-DEPOSIT RECORD";
                
                if (_jo_count.Count() > 0)
                {
                    // PAYMENT DEPOSIT MANUAL PROCESS DATA
                    char[] split = "*|*".ToCharArray();

                    for (int i = 0; i < _jo_count.Count(); i++)
                    {
                        Application.DoEvents();
                        __display_count++;
                        label_total_records.Text = __display_count.ToString("N0") + " of " + int.Parse(_jo_count_.ToString()).ToString("N0");

                        // ----- Username
                        JToken _username = _jo.SelectToken("$.PageData[" + i + "].Account").ToString();
                        // ----- Date && Submitted Date
                        JToken _date = _jo.SelectToken("$.PageData[" + i + "].Time").ToString().Replace("Date", "TestDate");
                        string _month = "";
                        string _submitted_time = "";
                        string _submitted_date_duration = "";
                        _date = _date.ToString().Replace("/TestDate(", "");
                        _date = _date.ToString().Replace(")/", "");
                        DateTime _date_replace = DateTime.ParseExact(_date.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                        _date = _date_replace.ToString("MM/dd/yyyy");
                        _submitted_time = _date_replace.ToString("hh:mm:ss tt");
                        _submitted_date_duration = _date_replace.ToString("yyyy-MM-dd HH:mm:ss");
                        _month = _date_replace.ToString("MMM-dd");
                        // ----- Transanction Number
                        JToken _transaction_num = _jo.SelectToken("$.PageData[" + i + "].Id").ToString();
                        // ----- VIP
                        JToken _vip = _jo.SelectToken("$.PageData[" + i + "].MemberLevelName").ToString();
                        for (int i_v = 0; i_v < __getdata_viplist.Count; i_v++)
                        {
                            string[] results = __getdata_viplist[i_v].Split(split);
                            if (results[0].Trim() == _vip.ToString())
                            {
                                _vip = results[3].Trim();
                                break;
                            }
                        }
                        if (_vip.ToString() == "")
                        {
                            // notify
                            _vip = "";
                        }
                        // ----- Amount
                        JToken _amount = _jo.SelectToken("$.PageData[" + i + "].Amount").ToString();
                        // ----- Status
                        string _status = "Success";
                        // ----- Updated Date
                        JToken _updated_time = _jo.SelectToken("$.PageData[" + i + "].Time").ToString().Replace("Date", "TestDate");
                        string _updated_date_duration = "";
                        if (_updated_time.ToString() != "")
                        {
                            _updated_time = _updated_time.ToString().Replace("/TestDate(", "");
                            _updated_time = _updated_time.ToString().Replace(")/", "");
                            DateTime _updated_time_replace = DateTime.ParseExact(_updated_time.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                            _updated_time = _updated_time_replace.ToString("hh:mm:ss tt");
                            _updated_date_duration = _updated_time_replace.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                        else
                        {
                            _updated_time = "";
                        }
                        // ----- Payment Method && PG Company && PG Type && Memo
                        JToken _memo = _jo.SelectToken("$.PageData[" + i + "].Memo").ToString();
                        _memo = Regex.Replace(_memo.ToString(), @"\t|\n|\r", "");
                        string _payment_method = "Manual Adjustment";
                        string _pg_company = "";
                        string _pg_type = "";
                        if (!_memo.ToString().ToLower().Contains("wechat") || !_memo.ToString().ToLower().Contains("wc") || !_memo.ToString().ToLower().Contains("approve"))
                        {
                            _pg_company = "LOCAL BANK";
                            _pg_type = "LOCAL BANK";
                        }
                        else
                        {
                            _pg_company = "MANUAL WECHAT";
                            _pg_type = "MANUAL WECHAT";
                        }

                        // ----- Duration Time
                        string _duration = "";
                        string _process_duration = "";
                        if (_updated_date_duration.ToString() != "")
                        {
                            DateTime start_date = DateTime.ParseExact(_submitted_date_duration.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            DateTime end_date = DateTime.ParseExact(_updated_date_duration.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            TimeSpan span = end_date - start_date;

                            if (span.Hours.ToString().Length == 1)
                            {
                                _duration += "0" + span.Hours + ":";
                            }
                            else
                            {
                                _duration += span.Hours + ":";
                            }
                            if (span.Minutes.ToString().Length == 1)
                            {
                                _duration += "0" + span.Minutes + ":";
                            }
                            else
                            {
                                _duration += span.Minutes + ":";
                            }
                            if (span.Seconds.ToString().Length == 1)
                            {
                                _duration += "0" + span.Seconds;
                            }
                            else
                            {
                                _duration += span.Seconds;
                            }

                            double totalMinutes = Math.Floor(span.TotalMinutes);
                            if (totalMinutes <= 5)
                            {
                                // 0-5
                                _process_duration = "0-5min";
                            }
                            else if (totalMinutes <= 10)
                            {
                                // 6-10
                                _process_duration = "6-10min";
                            }
                            else if (totalMinutes <= 15)
                            {
                                // 11-15
                                _process_duration = "11-15min";
                            }
                            else if (totalMinutes <= 20)
                            {
                                // 16-20
                                _process_duration = "16-20min";
                            }
                            else if (totalMinutes <= 25)
                            {
                                // 21-25
                                _process_duration = "21-25min";
                            }
                            else if (totalMinutes <= 30)
                            {
                                // 26-30
                                _process_duration = "26-30min";
                            }
                            else if (totalMinutes <= 60)
                            {
                                // 31-60
                                _process_duration = "31-60min";
                            }
                            else if (totalMinutes >= 61)
                            {
                                // >60
                                _process_duration = ">60min";
                            }
                        }
                        // ----- Last Deposit Date
                        // ----- First Deposit PageData
                        string _fd_ld_date = await ___REGISTRATION_FIRSTLASTDEPOSITAsync(_username.ToString());
                        string[] _fd_ld_date_replace = _fd_ld_date.Split('|');
                        string _fd_date = "";
                        string _ld_date = "";
                        string _fd_date_month = "";
                        int _count_fd_ld = 0;
                        foreach (string _detail in _fd_ld_date_replace)
                        {
                            _count_fd_ld++;

                            if (_count_fd_ld == 1)
                            {
                                _fd_date = _detail;
                            }
                            else if (_count_fd_ld == 2)
                            {
                                _ld_date = _detail;
                            }
                            else if (_count_fd_ld == 3)
                            {
                                _fd_date_month = _detail;
                            }
                        }
                        // ----- New
                        string _new = "";
                        string _retained = "";
                        string _reactivated = "";
                        if (_status.ToString() == "Success" && !_username.ToString().ToLower().Contains("test"))
                        {
                            if (_fd_date != "" && _ld_date != "")
                            {
                                DateTime _fd_date_ = DateTime.ParseExact(_fd_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                                DateTime _ld_date_ = DateTime.ParseExact(_ld_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                                var _last2months = DateTime.Today.AddMonths(-2);
                                DateTime _last2months_ = DateTime.ParseExact(_last2months.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                                if (_ld_date_ >= _last2months_)
                                {
                                    _retained = "Retained";
                                }
                                else
                                {
                                    _retained = "Not Retained";
                                }

                                string _month_ = DateTime.Now.Month.ToString();
                                string _year_ = DateTime.Now.Year.ToString();
                                string _year_month = _year_ + "-" + _month_;

                                // new
                                if (_fd_date_.ToString("yyyy-MM") == _year_month)
                                {
                                    _new = "New";
                                }
                                else
                                {
                                    _new = "Not New";
                                }

                                // reactivated
                                if (_retained == "Not Retained" && _new == "Not New")
                                {
                                    _reactivated = "Reactivated";
                                }
                                else
                                {
                                    _reactivated = "Not Reactivated";
                                }
                            }
                        }
                        else
                        {
                            _fd_date = "";
                        }

                        if (__detect_header)
                        {
                            var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", __brand_code, "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _transaction_num + "\"", "\"" + _username + "\"", "\"" + _vip + "\"", "\"" + _submitted_time + "\"", "\"" + _amount + "\"", "\"" + _status + "\"", "\"" + _updated_time + "\"", "\"" + _payment_method + "\"", "\"" + _pg_company + "\"", "\"" + _pg_type + "\"", "\"" + _duration + "\"", "\"" + _process_duration + "\"", "\"" + "Deposit" + "\"", "\"" + _memo + "\"", "\"" + _fd_date + "\"", "\"" + _new + "\"", "\"" + _retained + "\"", "\"" + _reactivated + "\"");
                            __DATA.AppendLine(data);
                        }
                        else
                        {
                            if (__display_count == 1)
                            {
                                __detect_header = true;
                                var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", "Brand", "Month", "Date", "Transaction #", "Username", "VIP", "Submit Time", "Amount", "Status", "Update Time", "Payment Method", "PG Company", "PG Type", "Duration", "Process Duration", "Transaction Type", "Memo", "FD Date", "New", "Retained", "Reactivated");
                                __DATA.AppendLine(header);
                            }
                            var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", __brand_code, "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _transaction_num + "\"", "\"" + _username + "\"", "\"" + _vip + "\"", "\"" + _submitted_time + "\"", "\"" + _amount + "\"", "\"" + _status + "\"", "\"" + _updated_time + "\"", "\"" + _payment_method + "\"", "\"" + _pg_company + "\"", "\"" + _pg_type + "\"", "\"" + _duration + "\"", "\"" + _process_duration + "\"", "\"" + "Deposit" + "\"", "\"" + _memo + "\"", "\"" + _fd_date + "\"", "\"" + _new + "\"", "\"" + _retained + "\"", "\"" + _reactivated + "\"");
                            __DATA.AppendLine(data);
                        }
                    }

                    await ___PAYMENT_DEPOSITMANUALAsync(__index_dep++);
                }
                else
                {
                    __display_count = 0;
                    await ___PAYMENT_WITHDRAWALAsync();
                }

                __send = 0;
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___PAYMENT_DEPOSITMANUALAsync(__index_dep);
                    }
                }
            }
        }

        private async Task ___PAYMENT_WITHDRAWALAsync()
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                start = datetime_start.ToString("yyyy/MM/dd");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                end = datetime_end.ToString("yyyy/MM/dd");
                
                string responsebody = await wc.UploadStringTaskAsync("http://sn.gk001.gpk456.com/VerifyWithdraw/Load", "{\"count\":" + __page_size + ",\"minId\":null,\"query\":{\"search\":\"true\",\"ApplyDateBegin\":\"" + start + "\",\"ApplyDateEnd\":\"" + end + "\"}}");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_count = _jo.SelectToken("$.Data");
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- WITHDRAWAL RECORD";

                // PAYMENT WITHDRAWAL PROCESS DATA
                char[] split = "*|*".ToCharArray();

                for (int i = 0; i < _jo_count.Count(); i++)
                {
                    Application.DoEvents();
                    __display_count++;
                    label_total_records.Text = __display_count.ToString("N0") + " of " + int.Parse(_jo_count.Count().ToString()).ToString("N0");

                    // ----- Username
                    JToken _username = _jo.SelectToken("$.Data[" + i + "].MemberAccount").ToString();
                    // ----- Date && Submitted Date
                    JToken _date = _jo.SelectToken("$.Data[" + i + "].ApplyTime").ToString().Replace("Date", "TestDate");
                    string _month = "";
                    string _submitted_time = "";
                    string _submitted_date_duration = "";
                    _date = _date.ToString().Replace("/TestDate(", "");
                    _date = _date.ToString().Replace(")/", "");
                    DateTime _date_replace = DateTime.ParseExact(_date.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                    _date = _date_replace.ToString("MM/dd/yyyy");
                    _submitted_time = _date_replace.ToString("hh:mm:ss tt");
                    _submitted_date_duration = _date_replace.ToString("yyyy-MM-dd HH:mm:ss");
                    _month = _date_replace.ToString("MMM-dd");
                    // ----- Transanction Number
                    JToken _transaction_num = _jo.SelectToken("$.Data[" + i + "].Id").ToString();
                    // ----- VIP
                    JToken _vip = _jo.SelectToken("$.Data[" + i + "].MemberLevelName").ToString();
                    for (int i_v = 0; i_v < __getdata_viplist.Count; i_v++)
                    {
                        string[] results = __getdata_viplist[i_v].Split(split);
                        if (results[0].Trim() == _vip.ToString())
                        {
                            _vip = results[3].Trim();
                            break;
                        }
                    }
                    if (_vip.ToString() == "")
                    {
                        // notify
                        _vip = "";
                    }
                    // ----- Amount
                    JToken _amount = _jo.SelectToken("$.Data[" + i + "].Amount").ToString();
                    // ----- Status
                    JToken _status = _jo.SelectToken("$.Data[" + i + "].State").ToString();
                    // ----- Updated Date
                    JToken _updated_time = _jo.SelectToken("$.Data[" + i + "].ProcessTime").ToString().Replace("Date", "TestDate");
                    string _updated_date_duration = "";
                    if (_updated_time.ToString() != "")
                    {
                        _updated_time = _updated_time.ToString().Replace("/TestDate(", "");
                        _updated_time = _updated_time.ToString().Replace(")/", "");
                        DateTime _updated_time_replace = DateTime.ParseExact(_updated_time.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                        _updated_time = _updated_time_replace.ToString("hh:mm:ss tt");
                        _updated_date_duration = _updated_time_replace.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        _updated_time = "";
                    }
                    // ----- Payment Method && PG Company && PG Type && Memo
                    JToken _memo = _jo.SelectToken("$.Data[" + i + "].Memo").ToString();
                    _memo = Regex.Replace(_memo.ToString(), @"\t|\n|\r", "");
                    string _payment_method = "Manual Adjustment";
                    string _pg_company = "";
                    string _pg_type = "";
                    if (!_memo.ToString().ToLower().Contains("wechat") || !_memo.ToString().ToLower().Contains("wc") || !_memo.ToString().ToLower().Contains("approve"))
                    {
                        _pg_company = "LOCAL BANK";
                        _pg_type = "LOCAL BANK";
                    }
                    else
                    {
                        _pg_company = "MANUAL WECHAT";
                        _pg_type = "MANUAL WECHAT";
                    }

                    // ----- Duration Time
                    string _duration = "";
                    string _process_duration = "";
                    if (_updated_date_duration.ToString() != "")
                    {
                        DateTime start_date = DateTime.ParseExact(_submitted_date_duration.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        DateTime end_date = DateTime.ParseExact(_updated_date_duration.ToString(), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        TimeSpan span = end_date - start_date;

                        if (span.Hours.ToString().Length == 1)
                        {
                            _duration += "0" + span.Hours + ":";
                        }
                        else
                        {
                            _duration += span.Hours + ":";
                        }
                        if (span.Minutes.ToString().Length == 1)
                        {
                            _duration += "0" + span.Minutes + ":";
                        }
                        else
                        {
                            _duration += span.Minutes + ":";
                        }
                        if (span.Seconds.ToString().Length == 1)
                        {
                            _duration += "0" + span.Seconds;
                        }
                        else
                        {
                            _duration += span.Seconds;
                        }

                        double totalMinutes = Math.Floor(span.TotalMinutes);
                        if (totalMinutes <= 5)
                        {
                            // 0-5
                            _process_duration = "0-5min";
                        }
                        else if (totalMinutes <= 10)
                        {
                            // 6-10
                            _process_duration = "6-10min";
                        }
                        else if (totalMinutes <= 15)
                        {
                            // 11-15
                            _process_duration = "11-15min";
                        }
                        else if (totalMinutes <= 20)
                        {
                            // 16-20
                            _process_duration = "16-20min";
                        }
                        else if (totalMinutes <= 25)
                        {
                            // 21-25
                            _process_duration = "21-25min";
                        }
                        else if (totalMinutes <= 30)
                        {
                            // 26-30
                            _process_duration = "26-30min";
                        }
                        else if (totalMinutes <= 60)
                        {
                            // 31-60
                            _process_duration = "31-60min";
                        }
                        else if (totalMinutes >= 61)
                        {
                            // >60
                            _process_duration = ">60min";
                        }
                    }
                    // ----- Last Deposit Date
                    // ----- First Deposit Data
                    string _fd_ld_date = await ___REGISTRATION_FIRSTLASTDEPOSITAsync(_username.ToString());
                    string[] _fd_ld_date_replace = _fd_ld_date.Split('|');
                    string _fd_date = "";
                    string _ld_date = "";
                    string _fd_date_month = "";
                    int _count_fd_ld = 0;
                    foreach (string _detail in _fd_ld_date_replace)
                    {
                        _count_fd_ld++;

                        if (_count_fd_ld == 1)
                        {
                            _fd_date = _detail;
                        }
                        else if (_count_fd_ld == 2)
                        {
                            _ld_date = _detail;
                        }
                        else if (_count_fd_ld == 3)
                        {
                            _fd_date_month = _detail;
                        }
                    }
                    // ----- New
                    string _new = "";
                    string _retained = "";
                    string _reactivated = "";
                    if (_status.ToString() == "Success" && !_username.ToString().ToLower().Contains("test"))
                    {
                        if (_fd_date != "" && _ld_date != "")
                        {
                            DateTime _fd_date_ = DateTime.ParseExact(_fd_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                            DateTime _ld_date_ = DateTime.ParseExact(_ld_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                            var _last2months = DateTime.Today.AddMonths(-2);
                            DateTime _last2months_ = DateTime.ParseExact(_last2months.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            if (_ld_date_ >= _last2months_)
                            {
                                _retained = "Retained";
                            }
                            else
                            {
                                _retained = "Not Retained";
                            }

                            string _month_ = DateTime.Now.Month.ToString();
                            string _year_ = DateTime.Now.Year.ToString();
                            string _year_month = _year_ + "-" + _month_;

                            // new
                            if (_fd_date_.ToString("yyyy-MM") == _year_month)
                            {
                                _new = "New";
                            }
                            else
                            {
                                _new = "Not New";
                            }

                            // reactivated
                            if (_retained == "Not Retained" && _new == "Not New")
                            {
                                _reactivated = "Reactivated";
                            }
                            else
                            {
                                _reactivated = "Not Reactivated";
                            }
                        }
                    }
                    else
                    {
                        _fd_date = "";
                    }

                    if (__detect_header)
                    {
                        var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", __brand_code, "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _transaction_num + "\"", "\"" + _username + "\"", "\"" + _vip + "\"", "\"" + _submitted_time + "\"", "\"" + _amount + "\"", "\"" + _status + "\"", "\"" + _updated_time + "\"", "\"" + _payment_method + "\"", "\"" + _pg_company + "\"", "\"" + _pg_type + "\"", "\"" + _duration + "\"", "\"" + _process_duration + "\"", "\"" + "Withdrawal" + "\"", "\"" + _memo + "\"", "\"" + _fd_date + "\"", "\"" + _new + "\"", "\"" + _retained + "\"", "\"" + _reactivated + "\"");
                        __DATA.AppendLine(data);
                    }
                    else
                    {
                        if (__display_count == 1)
                        {
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", "Brand", "Month", "Date", "Transaction #", "Username", "VIP", "Submit Time", "Amount", "Status", "Update Time", "Payment Method", "PG Company", "PG Type", "Duration", "Process Duration", "Transaction Type", "Memo", "FD Date", "New", "Retained", "Reactivated");
                            __DATA.AppendLine(header);
                        }
                        var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20}", __brand_code, "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _transaction_num + "\"", "\"" + _username + "\"", "\"" + _vip + "\"", "\"" + _submitted_time + "\"", "\"" + _amount + "\"", "\"" + _status + "\"", "\"" + _updated_time + "\"", "\"" + _payment_method + "\"", "\"" + _pg_company + "\"", "\"" + _pg_type + "\"", "\"" + _duration + "\"", "\"" + _process_duration + "\"", "\"" + "Withdrawal" + "\"", "\"" + _memo + "\"", "\"" + _fd_date + "\"", "\"" + _new + "\"", "\"" + _retained + "\"", "\"" + _reactivated + "\"");
                        __DATA.AppendLine(data);
                    }
                }

                // PAYMENT SAVING TO EXCEL
                string _current_datetime = DateTime.Now.ToString("yyyy-MM-dd");

                label_cl_status.ForeColor = Color.FromArgb(78, 122, 159);
                label_cl_status.Text = "status: saving excel... --- PAYMENT RECORD";

                if (!Directory.Exists(__file_location + "\\Cronos Data"))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data");
                }

                if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code);
                }

                if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime);
                }

                if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Payment Report"))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Payment Report");
                }

                string _folder_path_result = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Payment Report\\" + __brand_code + "_PaymentReport_" + _current_datetime + "_1.txt";
                string _folder_path_result_xlsx = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Payment Report\\" + __brand_code + "_PaymentReport_" + _current_datetime + "_1.xlsx";

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                if (File.Exists(_folder_path_result_xlsx))
                {
                    File.Delete(_folder_path_result_xlsx);
                }

                File.WriteAllText(_folder_path_result, __DATA.ToString(), Encoding.UTF8);

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                worksheet.Columns[2].NumberFormat = "MMM-dd";
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as string;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(27, 96, 168);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                        }

                        break;
                    }

                    count++;
                }
                int i_;
                for (i_ = 1; i_ <= 21; i_++)
                {
                    worksheet.Columns[i_].ColumnWidth = 20;
                }
                wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                __DATA.Clear();
                __detect_header = false;
                __display_count = 0;

                // PAYMENT SEND TO DATABASE
                // AUTO START

                Properties.Settings.Default.______start_detect = "3";
                Properties.Settings.Default.Save();

                panel_status.Visible = false;
                label_cl_status.Text = "-";
                label_page_count.Text = "-";
                label_total_records.Text = "-";
                button_start.Visible = true;
                __index_dep = 1;
                if (__is_autostart)
                {
                    comboBox_list.SelectedIndex = 2;
                    button_start.PerformClick();
                }
                else
                {
                    panel_filter.Enabled = true;
                }

                __send = 0;
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___PAYMENT_WITHDRAWALAsync();
                    }
                }
            }
        }

        // BONUS -----
        private async Task ___BONUSAsync(int index)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                start = datetime_start.ToString("yyyy/MM/dd 00:00:00");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                end = datetime_end.ToString("yyyy/MM/dd 23:59:59");

                string responsebody = await wc.UploadStringTaskAsync("http://sn.gk001.gpk456.com/MemberTransaction/Search", "{\"TimeBegin\":\"" + start + "\",\"TimeEnd\":\"" + end + "\",\"IsReal\":\"false\",\"Types\":[\"Manual\",\"Bonus\",\"Other\",\"AnyTimeDiscount\",\"Discount\"],\"pageIndex\":" + index + "}");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_count = _jo.SelectToken("$.PageData");
                JToken _jo_count_ = _jo.SelectToken("$.Total");
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- BONUS RECORD";
                
                if (_jo_count.Count() > 0)
                {
                    // BONUS PROCESS DATA
                    char[] split = "*|*".ToCharArray();
                    
                    for (int i = 0; i < _jo_count.Count(); i++)
                    {
                        Application.DoEvents();
                        __display_count++;
                        label_total_records.Text = __display_count.ToString("N0") + " of " + int.Parse(_jo_count_.ToString()).ToString("N0");

                        // ----- Username
                        JToken _username = _jo.SelectToken("$.PageData[" + i + "].Account").ToString();
                        // ----- Date && Submitted Date
                        JToken _date = _jo.SelectToken("$.PageData[" + i + "].Time").ToString().Replace("Date", "TestDate");
                        string _month = "";
                        string _transaction_time = "";
                        string _submitted_date_duration = "";
                        _date = _date.ToString().Replace("/TestDate(", "");
                        _date = _date.ToString().Replace(")/", "");
                        DateTime _date_replace = DateTime.ParseExact(_date.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                        _date = _date_replace.ToString("MM/dd/yyyy");
                        _transaction_time = _date_replace.ToString("hh:mm:ss tt");
                        _month = _date_replace.ToString("MMM-dd");
                        // ----- Memo && Bonus Category && Purpose
                        string _purpose = "";
                        string _bonus_category = "";
                        JToken _remark = _jo.SelectToken("$.PageData[" + i + "].Memo").ToString();
                        if (_remark.ToString().ToLower().Contains(";"))
                        {
                            try
                            {
                                if (!_username.ToString().ToLower().Contains("test"))
                                {
                                    string _bonus_code_replace_ = "";
                                    foreach (char c in _remark.ToString())
                                    {
                                        if (c == ';')
                                        {
                                            break;
                                        }
                                        else
                                        {
                                            if (c != ' ')
                                            {
                                                _bonus_code_replace_ += c;
                                            }
                                        }
                                    }

                                    for (int i_b = 0; i_b < __getdata_bonuscode.Count; i_b++)
                                    {
                                        string[] results = __getdata_bonuscode[i_b].Split(split);
                                        if (results[0].Trim() == _bonus_code_replace_)
                                        {
                                            _remark = results[3].Trim();
                                            if (_purpose.ToString() == "")
                                            {
                                                _purpose = results[6].Trim();
                                            }
                                            break;
                                        }
                                    }

                                    if (_bonus_category == "" && _purpose == "")
                                    {
                                        string get1 = _remark.ToString().Substring(6, 3);

                                        string get2 = get1.Substring(0, 2);
                                        string get3 = get1.Substring(2);
                                        string get4 = get1.Substring(0, 2);

                                        ArrayList items_code = new ArrayList(new string[] { "AD", "FD", "FDB", "DP", "PZ", "RF", "RAF", "RB", "SU", "TO", "RR", "CB", "GW", "RW", "TE" });
                                        ArrayList items_bonus_category = new ArrayList(new string[] { "Adjustment", "FDB", "FDB", "Deposit", "Prize", "Refer friend", "Refer friend", "Reload", "Signup Bonus", "Turnover", "Rebate", "Cashback", "Goodwill Bonus", "Reward", "Test" });
                                        int count_ = 0;
                                        foreach (var item in items_code)
                                        {
                                            if (get2 == item.ToString())
                                            {
                                                _bonus_category = items_bonus_category[count_].ToString();
                                                break;
                                            }

                                            count_++;
                                        }

                                        if (_purpose.ToString() == "")
                                        {
                                            if (get3 == "0")
                                            {
                                                if (get4 == "FD" || get4 == "RA")
                                                {
                                                    get1 = _remark.ToString().Substring(6, 5);
                                                    get2 = get1.Substring(0, 4);
                                                    get3 = get1.Substring(4);
                                                }
                                                else
                                                {
                                                    get1 = _remark.ToString().Substring(6, 4);
                                                    get2 = get1.Substring(0, 3);
                                                    get3 = get1.Substring(3);
                                                }
                                            }

                                            ArrayList items_code_ = new ArrayList(new string[] { "0", "1", "2", "3", "4" });
                                            ArrayList items_bonus_category_ = new ArrayList(new string[] { "Retention", "Acquisition", "Conversion", "Retention", "Reactivation" });
                                            int count__ = 0;
                                            foreach (var item in items_code_)
                                            {
                                                if (get3 == item.ToString())
                                                {
                                                    _purpose = items_bonus_category_[count__].ToString();
                                                    break;
                                                }

                                                count__++;
                                            }
                                        }

                                        if (_bonus_category == "" && _purpose.ToString() == "")
                                        {
                                            _bonus_category = "Rebate";
                                            _purpose = "Retention";
                                        }
                                    }
                                }
                                else
                                {
                                    _bonus_category = "Other";
                                    _purpose = "Adjustment";
                                }
                            }
                            catch (Exception err)
                            {
                                _bonus_category = "Other";
                                _purpose = "Adjustment";

                                SendMyBot(err.ToString());
                            }
                        }
                        else
                        {
                            _remark = ___ToStringWithOrdinal(DateTime.Now.AddDays(-2));
                            _bonus_category = "Rebate";
                            _purpose = "Retention (Rebate)";
                        }
                        // ----- Transanction Number
                        JToken _transaction_num = _jo.SelectToken("$.PageData[" + i + "].Id").ToString();
                        // ----- Type
                        JToken _type = _jo.SelectToken("$.PageData[" + i + "].TypeString").ToString();
                        // ----- 
                        JToken _member_id = _jo.SelectToken("$.PageData[" + i + "].MemberId").ToString();
                        string _operator__processing_staff = await ___BONUS_DETAILSAsync(_member_id.ToString());
                        // ----- VIP
                        JToken _vip = _jo.SelectToken("$.PageData[" + i + "].MemberLevelName").ToString();
                        for (int i_v = 0; i_v < __getdata_viplist.Count; i_v++)
                        {
                            string[] results = __getdata_viplist[i_v].Split(split);
                            if (results[0].Trim() == _vip.ToString())
                            {
                                _vip = results[3].Trim();
                                break;
                            }
                        }
                        if (_vip.ToString() == "")
                        {
                            // notify
                            _vip = "";
                        }
                        // ----- Amount
                        JToken _amount = _jo.SelectToken("$.PageData[" + i + "].Amount").ToString();

                        if (__display_count == 1)
                        {
                            __detect_header = true;
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}", "Brand", "Month", "Date", "Transaction #", "Username", "VIP", "Transaction Time", "Type", "Amount", "Remark", "Bonus Category", "Purpose", "Operator", "Processing Staff");
                            __DATA.AppendLine(header);
                        }
                        var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13}", __brand_code, "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _transaction_num + "\"", "\"" + _username + "\"", "\"" + _vip + "\"", "\"" + _transaction_time + "\"", "\"" + _type + "\"", "\"" + _amount + "\"", "\"" + _remark + "\"", "\"" + _bonus_category + "\"", "\"" + _purpose + "\"", "\"" + _operator__processing_staff + "\"", "\"" + _operator__processing_staff + "\"");
                        __DATA.AppendLine(data);
                    }

                    await ___BONUSAsync(__index_bon++);
                }
                else
                {
                    // BONUS SAVING EXCEL
                    string _current_datetime = DateTime.Now.ToString("yyyy-MM-dd");

                    label_cl_status.ForeColor = Color.FromArgb(78, 122, 159);
                    label_cl_status.Text = "status: saving excel... --- BONUS RECORD";

                    if (!Directory.Exists(__file_location + "\\Cronos Data"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data");
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bonus Report"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bonus Report");
                    }

                    string _folder_path_result = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bonus Report\\" + __brand_code + "_BonusReport_" + _current_datetime + "_1.txt";
                    string _folder_path_result_xlsx = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bonus Report\\" + __brand_code + "_BonusReport_" + _current_datetime + "_1.xlsx";

                    if (File.Exists(_folder_path_result))
                    {
                        File.Delete(_folder_path_result);
                    }

                    if (File.Exists(_folder_path_result_xlsx))
                    {
                        File.Delete(_folder_path_result_xlsx);
                    }

                    File.WriteAllText(_folder_path_result, __DATA.ToString(), Encoding.UTF8);

                    Excel.Application app = new Excel.Application();
                    Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet worksheet = wb.ActiveSheet;
                    worksheet.Activate();
                    worksheet.Application.ActiveWindow.SplitRow = 1;
                    worksheet.Application.ActiveWindow.FreezePanes = true;
                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                    firstRow.AutoFilter(1,
                                        Type.Missing,
                                        Excel.XlAutoFilterOperator.xlAnd,
                                        Type.Missing,
                                        true);
                    worksheet.Columns[2].NumberFormat = "MMM-dd";
                    Excel.Range usedRange = worksheet.UsedRange;
                    Excel.Range rows = usedRange.Rows;
                    int count = 0;
                    foreach (Excel.Range row in rows)
                    {
                        if (count == 0)
                        {
                            Excel.Range firstCell = row.Cells[1];

                            string firstCellValue = firstCell.Value as string;

                            if (!string.IsNullOrEmpty(firstCellValue))
                            {
                                row.Interior.Color = Color.FromArgb(27, 96, 168);
                                row.Font.Color = Color.FromArgb(255, 255, 255);
                            }

                            break;
                        }

                        count++;
                    }
                    int i_;
                    for (i_ = 1; i_ <= 21; i_++)
                    {
                        worksheet.Columns[i_].ColumnWidth = 20;
                    }
                    wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wb.Close();
                    app.Quit();
                    Marshal.ReleaseComObject(app);

                    if (File.Exists(_folder_path_result))
                    {
                        File.Delete(_folder_path_result);
                    }

                    __DATA.Clear();
                    __detect_header = false;

                    // BONUS SEND TO DATABASE
                    // AUTO START

                    Properties.Settings.Default.______start_detect = "4";
                    Properties.Settings.Default.Save();

                    panel_status.Visible = false;
                    label_cl_status.Text = "-";
                    label_page_count.Text = "-";
                    label_total_records.Text = "-";
                    button_start.Visible = true;
                    if (__is_autostart)
                    {
                        comboBox_list.SelectedIndex = 3;
                        button_start.PerformClick();
                    }
                    else
                    {
                        panel_filter.Enabled = true;
                    }

                    __send = 0;
                }

                __send = 0;
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___BONUSAsync(__index_bon);
                    }
                }
            }
        }

        private async Task<string> ___BONUS_DETAILSAsync(string id)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                start = datetime_start.ToString("yyyy/MM/dd 00:00:00");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                end = datetime_end.ToString("yyyy/MM/dd 23:59:59");

                string responsebody = await wc.UploadStringTaskAsync("http://sn.gk001.gpk456.com/MemberTransaction/GetDetail", "{\"id\":\"" + id + "\"}");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _operator_name = _jo.SelectToken("$.Detail.CreateName").ToString();
                
                return _operator_name.ToString();
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___BONUS_DETAILSAsync(id);
                    }
                }
                
                return "";
            }
        }

        private string ___ToStringWithOrdinal(DateTime d)
        {
            switch (d.Day)
            {
                case 1:
                case 21:
                case 31:
                    return "Rebate for" + d.ToString(" dd'st' MMMM yyyy");
                case 2:
                case 22:
                    return "Rebate for" + d.ToString(" dd'nd' MMMM yyyy");
                case 3:
                case 23:
                    return "Rebate for" + d.ToString(" dd'rd' MMMM yyyy");
                default:
                    return "Rebate for" + d.ToString(" dd'th' MMMM yyyy");
            }
        }

        // TURNOVER ----- 
        private async Task ___TURNOVERAsync()
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                start = start.Replace("-", "%2F");
                start = start.Replace(" ", "+");
                start = start.Replace(":", "%3A");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                end = end.Replace("-", "%2F");
                end = end.Replace(" ", "+");
                end = end.Replace(":", "%3A");

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                string responsebody = await wc.DownloadStringTaskAsync("http://sn.gk001.gpk456.com/Statistics/Export?begin=" + start + "&end=" + end + "");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_success = _jo.SelectToken("$.IsSuccess");
                if (_jo_success.ToString() == "True")
                {
                    JToken _jo_path = _jo.SelectToken("$.FileVirtualPath").ToString();
                    string _current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
                    if (!Directory.Exists(__file_location + "\\Cronos Data"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data");
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record");
                    }

                    string _folder_path_result_xlsx = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record\\" + __brand_code + "_TurnoverRecord_" + _current_datetime + "_1.xlsx";

                    if (File.Exists(_folder_path_result_xlsx))
                    {
                        File.Delete(_folder_path_result_xlsx); 
                    }

                    await ___TURNOVER_DOWNLOADAync(__root_url + _jo_path + "?", _folder_path_result_xlsx, cookie);
                    
                    ___TURNOVER_PROCESSAsync(_folder_path_result_xlsx);
                }
                else
                {
                    if (__is_login)
                    {
                        __send++;
                        if (__send == 5)
                        {
                            //SendMyBot("Can't download Turnover Record at this moment.");
                            SendReportsTeam("Can't download Turnover Record at this moment.");

                            Properties.Settings.Default.______start_detect = "5";
                            Properties.Settings.Default.Save();

                            panel_status.Visible = false;
                            label_cl_status.Text = "-";
                            label_page_count.Text = "-";
                            label_total_records.Text = "-";
                            button_start.Visible = true;
                            __index_reg = 1;
                            if (__is_autostart)
                            {
                                comboBox_list.SelectedIndex = 4;
                                button_start.PerformClick();
                            }
                            else
                            {
                                panel_filter.Enabled = true;
                            }

                            __DATA.Clear();
                            __detect_header = false;
                            __send = 0;
                        }
                        else
                        {
                            ___WaitNSeconds(2);
                            await ___TURNOVERAsync();
                        }
                    }
                }

            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        //SendMyBot("Can't download Turnover Record at this moment.");
                        SendReportsTeam("Can't download Turnover Record at this moment.");
                        SendMyBot(err.ToString());
                        
                        Properties.Settings.Default.______start_detect = "5";
                        Properties.Settings.Default.Save();

                        panel_status.Visible = false;
                        label_cl_status.Text = "-";
                        label_page_count.Text = "-";
                        label_total_records.Text = "-";
                        button_start.Visible = true;
                        __index_reg = 1;
                        if (__is_autostart)
                        {
                            comboBox_list.SelectedIndex = 4;
                            button_start.PerformClick();
                        }
                        else
                        {
                            panel_filter.Enabled = true;
                        }

                        __DATA.Clear();
                        __detect_header = false;
                        __send = 0;
                    }
                    else
                    {
                        ___WaitNSeconds(2);
                        await ___TURNOVERAsync();
                    }
                }
            }
        }

        private async Task ___TURNOVER_DOWNLOADAync(string path_from, string path_to, string cookie)
        {
            WebClient wc = new WebClient();
            wc.Headers.Add("User-Agent", "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.2; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 1.1.4322; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)");
            wc.Headers.Add("Accept", "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*");
            wc.Headers.Add("Accept-Language", "en-US,en;q=0.55");
            wc.Encoding = Encoding.UTF8;
            //wc.Headers.Add("Content-Type", "image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*");
            wc.Headers.Add("Cookie", cookie);

            await wc.DownloadFileTaskAsync(
                new Uri(path_from),
                path_to
            );
        }

        private async void ___TURNOVER_PROCESSAsync(string path)
        {
            try
            {
                Excel.Application app_ = new Excel.Application();
                Excel.Workbook workbook_ = app_.Workbooks.Open(path);
                Excel._Worksheet worksheet_ = workbook_.Sheets[2];
                Excel.Range range_ = worksheet_.UsedRange;

                int rowCount = range_.Rows.Count;
                int colCount = range_.Columns.Count;

                __DATA.Clear();
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- TURNOVER RECORD";

                for (int i = 1; i <= rowCount; i++)
                {
                    if (i != 1)
                    {
                        Application.DoEvents();

                        int count_ = 0;
                        string _member = "";
                        string _product = "";
                        string _bet_count = "";
                        string _stake = "";
                        string _stake_ex_draw = "";
                        string _company_wl = "";
                        string _provider = "";
                        string _category = "";
                        string _vip = "";
                        string _date = DateTime.Now.AddDays(-1).ToString("MM/dd/yyyy");
                        string _month = DateTime.Now.AddDays(-1).ToString("MMM-dd");
                        string _reg_month = "";
                        string _fd_date = "";
                        string _ld_date = "";
                        char[] split = "*|*".ToCharArray();
                        
                        __display_count++;
                        label_total_records.Text = __display_count.ToString("N0");

                        for (int j = 1; j <= colCount; j++)
                        {
                            count_++;
                            if (count_ == 5)
                            {
                                _member = range_.Cells[i, j].Value2.ToString();
                                // Datails
                                string _details = await ___TURNOVER_VIP_REGMONTHAsync(_member);
                                string[] _details_replace = _details.Split('|');
                                int _count_details = 0;
                                foreach (string _detail in _details_replace)
                                {
                                    _count_details++;

                                    if (_count_details == 1)
                                    {
                                        _vip = _detail;
                                    }
                                    else if (_count_details == 2)
                                    {
                                        _reg_month = _detail;
                                    }
                                }

                                // FD LD Date
                                string _fd_ld_date = await ___REGISTRATION_FIRSTLASTDEPOSITAsync(_member.ToString());
                                string[] _fd_ld_date_replace = _fd_ld_date.Split('|');
                                int _count_fd_ld = 0;
                                foreach (string _detail in _fd_ld_date_replace)
                                {
                                    _count_fd_ld++;

                                    if (_count_fd_ld == 1)
                                    {
                                        _fd_date = _detail;
                                    }
                                    else if (_count_fd_ld == 2)
                                    {
                                        _ld_date = _detail;
                                    }
                                }
                            }
                            else if (count_ == 6)
                            {
                                _product = range_.Cells[i, j].Value2.ToString();
                                for (int i_p = 0; i_p < __getdata_productcode.Count; i_p++)
                                {
                                    string[] results = __getdata_productcode[i_p].Split(split);
                                    if (results[0].Trim() == _product)
                                    {
                                        _provider = results[3].Trim();
                                        _category = results[6].Trim();
                                        break;
                                    }
                                }
                            }
                            else if (count_ == 7)
                            {
                                _bet_count = range_.Cells[i, j].Value2.ToString();
                            }
                            else if (count_ == 8)
                            {
                                _stake = range_.Cells[i, j].Value2.ToString();
                            }
                            else if (count_ == 9)
                            {
                                _stake_ex_draw = range_.Cells[i, j].Value2.ToString();
                            }
                            else if (count_ == 10)
                            {
                                _company_wl = range_.Cells[i, j].Value2.ToString();
                            }
                        }

                        // ----- Retained
                        string _retained = "";
                        string _month_ = DateTime.Now.Month.ToString();
                        string _year_ = DateTime.Now.Year.ToString();
                        string _year_month = _year_ + "-" + _month_;
                        if (_fd_date != "" && _ld_date != "")
                        {
                            DateTime _fd_date_ = DateTime.ParseExact(_fd_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                            DateTime _ld_date_ = DateTime.ParseExact(_ld_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                            var _last2months = DateTime.Today.AddMonths(-2);
                            DateTime _last2months_ = DateTime.ParseExact(_last2months.ToString("yyyy-MM-dd"), "yyyy-MM-dd", CultureInfo.InvariantCulture);
                            if (_ld_date_ >= _last2months_)
                            {
                                _retained = "Yes";
                            }
                            else
                            {
                                _retained = "No";
                            }
                        }
                        else
                        {
                            _retained = "No";
                        }
                        // ----- New Based on Reg && Reg Month
                        string _new_based_on_reg = "";
                        if (_reg_month != "")
                        {
                            DateTime _reg_month_replace_ = DateTime.ParseExact(_reg_month, "dd-MMM-y", CultureInfo.InvariantCulture);
                            if (_reg_month_replace_.ToString("yyyy-MM") == _year_month)
                            {
                                _new_based_on_reg = "Yes";
                            }
                            else
                            {
                                _new_based_on_reg = "No";
                            }
                        }
                        else
                        {
                            _new_based_on_reg = "No";
                            _reg_month = "";
                        }
                        // ----- New Based on Dep
                        // ----- Real Player
                        string _real_player = "";
                        string _new_based_on_dep = "";
                        if (_fd_date != "")
                        {
                            DateTime _first_deposit = DateTime.ParseExact(_fd_date, "MM/dd/yyyy", CultureInfo.InvariantCulture);
                            if (_first_deposit.ToString("yyyy-MM") == _year_month)
                            {
                                _new_based_on_dep = "Yes";
                            }
                            else
                            {
                                _new_based_on_dep = "No";
                            }

                            _real_player = "Yes";
                        }
                        else
                        {
                            _new_based_on_dep = "No";
                            _real_player = "No";
                        }

                        if (__display_count == 1)
                        {
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}", "Brand", "Provider", "Category", "Month", "Date", "Member", "Currency", "Stake", "Stake Ex. Draw", "Bet Count", "Company Win Loss", "VIP", "Retained", "Reg Month", "First Dep Month", "New Based on Reg", "New Based on Dep", "Real Player");
                            __DATA.AppendLine(header);
                        }
                        var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17}", __brand_code, "\"" + _provider + "\"", "\"" + _category + "\"", "\"" + _month + "\"", "\"" + _date + "\"", "\"" + _member + "\"", "\"" + "CNY" + "\"", "\"" + _stake + "\"", "\"" + _stake_ex_draw + "\"", "\"" + _bet_count + "\"", "\"" + _company_wl + "\"", "\"" + _vip + "\"", "\"" + _retained + "\"", "\"" + _reg_month + "\"", "\"" + _fd_date + "\"", "\"" + _new_based_on_reg + "\"", "\"" + _new_based_on_dep + "\"", "\"" + _real_player + "\"");
                        __DATA.AppendLine(data);
                    }
                }

                workbook_.Close();
                app_.Quit();
                Marshal.ReleaseComObject(app_);

                // TURNOVER SAVING EXCEL
                string _current_datetime = DateTime.Now.ToString("yyyy-MM-dd");
                if (!Directory.Exists(__file_location + "\\Cronos Data"))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data");
                }

                if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code);
                }

                if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime);
                }

                if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record"))
                {
                    Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record");
                }

                string _folder_path_result = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record\\" + __brand_code + "_TurnoverRecord_" + _current_datetime + "_1.txt";
                string _folder_path_result_xlsx = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Turnover Record\\" + __brand_code + "_TurnoverRecord_" + _current_datetime + "_1.xlsx";

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                if (File.Exists(_folder_path_result_xlsx))
                {
                    File.Delete(_folder_path_result_xlsx);
                }

                File.WriteAllText(_folder_path_result, __DATA.ToString(), Encoding.UTF8);

                Excel.Application app = new Excel.Application();
                Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = wb.ActiveSheet;
                worksheet.Activate();
                worksheet.Application.ActiveWindow.SplitRow = 1;
                worksheet.Application.ActiveWindow.FreezePanes = true;
                Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                firstRow.AutoFilter(1,
                                    Type.Missing,
                                    Excel.XlAutoFilterOperator.xlAnd,
                                    Type.Missing,
                                    true);
                worksheet.Columns[4].NumberFormat = "MMM-dd";
                worksheet.Columns[14].NumberFormat = "dd-MMM-y";
                worksheet.Columns[15].NumberFormat = "dd-MMM-y";
                Excel.Range usedRange = worksheet.UsedRange;
                Excel.Range rows = usedRange.Rows;
                int count = 0;
                foreach (Excel.Range row in rows)
                {
                    if (count == 0)
                    {
                        Excel.Range firstCell = row.Cells[1];

                        string firstCellValue = firstCell.Value as string;

                        if (!string.IsNullOrEmpty(firstCellValue))
                        {
                            row.Interior.Color = Color.FromArgb(27, 96, 168);
                            row.Font.Color = Color.FromArgb(255, 255, 255);
                        }

                        break;
                    }

                    count++;
                }
                int i_;
                for (i_ = 1; i_ <= 21; i_++)
                {
                    worksheet.Columns[i_].ColumnWidth = 20;
                }
                wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wb.Close();
                app.Quit();
                Marshal.ReleaseComObject(app);

                if (File.Exists(_folder_path_result))
                {
                    File.Delete(_folder_path_result);
                }

                // TURNOVER SEND TO DATABASE
                // AUTO START

                Properties.Settings.Default.______start_detect = "5";
                Properties.Settings.Default.Save();

                panel_status.Visible = false;
                label_cl_status.Text = "-";
                label_page_count.Text = "-";
                label_total_records.Text = "-";
                button_start.Visible = true;
                __index_reg = 1;
                if (__is_autostart)
                {
                    comboBox_list.SelectedIndex = 4;
                    button_start.PerformClick();
                }
                else
                {
                    panel_filter.Enabled = true;
                }

                __DATA.Clear();
                __detect_header = false;
                __display_count = 0;
                __send = 0;
            }
            catch (Exception err)
            {
                ___WaitNSeconds(5);
                __DATA.Clear();
                __display_count = 0;
                ___TURNOVER_PROCESSAsync(path);
            }
        }

        private async Task<string> ___TURNOVER_VIP_REGMONTHAsync(string username)
        {
            var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
            WebClient wc = new WebClient();

            wc.Headers.Add("Cookie", cookie);
            wc.Encoding = Encoding.UTF8;
            wc.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            wc.Headers["X-Requested-With"] = "XMLHttpRequest";

            var reqparm = new NameValueCollection
            {
                {"Account", username},
                {"connectionId", __conn_id.ToString()},
            };

            byte[] result = await wc.UploadValuesTaskAsync("http://sn.gk001.gpk456.com/Member/Search", "POST", reqparm);
            string responsebody = Encoding.UTF8.GetString(result).Replace("Date", "TestDate");
            var deserialize_object = JsonConvert.DeserializeObject(responsebody);
            JObject _jo = JObject.Parse(deserialize_object.ToString());
            // ----- VIP
            JToken _vip = _jo.SelectToken("$.PageData[0].MemberLevelSettingName").ToString();
            char[] split = "*|*".ToCharArray();
            for (int i_v = 0; i_v < __getdata_viplist.Count; i_v++)
            {
                string[] results = __getdata_viplist[i_v].Split(split);
                if (results[0].Trim() == _vip.ToString())
                {
                    _vip = results[3].Trim();
                    break;
                }
            }
            if (_vip.ToString() == "")
            {
                // notify
                _vip = "";
            }
            // ----- Registration Time
            JToken _reg_month = _jo.SelectToken("$.PageData[0].JoinTime").ToString();
            _reg_month = _reg_month.ToString().Replace("/TestDate(", "");
            _reg_month = _reg_month.ToString().Replace(")/", "");
            DateTime _date_replace = new DateTime(1970, 1, 1, 0, 0, 0, 0).AddSeconds(Math.Round(Convert.ToDouble(_reg_month.ToString()) / 1000d)).ToLocalTime();
            _reg_month = _date_replace.ToString("dd-MMM-y");
            
            return _vip + "|" + _reg_month;
        }
        
        // BET -----
        private async Task ___BETAsync(int index)
        {
            try
            {
                var cookie = Cookie.GetCookieInternal(webBrowser.Url, false);
                WebClient wc = new WebClient();

                wc.Headers.Add("Cookie", cookie);
                wc.Encoding = Encoding.UTF8;
                wc.Headers[HttpRequestHeader.ContentType] = "application/json";
                wc.Headers["X-Requested-With"] = "XMLHttpRequest";

                string start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_start = DateTime.ParseExact(start, "yyyy-MM-dd 00:00:00", CultureInfo.InvariantCulture);
                start = datetime_start.ToString("yyyy/MM/dd 00:00:00");

                string end = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
                DateTime datetime_end = DateTime.ParseExact(end, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                end = datetime_end.ToString("yyyy/MM/dd 23:59:59");

                string responsebody = await wc.UploadStringTaskAsync("http://sn.gk001.gpk456.com/BetRecord/Search", "{\"WagersTimeBegin\":\"" + start + "\",\"WagersTimeEnd\":\"" + end + "\",\"connectionId\":\"" + __conn_id + "\", \"pageIndex\": " + index + "}");
                var deserialize_object = JsonConvert.DeserializeObject(responsebody);
                JObject _jo = JObject.Parse(deserialize_object.ToString());
                JToken _jo_count = _jo.SelectToken("$.PageData");
                label_page_count.Text = "1 of 1";
                label_cl_status.Text = "status: getting data... --- BET RECORD";
                
                if (_jo_count.Count() > 0)
                {
                    // BET RECORD PROCESS DATA
                    char[] split = "*|*".ToCharArray();
                    
                    for (int i = 0; i < _jo_count.Count(); i++)
                    {
                        Application.DoEvents();
                        __display_count++;
                        label_total_records.Text = __display_count.ToString("N0");

                        // ----- Username
                        JToken _username = _jo.SelectToken("$.PageData[" + i + "].Account").ToString();
                        // ----- Provider
                        JToken _provider = _jo.SelectToken("$.PageData[" + i + "].GameCategory").ToString();
                        // ----- Games
                        JToken _games = _jo.SelectToken("$.PageData[" + i + "].GameType").ToString();
                        // ----- Bet Time
                        JToken _bet_time = _jo.SelectToken("$.PageData[" + i + "].WagersTime").ToString();
                        string _bet_date = "";
                        if (_bet_time.ToString() != "")
                        {
                            _bet_time = _bet_time.ToString().Replace("/TestDate(", "");
                            _bet_time = _bet_time.ToString().Replace(")/", "");
                            DateTime _bet_time_replace = DateTime.ParseExact(_bet_time.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                            _bet_time = _bet_time_replace.ToString("yyyy/MM/dd HH:mm:ss");
                            _bet_date = _bet_time_replace.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            _bet_time = "";
                        }
                        // ----- Win Loss Time
                        JToken _win_loss_time = _jo.SelectToken("$.PageData[" + i + "].PayoffTime").ToString();
                        string _win_loss_date = "";
                        if (_win_loss_time.ToString() != "")
                        {
                            _win_loss_time = _win_loss_time.ToString().Replace("/TestDate(", "");
                            _win_loss_time = _win_loss_time.ToString().Replace(")/", "");
                            DateTime _win_loss_time_replace = DateTime.ParseExact(_win_loss_time.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture).AddHours(20);
                            _win_loss_time = _win_loss_time_replace.ToString("yyyy/MM/dd HH:mm:ss");
                            _win_loss_date = _win_loss_time_replace.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            _win_loss_time = "";
                        }
                        // ----- Turnover
                        JToken _turnover = _jo.SelectToken("$.PageData[" + i + "].BetAmount").ToString();
                        // ----- Effective Turnover
                        JToken _effective_turnover = _jo.SelectToken("$.PageData[" + i + "].Commissionable").ToString();
                        // ----- Member Win Loss
                        JToken _member_win_loss = _jo.SelectToken("$.PageData[" + i + "].Payoff").ToString();
                        
                        if (__display_count == 1)
                        {
                            __detect_header = true;
                            var header = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}", "Username", "Provider", "Games", "Bet Time", "Win Loss Time", "Turnover", "Effective Turnover", "Member Win Loss", "Bet Date", "Win Loss Date");
                            __DATA.AppendLine(header);
                        }
                        var data = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8},{9}", "\"" + _username + "\"", "\"" + _provider + "\"", "\"" + _games + "\"", "\"" + _bet_time + "\"", "\"" + _win_loss_time + "\"", "\"" + _turnover + "\"", "\"" + _effective_turnover + "\"", "\"" + _member_win_loss + "\"", "\"" + _bet_date + "\"", "\"" + _win_loss_date + "\"");
                        __DATA.AppendLine(data);
                    }

                    await ___BETAsync(__index_bet++);
                }
                else
                {
                    __display_count = 0;
                    // BET RECORD SAVING EXCEL
                    string _current_datetime = DateTime.Now.ToString("yyyy-MM-dd");

                    label_cl_status.ForeColor = Color.FromArgb(78, 122, 159);
                    label_cl_status.Text = "status: saving excel... --- BET RECORD";

                    if (!Directory.Exists(__file_location + "\\Cronos Data"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data");
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime);
                    }

                    if (!Directory.Exists(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bet Record"))
                    {
                        Directory.CreateDirectory(__file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bet Record");
                    }

                    string _folder_path_result = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bet Record\\" + __brand_code + "_BetRecord_" + _current_datetime + "_1.txt";
                    string _folder_path_result_xlsx = __file_location + "\\Cronos Data\\" + __brand_code + "\\" + _current_datetime + "\\Bet Record\\" + __brand_code + "_BetRecord_" + _current_datetime + "_1.xlsx";

                    if (File.Exists(_folder_path_result))
                    {
                        File.Delete(_folder_path_result);
                    }

                    if (File.Exists(_folder_path_result_xlsx))
                    {
                        File.Delete(_folder_path_result_xlsx);
                    }

                    File.WriteAllText(_folder_path_result, __DATA.ToString(), Encoding.UTF8);

                    Excel.Application app = new Excel.Application();
                    Excel.Workbook wb = app.Workbooks.Open(_folder_path_result, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet worksheet = wb.ActiveSheet;
                    worksheet.Activate();
                    worksheet.Application.ActiveWindow.SplitRow = 1;
                    worksheet.Application.ActiveWindow.FreezePanes = true;
                    Excel.Range firstRow = (Excel.Range)worksheet.Rows[1];
                    firstRow.AutoFilter(1,
                                        Type.Missing,
                                        Excel.XlAutoFilterOperator.xlAnd,
                                        Type.Missing,
                                        true);
                    worksheet.Columns[4].NumberFormat = "yyyy/MM/dd HH:mm:ss";
                    worksheet.Columns[5].NumberFormat = "yyyy/MM/dd HH:mm:ss";
                    worksheet.Columns[9].NumberFormat = "yyyy-MM-dd";
                    worksheet.Columns[10].NumberFormat = "yyyy-MM-dd";
                    Excel.Range usedRange = worksheet.UsedRange;
                    Excel.Range rows = usedRange.Rows;
                    int count = 0;
                    foreach (Excel.Range row in rows)
                    {
                        if (count == 0)
                        {
                            Excel.Range firstCell = row.Cells[1];

                            string firstCellValue = firstCell.Value as string;

                            if (!string.IsNullOrEmpty(firstCellValue))
                            {
                                row.Interior.Color = Color.FromArgb(27, 96, 168);
                                row.Font.Color = Color.FromArgb(255, 255, 255);
                            }

                            break;
                        }

                        count++;
                    }
                    int i_;
                    for (i_ = 1; i_ <= 21; i_++)
                    {
                        worksheet.Columns[i_].ColumnWidth = 20;
                    }
                    wb.SaveAs(_folder_path_result_xlsx, Excel.XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wb.Close();
                    app.Quit();
                    Marshal.ReleaseComObject(app);

                    if (File.Exists(_folder_path_result))
                    {
                        File.Delete(_folder_path_result);
                    }

                    __DATA.Clear();
                    __detect_header = false;
                    __display_count = 0;
                    __index_bet = 1;

                    // BET SEND TO DATABASE
                    // AUTO START

                    Properties.Settings.Default.______start_detect = "0";
                    Properties.Settings.Default.Save();

                    panel_status.Visible = false;
                    label_cl_status.Text = "-";
                    label_page_count.Text = "-";
                    label_total_records.Text = "-";
                    button_start.Visible = true;
                    button_start.Enabled = false;

                    if (__is_autostart)
                    {
                        comboBox_list.SelectedIndex = 0;

                        //SendMyBot("Reports has been completed.");
                        SendITSupport("Reports has been completed.");
                        SendReportsTeam("Reports has been completed.");
                    }
                    else
                    {
                        panel_filter.Enabled = true;
                    }
                }

                __send = 0;
            }
            catch (Exception err)
            {
                if (__is_login)
                {
                    __send++;
                    if (__send == 5)
                    {
                        // comment
                        //SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        await ___BETAsync(__index_bet);
                    }
                }
            }
        }

        private void ___GETDATA_VIPLIST()
        {
            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[" + __brand_code + ".VIP Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[" + __brand_code + ".VIP Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount.Text = "VIP List: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                if (i == 0)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 1)
                                {
                                    columns += reader[i].ToString();
                                }
                            }

                            __getdata_viplist.Add(columns);
                            columns = "";
                        }
                    }

                    conn.Close();
                }

                __send = 0;
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    // comment
                    //SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GETDATA_VIPLIST();
                }
            }
        }

        private void ___GETDATA_AFFIALIATELIST()
        {
            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[" + __brand_code + ".Affiliate Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[" + __brand_code + ".Affiliate Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount.Text = "Affiliate List: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                if (i == 0)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 1)
                                {
                                    columns += reader[i].ToString();
                                }
                            }

                            __getdata_affiliatelist.Add(columns);
                            columns = "";
                        }
                    }

                    conn.Close();
                }

                __send = 0;
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    // comment
                    //SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GETDATA_AFFIALIATELIST();
                }
            }
        }

        private void ___GETDATA_PAYMENTMETHODLIST()
        {
            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[" + __brand_code + ".Payment Type Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[" + __brand_code + ".Payment Type Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount.Text = "Payment Method List: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                if (i == 0)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 1)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 2)
                                {
                                    columns += reader[i].ToString();
                                }
                            }

                            __getdata_paymentmethodlist.Add(columns);
                            columns = "";
                        }
                    }

                    conn.Close();
                }

                __send = 0;
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    // comment
                    //SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GETDATA_PAYMENTMETHODLIST();
                }
            }
        }

        private void ___GETDATA_BONUSCODE()
        {
            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[" + __brand_code + ".Bonus Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[" + __brand_code + ".Bonus Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount.Text = "Bonus Code: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                if (i == 0)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 1)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 2)
                                {
                                    columns += reader[i].ToString();
                                }
                            }

                            __getdata_bonuscode.Add(columns);
                            columns = "";
                        }
                    }

                    conn.Close();
                }
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    // comment
                    //SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GETDATA_BONUSCODE();
                }
            }
        }

        private void ___GETDATA_PRODUCTCODE()
        {
            try
            {
                string connection = "Data Source=192.168.10.252;User ID=sa;password=Test@123;Initial Catalog=testrain;Integrated Security=True;Trusted_Connection=false;";

                using (SqlConnection conn = new SqlConnection(connection))
                {
                    conn.Open();
                    SqlCommand command = new SqlCommand("SELECT * FROM [testrain].[dbo].[" + __brand_code + ".Product Code]", conn);
                    SqlCommand command_count = new SqlCommand("SELECT COUNT(*) FROM [testrain].[dbo].[" + __brand_code + ".Product Code]", conn);
                    string columns = "";

                    Int32 getcount = (Int32)command_count.ExecuteScalar();

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            count++;
                            label_getdatacount.Text = "Product Code: " + count.ToString("N0") + " of " + getcount.ToString("N0");

                            Application.DoEvents();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                Application.DoEvents();
                                if (i == 0)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 1)
                                {
                                    columns += reader[i].ToString() + "*|*";
                                }
                                else if (i == 2)
                                {
                                    columns += reader[i].ToString();
                                }
                            }

                            __getdata_productcode.Add(columns);
                            columns = "";
                        }
                    }

                    panel_cl.Enabled = true;
                    label_getdatacount.Visible = false;
                    label_getdatacount.Text = "-";

                    conn.Close();
                }
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    // comment
                    //SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    ___GETDATA_PRODUCTCODE();
                }
            }
        }

        private void SendMyBot(string message)
        {
            try
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                string apiToken = "772918363:AAHn2ufmP3ocLEilQ1V-IHcqYMcSuFJHx5g";
                string chatId = "@allandrake";
                string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20" + Properties.Settings.Default.______server_ip + "%0ALocation:%20" + Properties.Settings.Default.______server_location + "%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message + "";
                urlString = String.Format(urlString, apiToken, chatId, text);
                WebRequest request = WebRequest.Create(urlString);
                Stream rs = request.GetResponse().GetResponseStream();
                StreamReader reader = new StreamReader(rs);
                string line = "";
                StringBuilder sb = new StringBuilder();
                while (line != null)
                {
                    line = reader.ReadLine();
                    if (line != null)
                        sb.Append(line);
                }
            }
            catch (Exception err)
            {
                if (err.ToString().ToLower().Contains("hexadecimal"))
                {
                    string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                    string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                    string apiToken = "772918363:AAHn2ufmP3ocLEilQ1V-IHcqYMcSuFJHx5g";
                    string chatId = "@allandrake";
                    string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20192.168.10.60%0ALocation:%20192.168.10.60%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message + "";
                    urlString = String.Format(urlString, apiToken, chatId, text);
                    WebRequest request = WebRequest.Create(urlString);
                    Stream rs = request.GetResponse().GetResponseStream();
                    StreamReader reader = new StreamReader(rs);
                    string line = "";
                    StringBuilder sb = new StringBuilder();
                    while (line != null)
                    {
                        line = reader.ReadLine();
                        if (line != null)
                            sb.Append(line);
                    }

                    __is_close = false;
                    Environment.Exit(0);
                }
                else
                {
                    __send++;
                    if (__send == 5)
                    {
                        SendITSupport("There's a problem to the server, please re-open the application.");
                        SendMyBot(err.ToString());

                        __is_close = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        ___WaitNSeconds(10);
                        SendMyBot(message);
                    }
                }
            }
        }

        private void SendITSupport(string message)
        {
            if (__is_send)
            {
                try
                {
                    string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                    string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                    string apiToken = "612187347:AAE9doWWcStpWrDrfpOod89qGSxCJ5JwQO4";
                    string chatId = "@it_support_ssi";
                    string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20" + Properties.Settings.Default.______server_ip + "%0ALocation:%20" + Properties.Settings.Default.______server_location + "%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message + "";
                    urlString = String.Format(urlString, apiToken, chatId, text);
                    WebRequest request = WebRequest.Create(urlString);
                    Stream rs = request.GetResponse().GetResponseStream();
                    StreamReader reader = new StreamReader(rs);
                    string line = "";
                    StringBuilder sb = new StringBuilder();
                    while (line != null)
                    {
                        line = reader.ReadLine();
                        if (line != null)
                        {
                            sb.Append(line);
                        }
                    }
                }
                catch (Exception err)
                {
                    if (err.ToString().ToLower().Contains("hexadecimal"))
                    {
                        string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                        string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                        string apiToken = "612187347:AAE9doWWcStpWrDrfpOod89qGSxCJ5JwQO4";
                        string chatId = "@it_support_ssi";
                        string text = "-----" + __brand_code + " " + __app + "-----%0A%0AIP:%20192.168.10.60%0ALocation:%20192.168.10.60%0ADate%20and%20Time:%20[" + datetime + "]%0AMessage:%20" + message + "";
                        urlString = String.Format(urlString, apiToken, chatId, text);
                        WebRequest request = WebRequest.Create(urlString);
                        Stream rs = request.GetResponse().GetResponseStream();
                        StreamReader reader = new StreamReader(rs);
                        string line = "";
                        StringBuilder sb = new StringBuilder();
                        while (line != null)
                        {
                            line = reader.ReadLine();
                            if (line != null)
                            {
                                sb.Append(line);
                            }
                        }

                        __is_close = false;
                        Environment.Exit(0);
                    }
                    else
                    {
                        __send++;
                        if (__send == 5)
                        {
                            SendITSupport("There's a problem to the server, please re-open the application.");
                            SendMyBot(err.ToString());

                            __is_close = false;
                            Environment.Exit(0);
                        }
                        else
                        {
                            ___WaitNSeconds(10);
                            SendITSupport(message);
                        }
                    }
                }
            }
        }

        private void SendReportsTeam(string message)
        {
            try
            {
                string datetime = DateTime.Now.ToString("dd MMM HH:mm:ss");
                string urlString = "https://api.telegram.org/bot{0}/sendMessage?chat_id={1}&text={2}";
                string apiToken = "762890741:AAFwjSml3OgWrN07G_41YgIIzFAyxYLGE8Q";
                string chatId = "@cronos_data_reports_team";
                string text = "Brand:%20-----" + __brand_code + "-----%0ATime:%20[" + datetime + "]%0AMessage:%20" + message + "";
                urlString = String.Format(urlString, apiToken, chatId, text);
                WebRequest request = WebRequest.Create(urlString);
                Stream rs = request.GetResponse().GetResponseStream();
                StreamReader reader = new StreamReader(rs);
                string line = "";
                StringBuilder sb = new StringBuilder();
                while (line != null)
                {
                    line = reader.ReadLine();
                    if (line != null)
                        sb.Append(line);
                }

                __send = 0;
            }
            catch (Exception err)
            {
                __send++;
                if (__send == 5)
                {
                    //SendITSupport("There's a problem to the server, please re-open the application.");
                    SendMyBot(err.ToString());

                    Environment.Exit(0);
                }
                else
                {
                    ___WaitNSeconds(10);
                    SendReportsTeam(message);
                }
            }
        }


        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        const UInt32 WM_CLOSE = 0x0010;

        void ___CloseMessageBox()
        {
            IntPtr windowPtr = FindWindowByCaption(IntPtr.Zero, "Message from webpage");

            if (windowPtr == IntPtr.Zero)
            {
                return;
            }

            SendMessage(windowPtr, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        private void timer_close_message_box_Tick(object sender, EventArgs e)
        {
            ___CloseMessageBox();
        }

        private void timer_flush_memory_Tick(object sender, EventArgs e)
        {
            ___FlushMemory();
        }

        public static void ___FlushMemory()
        {
            Process prs = Process.GetCurrentProcess();
            try
            {
                prs.MinWorkingSet = (IntPtr)(300000);
            }
            catch (Exception err)
            {
                // leave blank
            }
        }

        private void timer_midnight_Tick(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.______midnight_time == "")
            {
                DateTime today = DateTime.Now;
                DateTime date = today.AddDays(1);
                Properties.Settings.Default.______midnight_time = date.ToString("yyyy-MM-dd 00:30");
                Properties.Settings.Default.Save();
            }
            else
            {
                DateTime today = DateTime.Now;
                if (Properties.Settings.Default.______midnight_time == today.ToString("yyyy-MM-dd HH:mm"))
                {
                    if (Properties.Settings.Default.______start_detect == "0")
                    {
                        Properties.Settings.Default.______midnight_time = "";
                        Properties.Settings.Default.Save();

                        __DATA.Clear();
                            __conn_id = "";
                        __index_reg = 1;
                        __index_dep = 1;
                        __index_bon = 1;
                        __index_bet = 1;
                        __display_count = 0;
                        __getdata_viplist.Clear();
                        __getdata_affiliatelist.Clear();
                        __getdata_paymentmethodlist.Clear();
                        __getdata_bonuscode.Clear();
                        __getdata_productcode.Clear();
                        __DATA.Clear();
                        ___GETDATA_VIPLIST();
                        ___GETDATA_AFFIALIATELIST();
                        ___GETDATA_PAYMENTMETHODLIST();
                        ___GETDATA_BONUSCODE();
                        ___GETDATA_PRODUCTCODE();
                        Properties.Settings.Default.______start_detect = "1";
                        Properties.Settings.Default.Save();
                        comboBox_list.SelectedIndex = 1;
                        comboBox_list.SelectedIndex = 0;
                        button_start.Enabled = true;
                        button_start.PerformClick();
                    }
                }
                else
                {
                    string start_datetime = today.ToString("yyyy-MM-dd HH:mm");
                    DateTime start = DateTime.ParseExact(start_datetime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                    string end_datetime = Properties.Settings.Default.______midnight_time;
                    DateTime end = DateTime.ParseExact(end_datetime, "yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                    if (start > end)
                    {
                        if (Properties.Settings.Default.______start_detect == "0")
                        {
                            Properties.Settings.Default.______midnight_time = "";
                            Properties.Settings.Default.Save();

                            __DATA.Clear();
                            __conn_id = "";
                            __index_reg = 1;
                            __index_dep = 1;
                            __index_bon = 1;
                            __index_bet = 1;
                            __display_count = 0;
                            __getdata_viplist.Clear();
                            __getdata_affiliatelist.Clear();
                            __getdata_paymentmethodlist.Clear();
                            __getdata_bonuscode.Clear();
                            __getdata_productcode.Clear();
                            __DATA.Clear();
                            ___GETDATA_VIPLIST();
                            ___GETDATA_AFFIALIATELIST();
                            ___GETDATA_PAYMENTMETHODLIST();
                            ___GETDATA_BONUSCODE();
                            ___GETDATA_PRODUCTCODE();
                            Properties.Settings.Default.______start_detect = "1";
                            Properties.Settings.Default.Save();
                            comboBox_list.SelectedIndex = 1;
                            comboBox_list.SelectedIndex = 0;
                            button_start.Enabled = true;
                            button_start.PerformClick();
                        }
                    }
                }
            }
        }

        private void label_navigate_up_Click(object sender, EventArgs e)
        {
            __mainform_handler = Application.OpenForms[0];
            __mainform_handler.Size = new Size(569, 208);
            panel_loader.Visible = true;
            label_navigate_up.Enabled = false;
        }

        private void label_navigate_down_Click(object sender, EventArgs e)
        {
            __mainform_handler = Application.OpenForms[0];
            __mainform_handler.Size = new Size(569, 514);
            panel_loader.Visible = false;
            label_navigate_up.Enabled = true;
        }

        private void panel_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (__is_send)
            {
                __is_send = false;
                MessageBox.Show("Telegram Notification is Disabled.", __brand_code + " " + __app, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                __is_send = true;
                MessageBox.Show("Telegram Notification is Enabled.", __brand_code + " " + __app, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ___WaitNSeconds(int sec)
        {
            if (sec < 1) return;
            DateTime _desired = DateTime.Now.AddSeconds(sec);
            while (DateTime.Now < _desired)
            {
                Application.DoEvents();
            }
        }
    }
}

// clear
//private int __index_reg = 1;
//private int __index_dep = 1;
//private int __index_bon = 1;
//private int __index_bet = 1;
//private int __display_count = 0;
// __getdata_viplist
// __getdata_affiliatelist
// __getdata_paymentmethodlist
// __DATA