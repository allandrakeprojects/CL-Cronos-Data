namespace CL_Cronos_Data
{
    partial class Main_Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main_Form));
            this.panel_header = new System.Windows.Forms.Panel();
            this.panel = new System.Windows.Forms.Panel();
            this.label_title = new System.Windows.Forms.Label();
            this.pictureBox_minimize = new System.Windows.Forms.PictureBox();
            this.pictureBox_close = new System.Windows.Forms.PictureBox();
            this.label_title_cl = new System.Windows.Forms.Label();
            this.panel_cl = new System.Windows.Forms.Panel();
            this.webBrowser = new System.Windows.Forms.WebBrowser();
            this.panel_status = new System.Windows.Forms.Panel();
            this.label_insert = new System.Windows.Forms.Label();
            this.button_proceed = new System.Windows.Forms.Button();
            this.label_locatefolder = new System.Windows.Forms.Label();
            this.panel_datetime = new System.Windows.Forms.Panel();
            this.label_elapsed = new System.Windows.Forms.Label();
            this.label_elapsed_1 = new System.Windows.Forms.Label();
            this.label_start_datetime_1 = new System.Windows.Forms.Label();
            this.label_finish_datetime = new System.Windows.Forms.Label();
            this.label_finish_datetime_1 = new System.Windows.Forms.Label();
            this.label_start_datetime = new System.Windows.Forms.Label();
            this.pictureBox_loader = new System.Windows.Forms.PictureBox();
            this.label_total_records = new System.Windows.Forms.Label();
            this.label_page_count = new System.Windows.Forms.Label();
            this.label_page_count_1 = new System.Windows.Forms.Label();
            this.label_total_records_1 = new System.Windows.Forms.Label();
            this.label_cl_status = new System.Windows.Forms.Label();
            this.label_navigate_up = new System.Windows.Forms.Label();
            this.panel_filter = new System.Windows.Forms.Panel();
            this.comboBox_list = new System.Windows.Forms.ComboBox();
            this.comboBox = new System.Windows.Forms.ComboBox();
            this.dateTimePicker_end = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker_start = new System.Windows.Forms.DateTimePicker();
            this.label_start = new System.Windows.Forms.Label();
            this.label_end = new System.Windows.Forms.Label();
            this.button_stop = new System.Windows.Forms.Button();
            this.button_start = new System.Windows.Forms.Button();
            this.panel_footer = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.label_status = new System.Windows.Forms.Label();
            this.label_status_1 = new System.Windows.Forms.Label();
            this.label_cycle_in = new System.Windows.Forms.Label();
            this.label_cycle_in_1 = new System.Windows.Forms.Label();
            this.label_version = new System.Windows.Forms.Label();
            this.label_updates = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label_getdatacount = new System.Windows.Forms.Label();
            this.label_count = new System.Windows.Forms.Label();
            this.timer_landing = new System.Windows.Forms.Timer(this.components);
            this.timer_cycle_in = new System.Windows.Forms.Timer(this.components);
            this.timer_start_button = new System.Windows.Forms.Timer(this.components);
            this.timer_elapsed = new System.Windows.Forms.Timer(this.components);
            this.timer_close_message_box = new System.Windows.Forms.Timer(this.components);
            this.timer_flush_memory = new System.Windows.Forms.Timer(this.components);
            this.timer_midnight = new System.Windows.Forms.Timer(this.components);
            this.panel_loader = new System.Windows.Forms.Panel();
            this.label_navigate_down = new System.Windows.Forms.Label();
            this.label_brand = new System.Windows.Forms.Label();
            this.pictureBox_loader_loader = new System.Windows.Forms.PictureBox();
            this.panel_landing = new System.Windows.Forms.Panel();
            this.pictureBox_landing = new System.Windows.Forms.PictureBox();
            this.panel_header.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).BeginInit();
            this.panel_cl.SuspendLayout();
            this.panel_status.SuspendLayout();
            this.panel_datetime.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_loader)).BeginInit();
            this.panel_filter.SuspendLayout();
            this.panel_footer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel_loader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_loader_loader)).BeginInit();
            this.panel_landing.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_landing)).BeginInit();
            this.SuspendLayout();
            // 
            // panel_header
            // 
            this.panel_header.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.panel_header.Controls.Add(this.panel);
            this.panel_header.Controls.Add(this.label_title);
            this.panel_header.Controls.Add(this.pictureBox_minimize);
            this.panel_header.Controls.Add(this.pictureBox_close);
            this.panel_header.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel_header.Location = new System.Drawing.Point(0, 0);
            this.panel_header.Name = "panel_header";
            this.panel_header.Size = new System.Drawing.Size(569, 45);
            this.panel_header.TabIndex = 0;
            this.panel_header.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel_header_MouseDown);
            // 
            // panel
            // 
            this.panel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(96)))), ((int)(((byte)(168)))));
            this.panel.Location = new System.Drawing.Point(-12, -5);
            this.panel.Name = "panel";
            this.panel.Size = new System.Drawing.Size(170, 10);
            this.panel.TabIndex = 1;
            this.panel.Visible = false;
            this.panel.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.panel_MouseDoubleClick);
            this.panel.MouseDown += new System.Windows.Forms.MouseEventHandler(this.panel_MouseDown);
            // 
            // label_title
            // 
            this.label_title.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_title.ForeColor = System.Drawing.Color.WhiteSmoke;
            this.label_title.Location = new System.Drawing.Point(2, 0);
            this.label_title.Name = "label_title";
            this.label_title.Size = new System.Drawing.Size(166, 45);
            this.label_title.TabIndex = 2;
            this.label_title.Text = "Cronos Data";
            this.label_title.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_title.Visible = false;
            this.label_title.MouseDown += new System.Windows.Forms.MouseEventHandler(this.label_title_MouseDown);
            // 
            // pictureBox_minimize
            // 
            this.pictureBox_minimize.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_minimize.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_minimize.Image = global::CL_Cronos_Data.Properties.Resources.minus;
            this.pictureBox_minimize.Location = new System.Drawing.Point(481, 10);
            this.pictureBox_minimize.Name = "pictureBox_minimize";
            this.pictureBox_minimize.Size = new System.Drawing.Size(24, 24);
            this.pictureBox_minimize.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox_minimize.TabIndex = 1;
            this.pictureBox_minimize.TabStop = false;
            this.pictureBox_minimize.Visible = false;
            this.pictureBox_minimize.Click += new System.EventHandler(this.pictureBox_minimize_Click);
            // 
            // pictureBox_close
            // 
            this.pictureBox_close.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.pictureBox_close.Cursor = System.Windows.Forms.Cursors.Hand;
            this.pictureBox_close.Image = global::CL_Cronos_Data.Properties.Resources.close;
            this.pictureBox_close.Location = new System.Drawing.Point(519, 10);
            this.pictureBox_close.Name = "pictureBox_close";
            this.pictureBox_close.Size = new System.Drawing.Size(24, 24);
            this.pictureBox_close.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox_close.TabIndex = 0;
            this.pictureBox_close.TabStop = false;
            this.pictureBox_close.Visible = false;
            this.pictureBox_close.Click += new System.EventHandler(this.pictureBox_close_Click);
            // 
            // label_title_cl
            // 
            this.label_title_cl.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_title_cl.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(96)))), ((int)(((byte)(168)))));
            this.label_title_cl.Location = new System.Drawing.Point(9, 4);
            this.label_title_cl.Name = "label_title_cl";
            this.label_title_cl.Size = new System.Drawing.Size(472, 30);
            this.label_title_cl.TabIndex = 2;
            this.label_title_cl.Text = "Chang Lee";
            this.label_title_cl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel_cl
            // 
            this.panel_cl.Controls.Add(this.webBrowser);
            this.panel_cl.Controls.Add(this.panel_status);
            this.panel_cl.Controls.Add(this.label_navigate_up);
            this.panel_cl.Controls.Add(this.label_title_cl);
            this.panel_cl.Controls.Add(this.panel_filter);
            this.panel_cl.Enabled = false;
            this.panel_cl.Location = new System.Drawing.Point(17, 70);
            this.panel_cl.Name = "panel_cl";
            this.panel_cl.Size = new System.Drawing.Size(534, 408);
            this.panel_cl.TabIndex = 4;
            this.panel_cl.Paint += new System.Windows.Forms.PaintEventHandler(this.panel_cl_Paint);
            // 
            // webBrowser
            // 
            this.webBrowser.Location = new System.Drawing.Point(6, 35);
            this.webBrowser.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser.Name = "webBrowser";
            this.webBrowser.Size = new System.Drawing.Size(522, 367);
            this.webBrowser.TabIndex = 13;
            this.webBrowser.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.WebBrowser_DocumentCompletedAsync);
            // 
            // panel_status
            // 
            this.panel_status.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.panel_status.Controls.Add(this.label_insert);
            this.panel_status.Controls.Add(this.button_proceed);
            this.panel_status.Controls.Add(this.label_locatefolder);
            this.panel_status.Controls.Add(this.panel_datetime);
            this.panel_status.Controls.Add(this.pictureBox_loader);
            this.panel_status.Controls.Add(this.label_total_records);
            this.panel_status.Controls.Add(this.label_page_count);
            this.panel_status.Controls.Add(this.label_page_count_1);
            this.panel_status.Controls.Add(this.label_total_records_1);
            this.panel_status.Controls.Add(this.label_cl_status);
            this.panel_status.Location = new System.Drawing.Point(7, 121);
            this.panel_status.Name = "panel_status";
            this.panel_status.Size = new System.Drawing.Size(524, 284);
            this.panel_status.TabIndex = 23;
            this.panel_status.Visible = false;
            // 
            // label_insert
            // 
            this.label_insert.Location = new System.Drawing.Point(382, 207);
            this.label_insert.Name = "label_insert";
            this.label_insert.Size = new System.Drawing.Size(125, 23);
            this.label_insert.TabIndex = 34;
            this.label_insert.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // button_proceed
            // 
            this.button_proceed.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_proceed.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button_proceed.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_proceed.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_proceed.ForeColor = System.Drawing.Color.White;
            this.button_proceed.Location = new System.Drawing.Point(382, 233);
            this.button_proceed.Name = "button_proceed";
            this.button_proceed.Size = new System.Drawing.Size(126, 28);
            this.button_proceed.TabIndex = 23;
            this.button_proceed.Text = "PROCEED";
            this.button_proceed.UseVisualStyleBackColor = false;
            this.button_proceed.Visible = false;
            // 
            // label_locatefolder
            // 
            this.label_locatefolder.AutoSize = true;
            this.label_locatefolder.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_locatefolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_locatefolder.Location = new System.Drawing.Point(436, 264);
            this.label_locatefolder.Name = "label_locatefolder";
            this.label_locatefolder.Size = new System.Drawing.Size(72, 13);
            this.label_locatefolder.TabIndex = 29;
            this.label_locatefolder.Text = "Locate Folder";
            this.label_locatefolder.Visible = false;
            // 
            // panel_datetime
            // 
            this.panel_datetime.Controls.Add(this.label_elapsed);
            this.panel_datetime.Controls.Add(this.label_elapsed_1);
            this.panel_datetime.Controls.Add(this.label_start_datetime_1);
            this.panel_datetime.Controls.Add(this.label_finish_datetime);
            this.panel_datetime.Controls.Add(this.label_finish_datetime_1);
            this.panel_datetime.Controls.Add(this.label_start_datetime);
            this.panel_datetime.Location = new System.Drawing.Point(66, 226);
            this.panel_datetime.Name = "panel_datetime";
            this.panel_datetime.Size = new System.Drawing.Size(287, 58);
            this.panel_datetime.TabIndex = 28;
            // 
            // label_elapsed
            // 
            this.label_elapsed.AutoSize = true;
            this.label_elapsed.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_elapsed.Location = new System.Drawing.Point(66, 38);
            this.label_elapsed.Name = "label_elapsed";
            this.label_elapsed.Size = new System.Drawing.Size(11, 15);
            this.label_elapsed.TabIndex = 29;
            this.label_elapsed.Text = "-";
            // 
            // label_elapsed_1
            // 
            this.label_elapsed_1.AutoSize = true;
            this.label_elapsed_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_elapsed_1.Location = new System.Drawing.Point(3, 36);
            this.label_elapsed_1.Name = "label_elapsed_1";
            this.label_elapsed_1.Size = new System.Drawing.Size(55, 15);
            this.label_elapsed_1.TabIndex = 28;
            this.label_elapsed_1.Text = "Elapsed:";
            // 
            // label_start_datetime_1
            // 
            this.label_start_datetime_1.AutoSize = true;
            this.label_start_datetime_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_start_datetime_1.Location = new System.Drawing.Point(3, 5);
            this.label_start_datetime_1.Name = "label_start_datetime_1";
            this.label_start_datetime_1.Size = new System.Drawing.Size(35, 15);
            this.label_start_datetime_1.TabIndex = 24;
            this.label_start_datetime_1.Text = "Start:";
            // 
            // label_finish_datetime
            // 
            this.label_finish_datetime.AutoSize = true;
            this.label_finish_datetime.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_finish_datetime.Location = new System.Drawing.Point(66, 21);
            this.label_finish_datetime.Name = "label_finish_datetime";
            this.label_finish_datetime.Size = new System.Drawing.Size(11, 15);
            this.label_finish_datetime.TabIndex = 27;
            this.label_finish_datetime.Text = "-";
            // 
            // label_finish_datetime_1
            // 
            this.label_finish_datetime_1.AutoSize = true;
            this.label_finish_datetime_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_finish_datetime_1.Location = new System.Drawing.Point(3, 21);
            this.label_finish_datetime_1.Name = "label_finish_datetime_1";
            this.label_finish_datetime_1.Size = new System.Drawing.Size(43, 15);
            this.label_finish_datetime_1.TabIndex = 25;
            this.label_finish_datetime_1.Text = "Finish:";
            // 
            // label_start_datetime
            // 
            this.label_start_datetime.AutoSize = true;
            this.label_start_datetime.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_start_datetime.Location = new System.Drawing.Point(66, 5);
            this.label_start_datetime.Name = "label_start_datetime";
            this.label_start_datetime.Size = new System.Drawing.Size(11, 15);
            this.label_start_datetime.TabIndex = 26;
            this.label_start_datetime.Text = "-";
            // 
            // pictureBox_loader
            // 
            this.pictureBox_loader.Image = global::CL_Cronos_Data.Properties.Resources.loader;
            this.pictureBox_loader.Location = new System.Drawing.Point(3, 180);
            this.pictureBox_loader.Name = "pictureBox_loader";
            this.pictureBox_loader.Size = new System.Drawing.Size(60, 101);
            this.pictureBox_loader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox_loader.TabIndex = 23;
            this.pictureBox_loader.TabStop = false;
            this.pictureBox_loader.Visible = false;
            // 
            // label_total_records
            // 
            this.label_total_records.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_total_records.Location = new System.Drawing.Point(258, 116);
            this.label_total_records.Name = "label_total_records";
            this.label_total_records.Size = new System.Drawing.Size(250, 18);
            this.label_total_records.TabIndex = 12;
            this.label_total_records.Text = "-";
            // 
            // label_page_count
            // 
            this.label_page_count.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_page_count.Location = new System.Drawing.Point(259, 86);
            this.label_page_count.Name = "label_page_count";
            this.label_page_count.Size = new System.Drawing.Size(249, 18);
            this.label_page_count.TabIndex = 13;
            this.label_page_count.Text = "-";
            // 
            // label_page_count_1
            // 
            this.label_page_count_1.AutoSize = true;
            this.label_page_count_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_page_count_1.Location = new System.Drawing.Point(134, 86);
            this.label_page_count_1.Name = "label_page_count_1";
            this.label_page_count_1.Size = new System.Drawing.Size(46, 18);
            this.label_page_count_1.TabIndex = 20;
            this.label_page_count_1.Text = "Page:";
            // 
            // label_total_records_1
            // 
            this.label_total_records_1.AutoSize = true;
            this.label_total_records_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_total_records_1.Location = new System.Drawing.Point(132, 116);
            this.label_total_records_1.Name = "label_total_records_1";
            this.label_total_records_1.Size = new System.Drawing.Size(98, 18);
            this.label_total_records_1.TabIndex = 18;
            this.label_total_records_1.Text = "Total Record:";
            // 
            // label_cl_status
            // 
            this.label_cl_status.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_cl_status.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.label_cl_status.Location = new System.Drawing.Point(3, 42);
            this.label_cl_status.Name = "label_cl_status";
            this.label_cl_status.Size = new System.Drawing.Size(518, 25);
            this.label_cl_status.TabIndex = 17;
            this.label_cl_status.Text = "-";
            this.label_cl_status.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label_navigate_up
            // 
            this.label_navigate_up.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_navigate_up.Enabled = false;
            this.label_navigate_up.Image = global::CL_Cronos_Data.Properties.Resources.navigate_up;
            this.label_navigate_up.Location = new System.Drawing.Point(303, 8);
            this.label_navigate_up.Name = "label_navigate_up";
            this.label_navigate_up.Size = new System.Drawing.Size(35, 23);
            this.label_navigate_up.TabIndex = 43;
            this.label_navigate_up.Click += new System.EventHandler(this.label_navigate_up_Click);
            // 
            // panel_filter
            // 
            this.panel_filter.Controls.Add(this.comboBox_list);
            this.panel_filter.Controls.Add(this.comboBox);
            this.panel_filter.Controls.Add(this.dateTimePicker_end);
            this.panel_filter.Controls.Add(this.dateTimePicker_start);
            this.panel_filter.Controls.Add(this.label_start);
            this.panel_filter.Controls.Add(this.label_end);
            this.panel_filter.Location = new System.Drawing.Point(3, 35);
            this.panel_filter.Name = "panel_filter";
            this.panel_filter.Size = new System.Drawing.Size(528, 80);
            this.panel_filter.TabIndex = 24;
            // 
            // comboBox_list
            // 
            this.comboBox_list.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_list.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox_list.FormattingEnabled = true;
            this.comboBox_list.Items.AddRange(new object[] {
            "Registration",
            "Payment Report",
            "Bonus Report",
            "Turnover Record",
            "Bet Record"});
            this.comboBox_list.Location = new System.Drawing.Point(69, 47);
            this.comboBox_list.Name = "comboBox_list";
            this.comboBox_list.Size = new System.Drawing.Size(133, 23);
            this.comboBox_list.TabIndex = 12;
            this.comboBox_list.SelectedIndexChanged += new System.EventHandler(this.comboBox_list_SelectedIndexChanged);
            // 
            // comboBox
            // 
            this.comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBox.FormattingEnabled = true;
            this.comboBox.Items.AddRange(new object[] {
            "Yesterday",
            "Last week",
            "Last month"});
            this.comboBox.Location = new System.Drawing.Point(69, 15);
            this.comboBox.Name = "comboBox";
            this.comboBox.Size = new System.Drawing.Size(133, 23);
            this.comboBox.TabIndex = 7;
            this.comboBox.SelectedIndexChanged += new System.EventHandler(this.comboBox_SelectedIndexChanged);
            // 
            // dateTimePicker_end
            // 
            this.dateTimePicker_end.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_end.Location = new System.Drawing.Point(296, 44);
            this.dateTimePicker_end.Name = "dateTimePicker_end";
            this.dateTimePicker_end.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_end.TabIndex = 11;
            // 
            // dateTimePicker_start
            // 
            this.dateTimePicker_start.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker_start.Location = new System.Drawing.Point(296, 16);
            this.dateTimePicker_start.Name = "dateTimePicker_start";
            this.dateTimePicker_start.Size = new System.Drawing.Size(169, 21);
            this.dateTimePicker_start.TabIndex = 8;
            // 
            // label_start
            // 
            this.label_start.AutoSize = true;
            this.label_start.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_start.Location = new System.Drawing.Point(225, 20);
            this.label_start.Name = "label_start";
            this.label_start.Size = new System.Drawing.Size(66, 15);
            this.label_start.TabIndex = 9;
            this.label_start.Text = "Start Time:";
            // 
            // label_end
            // 
            this.label_end.AutoSize = true;
            this.label_end.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_end.Location = new System.Drawing.Point(225, 49);
            this.label_end.Name = "label_end";
            this.label_end.Size = new System.Drawing.Size(63, 15);
            this.label_end.TabIndex = 10;
            this.label_end.Text = "End Time:";
            // 
            // button_stop
            // 
            this.button_stop.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_stop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button_stop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_stop.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_stop.ForeColor = System.Drawing.Color.White;
            this.button_stop.Image = global::CL_Cronos_Data.Properties.Resources.stop;
            this.button_stop.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_stop.Location = new System.Drawing.Point(204, 290);
            this.button_stop.Name = "button_stop";
            this.button_stop.Padding = new System.Windows.Forms.Padding(15, 0, 22, 0);
            this.button_stop.Size = new System.Drawing.Size(153, 59);
            this.button_stop.TabIndex = 33;
            this.button_stop.Text = "STOP";
            this.button_stop.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_stop.UseVisualStyleBackColor = false;
            this.button_stop.Visible = false;
            this.button_stop.Click += new System.EventHandler(this.button_stop_Click);
            // 
            // button_start
            // 
            this.button_start.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button_start.Cursor = System.Windows.Forms.Cursors.Hand;
            this.button_start.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button_start.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button_start.ForeColor = System.Drawing.Color.White;
            this.button_start.Image = global::CL_Cronos_Data.Properties.Resources.start;
            this.button_start.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_start.Location = new System.Drawing.Point(204, 290);
            this.button_start.Name = "button_start";
            this.button_start.Padding = new System.Windows.Forms.Padding(15, 0, 15, 0);
            this.button_start.Size = new System.Drawing.Size(153, 59);
            this.button_start.TabIndex = 1;
            this.button_start.Text = "START";
            this.button_start.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button_start.UseVisualStyleBackColor = false;
            this.button_start.Visible = false;
            this.button_start.Click += new System.EventHandler(this.button_start_Click);
            // 
            // panel_footer
            // 
            this.panel_footer.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.panel_footer.Controls.Add(this.button1);
            this.panel_footer.Controls.Add(this.label_status);
            this.panel_footer.Controls.Add(this.label_status_1);
            this.panel_footer.Controls.Add(this.label_cycle_in);
            this.panel_footer.Controls.Add(this.label_cycle_in_1);
            this.panel_footer.Controls.Add(this.label_version);
            this.panel_footer.Controls.Add(this.label_updates);
            this.panel_footer.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel_footer.Location = new System.Drawing.Point(0, 487);
            this.panel_footer.Name = "panel_footer";
            this.panel_footer.Size = new System.Drawing.Size(569, 27);
            this.panel_footer.TabIndex = 4;
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(433, 6);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 16);
            this.button1.TabIndex = 1;
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Visible = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label_status
            // 
            this.label_status.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_status.ForeColor = System.Drawing.Color.White;
            this.label_status.Location = new System.Drawing.Point(50, 0);
            this.label_status.Name = "label_status";
            this.label_status.Size = new System.Drawing.Size(57, 27);
            this.label_status.TabIndex = 3;
            this.label_status.Text = "-";
            this.label_status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label_status.Visible = false;
            // 
            // label_status_1
            // 
            this.label_status_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_status_1.ForeColor = System.Drawing.Color.White;
            this.label_status_1.Location = new System.Drawing.Point(-1, 1);
            this.label_status_1.Name = "label_status_1";
            this.label_status_1.Size = new System.Drawing.Size(69, 27);
            this.label_status_1.TabIndex = 4;
            this.label_status_1.Text = "Status:";
            this.label_status_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_status_1.Visible = false;
            // 
            // label_cycle_in
            // 
            this.label_cycle_in.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_cycle_in.ForeColor = System.Drawing.Color.White;
            this.label_cycle_in.Location = new System.Drawing.Point(157, 0);
            this.label_cycle_in.Name = "label_cycle_in";
            this.label_cycle_in.Size = new System.Drawing.Size(147, 27);
            this.label_cycle_in.TabIndex = 1;
            this.label_cycle_in.Text = "-";
            this.label_cycle_in.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.label_cycle_in.Visible = false;
            // 
            // label_cycle_in_1
            // 
            this.label_cycle_in_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_cycle_in_1.ForeColor = System.Drawing.Color.White;
            this.label_cycle_in_1.Location = new System.Drawing.Point(105, 1);
            this.label_cycle_in_1.Name = "label_cycle_in_1";
            this.label_cycle_in_1.Size = new System.Drawing.Size(69, 27);
            this.label_cycle_in_1.TabIndex = 2;
            this.label_cycle_in_1.Text = "Cycle In:";
            this.label_cycle_in_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_cycle_in_1.Visible = false;
            // 
            // label_version
            // 
            this.label_version.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label_version.AutoSize = true;
            this.label_version.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_version.ForeColor = System.Drawing.Color.White;
            this.label_version.Location = new System.Drawing.Point(512, 8);
            this.label_version.Name = "label_version";
            this.label_version.Size = new System.Drawing.Size(43, 13);
            this.label_version.TabIndex = 1;
            this.label_version.Text = "v1.0.0";
            this.label_version.Visible = false;
            // 
            // label_updates
            // 
            this.label_updates.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label_updates.AutoSize = true;
            this.label_updates.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_updates.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_updates.ForeColor = System.Drawing.Color.White;
            this.label_updates.Location = new System.Drawing.Point(406, 7);
            this.label_updates.Name = "label_updates";
            this.label_updates.Size = new System.Drawing.Size(99, 13);
            this.label_updates.TabIndex = 0;
            this.label_updates.Text = "Check for Updates.";
            this.label_updates.Visible = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1});
            this.dataGridView1.Location = new System.Drawing.Point(1167, 55);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(242, 170);
            this.dataGridView1.TabIndex = 30;
            // 
            // Column1
            // 
            this.Column1.HeaderText = "Column1";
            this.Column1.Name = "Column1";
            // 
            // label_getdatacount
            // 
            this.label_getdatacount.Location = new System.Drawing.Point(15, 47);
            this.label_getdatacount.Name = "label_getdatacount";
            this.label_getdatacount.Size = new System.Drawing.Size(534, 20);
            this.label_getdatacount.TabIndex = 34;
            this.label_getdatacount.Text = "-";
            this.label_getdatacount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label_count
            // 
            this.label_count.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_count.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(122)))), ((int)(((byte)(159)))));
            this.label_count.Location = new System.Drawing.Point(38, 350);
            this.label_count.Name = "label_count";
            this.label_count.Size = new System.Drawing.Size(498, 38);
            this.label_count.TabIndex = 0;
            this.label_count.Text = "-";
            this.label_count.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.label_count.Visible = false;
            // 
            // timer_landing
            // 
            this.timer_landing.Interval = 2000;
            this.timer_landing.Tick += new System.EventHandler(this.timer_landing_Tick);
            // 
            // timer_cycle_in
            // 
            this.timer_cycle_in.Enabled = true;
            this.timer_cycle_in.Interval = 1000;
            this.timer_cycle_in.Tick += new System.EventHandler(this.timer_cycle_in_Tick);
            // 
            // timer_start_button
            // 
            this.timer_start_button.Interval = 1000;
            this.timer_start_button.Tick += new System.EventHandler(this.timer_start_button_TickAsync);
            // 
            // timer_elapsed
            // 
            this.timer_elapsed.Interval = 1000;
            this.timer_elapsed.Tick += new System.EventHandler(this.timer_elapsed_Tick);
            // 
            // timer_close_message_box
            // 
            this.timer_close_message_box.Enabled = true;
            this.timer_close_message_box.Tick += new System.EventHandler(this.timer_close_message_box_Tick);
            // 
            // timer_flush_memory
            // 
            this.timer_flush_memory.Enabled = true;
            this.timer_flush_memory.Interval = 2000;
            this.timer_flush_memory.Tick += new System.EventHandler(this.timer_flush_memory_Tick);
            // 
            // timer_midnight
            // 
            this.timer_midnight.Enabled = true;
            this.timer_midnight.Interval = 10000;
            this.timer_midnight.Tick += new System.EventHandler(this.timer_midnight_Tick);
            // 
            // panel_loader
            // 
            this.panel_loader.Controls.Add(this.label_navigate_down);
            this.panel_loader.Controls.Add(this.label_brand);
            this.panel_loader.Controls.Add(this.pictureBox_loader_loader);
            this.panel_loader.Location = new System.Drawing.Point(-11, 47);
            this.panel_loader.Name = "panel_loader";
            this.panel_loader.Size = new System.Drawing.Size(600, 134);
            this.panel_loader.TabIndex = 37;
            this.panel_loader.Visible = false;
            // 
            // label_navigate_down
            // 
            this.label_navigate_down.Cursor = System.Windows.Forms.Cursors.Hand;
            this.label_navigate_down.Image = global::CL_Cronos_Data.Properties.Resources.navigate_down;
            this.label_navigate_down.Location = new System.Drawing.Point(331, 6);
            this.label_navigate_down.Name = "label_navigate_down";
            this.label_navigate_down.Size = new System.Drawing.Size(35, 23);
            this.label_navigate_down.TabIndex = 42;
            this.label_navigate_down.Click += new System.EventHandler(this.label_navigate_down_Click);
            // 
            // label_brand
            // 
            this.label_brand.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_brand.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(27)))), ((int)(((byte)(96)))), ((int)(((byte)(168)))));
            this.label_brand.Location = new System.Drawing.Point(18, -6);
            this.label_brand.Name = "label_brand";
            this.label_brand.Size = new System.Drawing.Size(510, 46);
            this.label_brand.TabIndex = 26;
            this.label_brand.Text = "Chang Lee";
            this.label_brand.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pictureBox_loader_loader
            // 
            this.pictureBox_loader_loader.Image = global::CL_Cronos_Data.Properties.Resources.loader;
            this.pictureBox_loader_loader.Location = new System.Drawing.Point(258, 32);
            this.pictureBox_loader_loader.Name = "pictureBox_loader_loader";
            this.pictureBox_loader_loader.Size = new System.Drawing.Size(60, 101);
            this.pictureBox_loader_loader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox_loader_loader.TabIndex = 41;
            this.pictureBox_loader_loader.TabStop = false;
            // 
            // panel_landing
            // 
            this.panel_landing.Controls.Add(this.pictureBox_landing);
            this.panel_landing.Location = new System.Drawing.Point(-1, 21);
            this.panel_landing.Name = "panel_landing";
            this.panel_landing.Size = new System.Drawing.Size(572, 472);
            this.panel_landing.TabIndex = 38;
            // 
            // pictureBox_landing
            // 
            this.pictureBox_landing.Image = global::CL_Cronos_Data.Properties.Resources.cronos_data;
            this.pictureBox_landing.Location = new System.Drawing.Point(226, 170);
            this.pictureBox_landing.Name = "pictureBox_landing";
            this.pictureBox_landing.Size = new System.Drawing.Size(135, 134);
            this.pictureBox_landing.TabIndex = 0;
            this.pictureBox_landing.TabStop = false;
            // 
            // Main_Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(569, 514);
            this.Controls.Add(this.panel_loader);
            this.Controls.Add(this.panel_landing);
            this.Controls.Add(this.button_start);
            this.Controls.Add(this.label_count);
            this.Controls.Add(this.button_stop);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.panel_footer);
            this.Controls.Add(this.panel_cl);
            this.Controls.Add(this.panel_header);
            this.Controls.Add(this.label_getdatacount);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Main_Form";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "YB Cronos Data";
            this.Load += new System.EventHandler(this.Main_Form_Load);
            this.Shown += new System.EventHandler(this.Main_Form_Shown);
            this.panel_header.ResumeLayout(false);
            this.panel_header.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_minimize)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_close)).EndInit();
            this.panel_cl.ResumeLayout(false);
            this.panel_status.ResumeLayout(false);
            this.panel_status.PerformLayout();
            this.panel_datetime.ResumeLayout(false);
            this.panel_datetime.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_loader)).EndInit();
            this.panel_filter.ResumeLayout(false);
            this.panel_filter.PerformLayout();
            this.panel_footer.ResumeLayout(false);
            this.panel_footer.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel_loader.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_loader_loader)).EndInit();
            this.panel_landing.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox_landing)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel_header;
        private System.Windows.Forms.Label label_title_cl;
        private System.Windows.Forms.Panel panel_cl;
        private System.Windows.Forms.Label label_title;
        private System.Windows.Forms.ComboBox comboBox;
        private System.Windows.Forms.DateTimePicker dateTimePicker_start;
        private System.Windows.Forms.Label label_start;
        private System.Windows.Forms.Label label_end;
        private System.Windows.Forms.DateTimePicker dateTimePicker_end;
        private System.Windows.Forms.Button button_start;
        private System.Windows.Forms.Panel panel_filter;
        private System.Windows.Forms.PictureBox pictureBox_minimize;
        private System.Windows.Forms.PictureBox pictureBox_close;
        private System.Windows.Forms.Panel panel_footer;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.Label label_updates;
        private System.Windows.Forms.Label label_version;
        private System.Windows.Forms.ComboBox comboBox_list;
        private System.Windows.Forms.Button button_stop;
        private System.Windows.Forms.Label label_getdatacount;
        private System.Windows.Forms.Panel panel_status;
        private System.Windows.Forms.Label label_insert;
        private System.Windows.Forms.Button button_proceed;
        private System.Windows.Forms.Label label_locatefolder;
        private System.Windows.Forms.Panel panel_datetime;
        private System.Windows.Forms.Label label_elapsed;
        private System.Windows.Forms.Label label_elapsed_1;
        private System.Windows.Forms.Label label_start_datetime_1;
        private System.Windows.Forms.Label label_finish_datetime;
        private System.Windows.Forms.Label label_finish_datetime_1;
        private System.Windows.Forms.Label label_start_datetime;
        private System.Windows.Forms.PictureBox pictureBox_loader;
        private System.Windows.Forms.Label label_total_records;
        private System.Windows.Forms.Label label_page_count;
        private System.Windows.Forms.Label label_page_count_1;
        private System.Windows.Forms.Label label_total_records_1;
        private System.Windows.Forms.Label label_cl_status;
        private System.Windows.Forms.Panel panel;
        private System.Windows.Forms.Label label_cycle_in;
        private System.Windows.Forms.Label label_cycle_in_1;
        private System.Windows.Forms.Label label_status;
        private System.Windows.Forms.Label label_status_1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label_navigate_up;
        private System.Windows.Forms.Label label_count;
        private System.Windows.Forms.Timer timer_landing;
        private System.Windows.Forms.Timer timer_cycle_in;
        private System.Windows.Forms.Timer timer_start_button;
        private System.Windows.Forms.Timer timer_elapsed;
        private System.Windows.Forms.WebBrowser webBrowser;
        private System.Windows.Forms.Timer timer_close_message_box;
        private System.Windows.Forms.Timer timer_flush_memory;
        private System.Windows.Forms.Timer timer_midnight;
        private System.Windows.Forms.Panel panel_loader;
        private System.Windows.Forms.Label label_navigate_down;
        private System.Windows.Forms.Label label_brand;
        private System.Windows.Forms.PictureBox pictureBox_loader_loader;
        private System.Windows.Forms.Panel panel_landing;
        private System.Windows.Forms.PictureBox pictureBox_landing;
    }
}