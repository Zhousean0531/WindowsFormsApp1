namespace WindowsFormsApp1
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.CylinderPage = new System.Windows.Forms.TabPage();
            this.CYLTypeBox = new System.Windows.Forms.ComboBox();
            this.CylinderTestDateBox = new System.Windows.Forms.DateTimePicker();
            this.CylinderBox = new System.Windows.Forms.DataGridView();
            this.CYLSN = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYLWeight = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Particle_In = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Particle_out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_IPA_in = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_IPA_out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Acetone_In = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Acetone_out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Nontarget_in = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Nontarget_out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYL_Pressure_Drop = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CYLMaterialSNBox = new System.Windows.Forms.TextBox();
            this.ReCylinderNOBox = new System.Windows.Forms.TextBox();
            this.CylinderCustmorBox = new System.Windows.Forms.TextBox();
            this.CylinderNOBox = new System.Windows.Forms.TextBox();
            this.CYLType = new System.Windows.Forms.Label();
            this.CylinderMaterialSN = new System.Windows.Forms.Label();
            this.CylinderReportNOBox = new System.Windows.Forms.TextBox();
            this.ReCylinderNO = new System.Windows.Forms.Label();
            this.CylinderCustmor = new System.Windows.Forms.Label();
            this.CylinderNO = new System.Windows.Forms.Label();
            this.CylinderReportNO = new System.Windows.Forms.Label();
            this.CylinderTestDate = new System.Windows.Forms.Label();
            this.CylinderRawPage = new System.Windows.Forms.TabPage();
            this.chkAsh = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.chkMoisture = new System.Windows.Forms.CheckBox();
            this.CylinderRawPressureBox = new System.Windows.Forms.TextBox();
            this.label20 = new System.Windows.Forms.Label();
            this.CylinderRawArriveDateBox = new System.Windows.Forms.DateTimePicker();
            this.CylinderRawTestDateBox = new System.Windows.Forms.DateTimePicker();
            this.CylinderRawTypeBox = new System.Windows.Forms.ComboBox();
            this.CylinderRawVOCsOutletBox = new System.Windows.Forms.TextBox();
            this.CylinderRawVOCsInletBox = new System.Windows.Forms.TextBox();
            this.CylinderRawMoistureTB = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.CylinderRawAshTB = new System.Windows.Forms.TextBox();
            this.CylinderRawWeightBox = new System.Windows.Forms.TextBox();
            this.CylinderRawNumberBox = new System.Windows.Forms.TextBox();
            this.CylinderRawQtyPackBox = new System.Windows.Forms.TextBox();
            this.CylinderRawQtyWeightBox = new System.Windows.Forms.TextBox();
            this.CylinderRawReportNoTB = new System.Windows.Forms.TextBox();
            this.CylinderRawLotBox = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.CylinderRawReportNo = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.CylinderRawEffPanel = new System.Windows.Forms.TableLayoutPanel();
            this.label76 = new System.Windows.Forms.Label();
            this.CylinderRawSO2CheckBox = new System.Windows.Forms.CheckBox();
            this.CylinderRawH2SCheckBox = new System.Windows.Forms.CheckBox();
            this.CylinderRawNH3CheckBox = new System.Windows.Forms.CheckBox();
            this.CylinderRawTolueneCheckBox = new System.Windows.Forms.CheckBox();
            this.CylinderRawIPACheckBox = new System.Windows.Forms.CheckBox();
            this.CylinderRawAcetoneCheckBox = new System.Windows.Forms.CheckBox();
            this.CylinderRawValue = new System.Windows.Forms.Label();
            this.CylinderRawSO2 = new System.Windows.Forms.Label();
            this.CylinderRawNH3 = new System.Windows.Forms.Label();
            this.CylinderRawToluene = new System.Windows.Forms.Label();
            this.CylinderRawIPA = new System.Windows.Forms.Label();
            this.CylinderRawAcetone = new System.Windows.Forms.Label();
            this.CylinderRawSO2ConcertrationBox = new System.Windows.Forms.TextBox();
            this.CylinderRawSO2BackGroundBox = new System.Windows.Forms.TextBox();
            this.CylinderRawH2SBackGroundBox = new System.Windows.Forms.TextBox();
            this.CylinderRawNH3BackGroundBox = new System.Windows.Forms.TextBox();
            this.CylinderRawTolueneBackGroundBox = new System.Windows.Forms.TextBox();
            this.CylinderRawIPABackGroundBox = new System.Windows.Forms.TextBox();
            this.CylinderRawAcetoneBackGroundBox = new System.Windows.Forms.TextBox();
            this.CylinderRawAcetoneConcertrationBox = new System.Windows.Forms.TextBox();
            this.CylinderRawIPAConcertrationBox = new System.Windows.Forms.TextBox();
            this.CylinderRawTolueneConcertrationBox = new System.Windows.Forms.TextBox();
            this.CylinderRawNH3ConcertrationBox = new System.Windows.Forms.TextBox();
            this.CylinderRawH2SConcertrationBox = new System.Windows.Forms.TextBox();
            this.CylinderRawSO2ValueBox = new System.Windows.Forms.TextBox();
            this.CylinderRawH2SValueBox = new System.Windows.Forms.TextBox();
            this.CylinderRawGas = new System.Windows.Forms.Label();
            this.CylinderRawNH3ValueBox = new System.Windows.Forms.TextBox();
            this.CylinderRawTolueneValueBox = new System.Windows.Forms.TextBox();
            this.CylinderRawIPAValueBox = new System.Windows.Forms.TextBox();
            this.CylinderRawAcetoneValueBox = new System.Windows.Forms.TextBox();
            this.CylinderRawBackGround = new System.Windows.Forms.Label();
            this.CylinderRawConcertration = new System.Windows.Forms.Label();
            this.CylinderRawH2S = new System.Windows.Forms.Label();
            this.CylinderRawEff = new System.Windows.Forms.Label();
            this.CylinderRawMeshBox = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn12 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn13 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CylinderRawParticle = new System.Windows.Forms.Label();
            this.FilterPage = new System.Windows.Forms.TabPage();
            this.FilterReportCustmorBox = new System.Windows.Forms.ComboBox();
            this.FilterModelBox = new System.Windows.Forms.ComboBox();
            this.FilterProductionBox = new System.Windows.Forms.DateTimePicker();
            this.FilterTestDateBox = new System.Windows.Forms.DateTimePicker();
            this.FilterBox = new System.Windows.Forms.DataGridView();
            this.生產序號 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.重量 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.length = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.width = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.height = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.diagonal = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Particle_In = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Particle_Out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IPA_In = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.IPA_Out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Acetone_In = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Acetone_Out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Nontarget_In = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Nontarget_Out = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Pressure_Drop_spec = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Pressure_Drop = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FilterMaterialNumerBox = new System.Windows.Forms.TextBox();
            this.FilterAlarmBox = new System.Windows.Forms.TextBox();
            this.FilterCarbonLotBox = new System.Windows.Forms.TextBox();
            this.ReFilterNOBox = new System.Windows.Forms.TextBox();
            this.FilterPackageNOBox = new System.Windows.Forms.TextBox();
            this.FilterMaterialNumer = new System.Windows.Forms.Label();
            this.FilterAlarm = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.FilterReportBox = new System.Windows.Forms.TextBox();
            this.ReFilterNO = new System.Windows.Forms.Label();
            this.FilterModel = new System.Windows.Forms.Label();
            this.FilterReportCustmor = new System.Windows.Forms.Label();
            this.FilterPackageNO = new System.Windows.Forms.Label();
            this.FilterProduction = new System.Windows.Forms.Label();
            this.FilterReport = new System.Windows.Forms.Label();
            this.FilterTestDate = new System.Windows.Forms.Label();
            this.FilterInProcessPage = new System.Windows.Forms.TabPage();
            this.FilterInProcessEffPanel = new System.Windows.Forms.TableLayoutPanel();
            this.label35 = new System.Windows.Forms.Label();
            this.FilterInGas = new System.Windows.Forms.Label();
            this.FilterInSO2Concertration = new System.Windows.Forms.Label();
            this.FilterInBackGround = new System.Windows.Forms.Label();
            this.FilterInValue = new System.Windows.Forms.Label();
            this.FilterInSO2 = new System.Windows.Forms.Label();
            this.FilterInH2S = new System.Windows.Forms.Label();
            this.FilterInNH3 = new System.Windows.Forms.Label();
            this.FilterInToluene = new System.Windows.Forms.Label();
            this.FilterInSO2CheckBox = new System.Windows.Forms.CheckBox();
            this.FilterInH2SCheckBox = new System.Windows.Forms.CheckBox();
            this.FilterInNH3CheckBox = new System.Windows.Forms.CheckBox();
            this.FilterInTolueneCheckBox = new System.Windows.Forms.CheckBox();
            this.FilterInSO2ConcentrationBox = new System.Windows.Forms.TextBox();
            this.FilterInH2SConcentrationBox = new System.Windows.Forms.TextBox();
            this.FilterInNH3ConcentrationBox = new System.Windows.Forms.TextBox();
            this.FilterInBackGroundTolueneBox = new System.Windows.Forms.TextBox();
            this.FilterInBackGroundNH3Box = new System.Windows.Forms.TextBox();
            this.FilterInBackGroundH2SBox = new System.Windows.Forms.TextBox();
            this.FilterInBackGroundSO2Box = new System.Windows.Forms.TextBox();
            this.FilterInTolueneConcentrationBox = new System.Windows.Forms.TextBox();
            this.FilterInValueSO2Box = new System.Windows.Forms.TextBox();
            this.FilterInValueH2SBox = new System.Windows.Forms.TextBox();
            this.FilterInValueNH3Box = new System.Windows.Forms.TextBox();
            this.FilterInValueTolueneBox = new System.Windows.Forms.TextBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.FilterInValueIPABox = new System.Windows.Forms.TextBox();
            this.FilterInValueAcetoneBox = new System.Windows.Forms.TextBox();
            this.FilterInIPAConcentrationBox = new System.Windows.Forms.TextBox();
            this.FilterInAcetoneConcentrationBox = new System.Windows.Forms.TextBox();
            this.FilterInBackGroundIPABox = new System.Windows.Forms.TextBox();
            this.FilterInBackGroundAcetoneBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessTypeBox = new System.Windows.Forms.ComboBox();
            this.FilterInProcessTestBox = new System.Windows.Forms.DateTimePicker();
            this.FilterInProcessProductionBox = new System.Windows.Forms.DateTimePicker();
            this.FilterInProcessCarbonInfoBox = new System.Windows.Forms.TextBox();
            this.FilterSizeTB = new System.Windows.Forms.TextBox();
            this.FilterInProcessPressureDropBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessTestGsmBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessWindBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessPressureBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessLowerBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessUpperBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessSpeedBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessGileBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessgsmBox = new System.Windows.Forms.TextBox();
            this.FilterMaterialNOBox = new System.Windows.Forms.TextBox();
            this.FilterInProcessCarbonOrderBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.FilterInProcessReportNOTB = new System.Windows.Forms.TextBox();
            this.FilterInProcessOrderBox = new System.Windows.Forms.TextBox();
            this.FilterSize = new System.Windows.Forms.Label();
            this.FilterInProcessWire = new System.Windows.Forms.Label();
            this.FilterInProcessEff = new System.Windows.Forms.Label();
            this.FilterInProcessTestGsm = new System.Windows.Forms.Label();
            this.FilterInProcessWind = new System.Windows.Forms.Label();
            this.FilterInProcessPressure = new System.Windows.Forms.Label();
            this.FilterInProcessLower = new System.Windows.Forms.Label();
            this.FilterInProcessUpper = new System.Windows.Forms.Label();
            this.FilterInProcessSpeed = new System.Windows.Forms.Label();
            this.FilterInProcessGile = new System.Windows.Forms.Label();
            this.FilterInProcessType = new System.Windows.Forms.Label();
            this.FilterMaterialNO = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.FilterInProcessReportNO = new System.Windows.Forms.Label();
            this.FilterInProcessgsm = new System.Windows.Forms.Label();
            this.FilterInProcessOrder = new System.Windows.Forms.Label();
            this.FilterInProcessTest = new System.Windows.Forms.Label();
            this.FilterInProcessProduction = new System.Windows.Forms.Label();
            this.FilterRawPage = new System.Windows.Forms.TabPage();
            this.FilterRawReportNoTB = new System.Windows.Forms.TextBox();
            this.FilterRawParticleSizeBox = new System.Windows.Forms.DataGridView();
            this.目數 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.weight = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FilterRawArriveDateBox = new System.Windows.Forms.DateTimePicker();
            this.FilterRawTestDateBox = new System.Windows.Forms.DateTimePicker();
            this.FilterRawTypeBox = new System.Windows.Forms.ComboBox();
            this.FilterRawEffvalue = new System.Windows.Forms.Label();
            this.FilterRawBackGround = new System.Windows.Forms.Label();
            this.FilterRawConcertration = new System.Windows.Forms.Label();
            this.FilterRawBackGroundBox = new System.Windows.Forms.TextBox();
            this.FilterRawEffvalueBox = new System.Windows.Forms.TextBox();
            this.FilterRawConcertrationBox = new System.Windows.Forms.TextBox();
            this.FilterRawPressureBox = new System.Windows.Forms.TextBox();
            this.FilterRawVOCsOutletBox = new System.Windows.Forms.TextBox();
            this.FilterRawVOCsInletBox = new System.Windows.Forms.TextBox();
            this.FilterRawWeightBox = new System.Windows.Forms.TextBox();
            this.FilterRawNumberBox = new System.Windows.Forms.TextBox();
            this.FilterRawQuantityBox = new System.Windows.Forms.TextBox();
            this.FilterRawQtyWeightBox = new System.Windows.Forms.TextBox();
            this.FilterRawBatchNOBox = new System.Windows.Forms.TextBox();
            this.FilterRawEff = new System.Windows.Forms.Label();
            this.FilterRawPressure = new System.Windows.Forms.Label();
            this.FilterRawVOCsOutlet = new System.Windows.Forms.Label();
            this.FilterRawVOCsInlet = new System.Windows.Forms.Label();
            this.FilterRawVOCs = new System.Windows.Forms.Label();
            this.FilterRawWeight = new System.Windows.Forms.Label();
            this.FilterRawParticleSize = new System.Windows.Forms.Label();
            this.FilterRawNumber = new System.Windows.Forms.Label();
            this.FilterRawQuantity = new System.Windows.Forms.Label();
            this.FilterRawBatchNO = new System.Windows.Forms.Label();
            this.FilterRawReportNo = new System.Windows.Forms.Label();
            this.FilterRawType = new System.Windows.Forms.Label();
            this.FilterRawArriveDate = new System.Windows.Forms.Label();
            this.FilterRawTestDate = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.RawMaterialPage = new System.Windows.Forms.TabPage();
            this.RawMaterialdgv = new System.Windows.Forms.DataGridView();
            this.MaterialTypeTB = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.MaterialTestDateBox = new System.Windows.Forms.DateTimePicker();
            this.RawMaterialNOtb = new System.Windows.Forms.TextBox();
            this.label23 = new System.Windows.Forms.Label();
            this.MaterialReportNOTB = new System.Windows.Forms.TextBox();
            this.label22 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.execute = new System.Windows.Forms.Button();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CylinderPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CylinderBox)).BeginInit();
            this.CylinderRawPage.SuspendLayout();
            this.CylinderRawEffPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CylinderRawMeshBox)).BeginInit();
            this.FilterPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FilterBox)).BeginInit();
            this.FilterInProcessPage.SuspendLayout();
            this.FilterInProcessEffPanel.SuspendLayout();
            this.FilterRawPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FilterRawParticleSizeBox)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.RawMaterialPage.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RawMaterialdgv)).BeginInit();
            this.SuspendLayout();
            // 
            // CylinderPage
            // 
            this.CylinderPage.BackColor = System.Drawing.Color.White;
            this.CylinderPage.Controls.Add(this.CYLTypeBox);
            this.CylinderPage.Controls.Add(this.CylinderTestDateBox);
            this.CylinderPage.Controls.Add(this.CylinderBox);
            this.CylinderPage.Controls.Add(this.CYLMaterialSNBox);
            this.CylinderPage.Controls.Add(this.ReCylinderNOBox);
            this.CylinderPage.Controls.Add(this.CylinderCustmorBox);
            this.CylinderPage.Controls.Add(this.CylinderNOBox);
            this.CylinderPage.Controls.Add(this.CYLType);
            this.CylinderPage.Controls.Add(this.CylinderMaterialSN);
            this.CylinderPage.Controls.Add(this.CylinderReportNOBox);
            this.CylinderPage.Controls.Add(this.ReCylinderNO);
            this.CylinderPage.Controls.Add(this.CylinderCustmor);
            this.CylinderPage.Controls.Add(this.CylinderNO);
            this.CylinderPage.Controls.Add(this.CylinderReportNO);
            this.CylinderPage.Controls.Add(this.CylinderTestDate);
            this.CylinderPage.Location = new System.Drawing.Point(4, 29);
            this.CylinderPage.Name = "CylinderPage";
            this.CylinderPage.Size = new System.Drawing.Size(1018, 625);
            this.CylinderPage.TabIndex = 4;
            this.CylinderPage.Text = "濾筒成品";
            // 
            // CYLTypeBox
            // 
            this.CYLTypeBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CYLTypeBox.FormattingEnabled = true;
            this.CYLTypeBox.Items.AddRange(new object[] {
            "MA",
            "MB",
            "MC",
            "MA+MB",
            "MA+MC",
            "MB+MC",
            "MA+MB+MC"});
            this.CYLTypeBox.Location = new System.Drawing.Point(175, 225);
            this.CYLTypeBox.Name = "CYLTypeBox";
            this.CYLTypeBox.Size = new System.Drawing.Size(100, 28);
            this.CYLTypeBox.TabIndex = 7;
            // 
            // CylinderTestDateBox
            // 
            this.CylinderTestDateBox.Checked = false;
            this.CylinderTestDateBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderTestDateBox.CustomFormat = "yyyy.MM.dd";
            this.CylinderTestDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.CylinderTestDateBox.Location = new System.Drawing.Point(175, 25);
            this.CylinderTestDateBox.Name = "CylinderTestDateBox";
            this.CylinderTestDateBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderTestDateBox.TabIndex = 1;
            this.CylinderTestDateBox.ValueChanged += new System.EventHandler(this.CylinderTestDateBox_ValueChanged);
            // 
            // CylinderBox
            // 
            this.CylinderBox.AllowUserToResizeColumns = false;
            this.CylinderBox.AllowUserToResizeRows = false;
            this.CylinderBox.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CylinderBox.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CYLSN,
            this.CYLWeight,
            this.CYL_Particle_In,
            this.CYL_Particle_out,
            this.CYL_IPA_in,
            this.CYL_IPA_out,
            this.CYL_Acetone_In,
            this.CYL_Acetone_out,
            this.CYL_Nontarget_in,
            this.CYL_Nontarget_out,
            this.CYL_Pressure_Drop});
            this.CylinderBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderBox.Location = new System.Drawing.Point(40, 310);
            this.CylinderBox.Name = "CylinderBox";
            this.CylinderBox.RowTemplate.Height = 24;
            this.CylinderBox.Size = new System.Drawing.Size(915, 270);
            this.CylinderBox.TabIndex = 9;
            this.CylinderBox.RowPostPaint += new System.Windows.Forms.DataGridViewRowPostPaintEventHandler(this.CylinderBox_RowPostPaint);
            // 
            // CYLSN
            // 
            this.CYLSN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.CYLSN.Frozen = true;
            this.CYLSN.HeaderText = "生產序號";
            this.CYLSN.Name = "CYLSN";
            // 
            // CYLWeight
            // 
            this.CYLWeight.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.CYLWeight.HeaderText = "重量";
            this.CYLWeight.Name = "CYLWeight";
            this.CYLWeight.Width = 65;
            // 
            // CYL_Particle_In
            // 
            this.CYL_Particle_In.HeaderText = "Particle_In";
            this.CYL_Particle_In.Name = "CYL_Particle_In";
            this.CYL_Particle_In.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // CYL_Particle_out
            // 
            this.CYL_Particle_out.HeaderText = "Particle_Out";
            this.CYL_Particle_out.Name = "CYL_Particle_out";
            this.CYL_Particle_out.Width = 105;
            // 
            // CYL_IPA_in
            // 
            this.CYL_IPA_in.HeaderText = "IPA_In";
            this.CYL_IPA_in.Name = "CYL_IPA_in";
            this.CYL_IPA_in.Width = 65;
            // 
            // CYL_IPA_out
            // 
            this.CYL_IPA_out.HeaderText = "IPA_Out";
            this.CYL_IPA_out.Name = "CYL_IPA_out";
            this.CYL_IPA_out.Width = 78;
            // 
            // CYL_Acetone_In
            // 
            this.CYL_Acetone_In.HeaderText = "Acetone_In";
            this.CYL_Acetone_In.Name = "CYL_Acetone_In";
            // 
            // CYL_Acetone_out
            // 
            this.CYL_Acetone_out.HeaderText = "Acetone_Out";
            this.CYL_Acetone_out.Name = "CYL_Acetone_out";
            this.CYL_Acetone_out.Width = 120;
            // 
            // CYL_Nontarget_in
            // 
            this.CYL_Nontarget_in.HeaderText = "Nontarget_In";
            this.CYL_Nontarget_in.Name = "CYL_Nontarget_in";
            this.CYL_Nontarget_in.Width = 125;
            // 
            // CYL_Nontarget_out
            // 
            this.CYL_Nontarget_out.HeaderText = "Nontarget_Out";
            this.CYL_Nontarget_out.Name = "CYL_Nontarget_out";
            this.CYL_Nontarget_out.Width = 130;
            // 
            // CYL_Pressure_Drop
            // 
            this.CYL_Pressure_Drop.HeaderText = "Pressure_Drop";
            this.CYL_Pressure_Drop.Name = "CYL_Pressure_Drop";
            this.CYL_Pressure_Drop.Width = 140;
            // 
            // CYLMaterialSNBox
            // 
            this.CYLMaterialSNBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CYLMaterialSNBox.Location = new System.Drawing.Point(175, 265);
            this.CYLMaterialSNBox.Name = "CYLMaterialSNBox";
            this.CYLMaterialSNBox.Size = new System.Drawing.Size(100, 29);
            this.CYLMaterialSNBox.TabIndex = 8;
            // 
            // ReCylinderNOBox
            // 
            this.ReCylinderNOBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.ReCylinderNOBox.Location = new System.Drawing.Point(175, 185);
            this.ReCylinderNOBox.Name = "ReCylinderNOBox";
            this.ReCylinderNOBox.Size = new System.Drawing.Size(100, 29);
            this.ReCylinderNOBox.TabIndex = 6;
            // 
            // CylinderCustmorBox
            // 
            this.CylinderCustmorBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderCustmorBox.Location = new System.Drawing.Point(175, 145);
            this.CylinderCustmorBox.Name = "CylinderCustmorBox";
            this.CylinderCustmorBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderCustmorBox.TabIndex = 4;
            this.CylinderCustmorBox.Text = "General";
            // 
            // CylinderNOBox
            // 
            this.CylinderNOBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderNOBox.Location = new System.Drawing.Point(175, 107);
            this.CylinderNOBox.Name = "CylinderNOBox";
            this.CylinderNOBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderNOBox.TabIndex = 3;
            // 
            // CYLType
            // 
            this.CYLType.AutoSize = true;
            this.CYLType.Location = new System.Drawing.Point(40, 228);
            this.CYLType.Name = "CYLType";
            this.CYLType.Size = new System.Drawing.Size(73, 20);
            this.CYLType.TabIndex = 13;
            this.CYLType.Text = "原料種類";
            // 
            // CylinderMaterialSN
            // 
            this.CylinderMaterialSN.AutoSize = true;
            this.CylinderMaterialSN.Location = new System.Drawing.Point(40, 268);
            this.CylinderMaterialSN.Name = "CylinderMaterialSN";
            this.CylinderMaterialSN.Size = new System.Drawing.Size(73, 20);
            this.CylinderMaterialSN.TabIndex = 13;
            this.CylinderMaterialSN.Text = "原料批號";
            // 
            // CylinderReportNOBox
            // 
            this.CylinderReportNOBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderReportNOBox.Location = new System.Drawing.Point(175, 67);
            this.CylinderReportNOBox.Name = "CylinderReportNOBox";
            this.CylinderReportNOBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderReportNOBox.TabIndex = 2;
            // 
            // ReCylinderNO
            // 
            this.ReCylinderNO.AutoSize = true;
            this.ReCylinderNO.Location = new System.Drawing.Point(40, 188);
            this.ReCylinderNO.Name = "ReCylinderNO";
            this.ReCylinderNO.Size = new System.Drawing.Size(73, 20);
            this.ReCylinderNO.TabIndex = 13;
            this.ReCylinderNO.Text = "再生次數";
            // 
            // CylinderCustmor
            // 
            this.CylinderCustmor.AutoSize = true;
            this.CylinderCustmor.Location = new System.Drawing.Point(40, 148);
            this.CylinderCustmor.Name = "CylinderCustmor";
            this.CylinderCustmor.Size = new System.Drawing.Size(41, 20);
            this.CylinderCustmor.TabIndex = 16;
            this.CylinderCustmor.Text = "客戶";
            // 
            // CylinderNO
            // 
            this.CylinderNO.AutoSize = true;
            this.CylinderNO.Location = new System.Drawing.Point(40, 110);
            this.CylinderNO.Name = "CylinderNO";
            this.CylinderNO.Size = new System.Drawing.Size(73, 20);
            this.CylinderNO.TabIndex = 17;
            this.CylinderNO.Text = "生產單號";
            // 
            // CylinderReportNO
            // 
            this.CylinderReportNO.AutoSize = true;
            this.CylinderReportNO.Location = new System.Drawing.Point(40, 70);
            this.CylinderReportNO.Name = "CylinderReportNO";
            this.CylinderReportNO.Size = new System.Drawing.Size(73, 20);
            this.CylinderReportNO.TabIndex = 18;
            this.CylinderReportNO.Text = "報告編號";
            // 
            // CylinderTestDate
            // 
            this.CylinderTestDate.AutoSize = true;
            this.CylinderTestDate.Location = new System.Drawing.Point(40, 30);
            this.CylinderTestDate.Name = "CylinderTestDate";
            this.CylinderTestDate.Size = new System.Drawing.Size(73, 20);
            this.CylinderTestDate.TabIndex = 19;
            this.CylinderTestDate.Text = "測試日期";
            // 
            // CylinderRawPage
            // 
            this.CylinderRawPage.BackColor = System.Drawing.Color.White;
            this.CylinderRawPage.Controls.Add(this.chkAsh);
            this.CylinderRawPage.Controls.Add(this.checkBox4);
            this.CylinderRawPage.Controls.Add(this.chkMoisture);
            this.CylinderRawPage.Controls.Add(this.CylinderRawPressureBox);
            this.CylinderRawPage.Controls.Add(this.label20);
            this.CylinderRawPage.Controls.Add(this.CylinderRawArriveDateBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawTestDateBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawTypeBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawVOCsOutletBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawVOCsInletBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawMoistureTB);
            this.CylinderRawPage.Controls.Add(this.textBox5);
            this.CylinderRawPage.Controls.Add(this.CylinderRawAshTB);
            this.CylinderRawPage.Controls.Add(this.CylinderRawWeightBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawNumberBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawQtyPackBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawQtyWeightBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawReportNoTB);
            this.CylinderRawPage.Controls.Add(this.CylinderRawLotBox);
            this.CylinderRawPage.Controls.Add(this.label6);
            this.CylinderRawPage.Controls.Add(this.label7);
            this.CylinderRawPage.Controls.Add(this.label8);
            this.CylinderRawPage.Controls.Add(this.label9);
            this.CylinderRawPage.Controls.Add(this.label10);
            this.CylinderRawPage.Controls.Add(this.label11);
            this.CylinderRawPage.Controls.Add(this.label13);
            this.CylinderRawPage.Controls.Add(this.label14);
            this.CylinderRawPage.Controls.Add(this.CylinderRawReportNo);
            this.CylinderRawPage.Controls.Add(this.label15);
            this.CylinderRawPage.Controls.Add(this.label16);
            this.CylinderRawPage.Controls.Add(this.label17);
            this.CylinderRawPage.Controls.Add(this.label18);
            this.CylinderRawPage.Controls.Add(this.label19);
            this.CylinderRawPage.Controls.Add(this.CylinderRawEffPanel);
            this.CylinderRawPage.Controls.Add(this.CylinderRawEff);
            this.CylinderRawPage.Controls.Add(this.CylinderRawMeshBox);
            this.CylinderRawPage.Controls.Add(this.CylinderRawParticle);
            this.CylinderRawPage.Location = new System.Drawing.Point(4, 29);
            this.CylinderRawPage.Name = "CylinderRawPage";
            this.CylinderRawPage.Size = new System.Drawing.Size(1018, 625);
            this.CylinderRawPage.TabIndex = 3;
            this.CylinderRawPage.Text = "濾筒原料";
            // 
            // chkAsh
            // 
            this.chkAsh.AutoSize = true;
            this.chkAsh.Location = new System.Drawing.Point(17, 555);
            this.chkAsh.Name = "chkAsh";
            this.chkAsh.Size = new System.Drawing.Size(15, 14);
            this.chkAsh.TabIndex = 53;
            this.chkAsh.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(18, 512);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(15, 14);
            this.checkBox4.TabIndex = 53;
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // chkMoisture
            // 
            this.chkMoisture.AutoSize = true;
            this.chkMoisture.Location = new System.Drawing.Point(18, 469);
            this.chkMoisture.Name = "chkMoisture";
            this.chkMoisture.Size = new System.Drawing.Size(15, 14);
            this.chkMoisture.TabIndex = 53;
            this.chkMoisture.UseVisualStyleBackColor = true;
            // 
            // CylinderRawPressureBox
            // 
            this.CylinderRawPressureBox.Location = new System.Drawing.Point(174, 427);
            this.CylinderRawPressureBox.Name = "CylinderRawPressureBox";
            this.CylinderRawPressureBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawPressureBox.TabIndex = 47;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(39, 430);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(41, 20);
            this.label20.TabIndex = 51;
            this.label20.Text = "壓損";
            // 
            // CylinderRawArriveDateBox
            // 
            this.CylinderRawArriveDateBox.Checked = false;
            this.CylinderRawArriveDateBox.CustomFormat = "yyyy.MM.dd";
            this.CylinderRawArriveDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.CylinderRawArriveDateBox.Location = new System.Drawing.Point(175, 66);
            this.CylinderRawArriveDateBox.Name = "CylinderRawArriveDateBox";
            this.CylinderRawArriveDateBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawArriveDateBox.TabIndex = 38;
            // 
            // CylinderRawTestDateBox
            // 
            this.CylinderRawTestDateBox.Checked = false;
            this.CylinderRawTestDateBox.CustomFormat = "yyyy.MM.dd";
            this.CylinderRawTestDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.CylinderRawTestDateBox.Location = new System.Drawing.Point(175, 26);
            this.CylinderRawTestDateBox.Name = "CylinderRawTestDateBox";
            this.CylinderRawTestDateBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawTestDateBox.TabIndex = 37;
            this.CylinderRawTestDateBox.ValueChanged += new System.EventHandler(this.CylinderRawTestDateBox_ValueChange);
            // 
            // CylinderRawTypeBox
            // 
            this.CylinderRawTypeBox.FormattingEnabled = true;
            this.CylinderRawTypeBox.Items.AddRange(new object[] {
            "IKP201",
            "IKP205",
            "IKP210",
            "SG017_A",
            "SG017_B",
            "SG017_C",
            "SG017_D",
            "SG029",
            "SG043",
            "CI001"});
            this.CylinderRawTypeBox.Location = new System.Drawing.Point(174, 135);
            this.CylinderRawTypeBox.Name = "CylinderRawTypeBox";
            this.CylinderRawTypeBox.Size = new System.Drawing.Size(100, 28);
            this.CylinderRawTypeBox.TabIndex = 39;
            this.CylinderRawTypeBox.SelectedIndexChanged += new System.EventHandler(this.CylinderRawTypeBox_SelectedIndexChanged);
            // 
            // CylinderRawVOCsOutletBox
            // 
            this.CylinderRawVOCsOutletBox.Location = new System.Drawing.Point(174, 381);
            this.CylinderRawVOCsOutletBox.Name = "CylinderRawVOCsOutletBox";
            this.CylinderRawVOCsOutletBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawVOCsOutletBox.TabIndex = 46;
            // 
            // CylinderRawVOCsInletBox
            // 
            this.CylinderRawVOCsInletBox.Location = new System.Drawing.Point(174, 341);
            this.CylinderRawVOCsInletBox.Name = "CylinderRawVOCsInletBox";
            this.CylinderRawVOCsInletBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawVOCsInletBox.TabIndex = 45;
            // 
            // CylinderRawMoistureTB
            // 
            this.CylinderRawMoistureTB.Location = new System.Drawing.Point(174, 462);
            this.CylinderRawMoistureTB.Name = "CylinderRawMoistureTB";
            this.CylinderRawMoistureTB.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawMoistureTB.TabIndex = 48;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(174, 505);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(100, 29);
            this.textBox5.TabIndex = 49;
            // 
            // CylinderRawAshTB
            // 
            this.CylinderRawAshTB.Location = new System.Drawing.Point(174, 548);
            this.CylinderRawAshTB.Name = "CylinderRawAshTB";
            this.CylinderRawAshTB.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawAshTB.TabIndex = 50;
            // 
            // CylinderRawWeightBox
            // 
            this.CylinderRawWeightBox.Location = new System.Drawing.Point(174, 297);
            this.CylinderRawWeightBox.Name = "CylinderRawWeightBox";
            this.CylinderRawWeightBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawWeightBox.TabIndex = 44;
            // 
            // CylinderRawNumberBox
            // 
            this.CylinderRawNumberBox.Location = new System.Drawing.Point(174, 255);
            this.CylinderRawNumberBox.Name = "CylinderRawNumberBox";
            this.CylinderRawNumberBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawNumberBox.TabIndex = 43;
            // 
            // CylinderRawQtyPackBox
            // 
            this.CylinderRawQtyPackBox.Location = new System.Drawing.Point(224, 215);
            this.CylinderRawQtyPackBox.Name = "CylinderRawQtyPackBox";
            this.CylinderRawQtyPackBox.Size = new System.Drawing.Size(50, 29);
            this.CylinderRawQtyPackBox.TabIndex = 42;
            // 
            // CylinderRawQtyWeightBox
            // 
            this.CylinderRawQtyWeightBox.Location = new System.Drawing.Point(174, 215);
            this.CylinderRawQtyWeightBox.Name = "CylinderRawQtyWeightBox";
            this.CylinderRawQtyWeightBox.Size = new System.Drawing.Size(50, 29);
            this.CylinderRawQtyWeightBox.TabIndex = 41;
            // 
            // CylinderRawReportNoTB
            // 
            this.CylinderRawReportNoTB.Location = new System.Drawing.Point(174, 100);
            this.CylinderRawReportNoTB.Name = "CylinderRawReportNoTB";
            this.CylinderRawReportNoTB.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawReportNoTB.TabIndex = 40;
            // 
            // CylinderRawLotBox
            // 
            this.CylinderRawLotBox.Location = new System.Drawing.Point(174, 175);
            this.CylinderRawLotBox.Name = "CylinderRawLotBox";
            this.CylinderRawLotBox.Size = new System.Drawing.Size(100, 29);
            this.CylinderRawLotBox.TabIndex = 40;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(99, 384);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(57, 20);
            this.label6.TabIndex = 36;
            this.label6.Text = "Outlet";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(99, 344);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 20);
            this.label7.TabIndex = 34;
            this.label7.Text = "Inlet";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(39, 366);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(51, 20);
            this.label8.TabIndex = 33;
            this.label8.Text = "VOCs";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(39, 465);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(57, 20);
            this.label9.TabIndex = 32;
            this.label9.Text = "含水率";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(39, 508);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(57, 20);
            this.label10.TabIndex = 31;
            this.label10.Text = "正丁烷";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(39, 551);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(41, 20);
            this.label11.TabIndex = 30;
            this.label11.Text = "灰分";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(39, 300);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(89, 20);
            this.label13.TabIndex = 28;
            this.label13.Text = "測試品重量";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(39, 250);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(109, 40);
            this.label14.TabIndex = 27;
            this.label14.Text = "原料編號\n(輸入#後兩位)";
            // 
            // CylinderRawReportNo
            // 
            this.CylinderRawReportNo.AutoSize = true;
            this.CylinderRawReportNo.Location = new System.Drawing.Point(39, 105);
            this.CylinderRawReportNo.Name = "CylinderRawReportNo";
            this.CylinderRawReportNo.Size = new System.Drawing.Size(73, 20);
            this.CylinderRawReportNo.TabIndex = 25;
            this.CylinderRawReportNo.Text = "報告編號";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(39, 220);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(73, 20);
            this.label15.TabIndex = 26;
            this.label15.Text = "進貨數量";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(39, 180);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(89, 20);
            this.label16.TabIndex = 25;
            this.label16.Text = "供應商批號";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(39, 140);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(73, 20);
            this.label17.TabIndex = 24;
            this.label17.Text = "原料種類";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(39, 70);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(73, 20);
            this.label18.TabIndex = 35;
            this.label18.Text = "到廠日期";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(39, 30);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(73, 20);
            this.label19.TabIndex = 23;
            this.label19.Text = "測試日期";
            // 
            // CylinderRawEffPanel
            // 
            this.CylinderRawEffPanel.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.CylinderRawEffPanel.ColumnCount = 5;
            this.CylinderRawEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 7.086237F));
            this.CylinderRawEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.32365F));
            this.CylinderRawEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.32365F));
            this.CylinderRawEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.32365F));
            this.CylinderRawEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 31.94281F));
            this.CylinderRawEffPanel.Controls.Add(this.label76, 0, 0);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawSO2CheckBox, 0, 1);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawH2SCheckBox, 0, 2);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawNH3CheckBox, 0, 3);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawTolueneCheckBox, 0, 4);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawIPACheckBox, 0, 5);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawAcetoneCheckBox, 0, 6);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawValue, 4, 0);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawSO2, 1, 1);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawNH3, 1, 3);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawToluene, 1, 4);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawIPA, 1, 5);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawAcetone, 1, 6);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawSO2ConcertrationBox, 2, 1);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawSO2BackGroundBox, 3, 1);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawH2SBackGroundBox, 3, 2);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawNH3BackGroundBox, 3, 3);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawTolueneBackGroundBox, 3, 4);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawIPABackGroundBox, 3, 5);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawAcetoneBackGroundBox, 3, 6);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawAcetoneConcertrationBox, 2, 6);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawIPAConcertrationBox, 2, 5);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawTolueneConcertrationBox, 2, 4);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawNH3ConcertrationBox, 2, 3);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawH2SConcertrationBox, 2, 2);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawSO2ValueBox, 4, 1);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawH2SValueBox, 4, 2);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawGas, 1, 0);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawNH3ValueBox, 4, 3);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawTolueneValueBox, 4, 4);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawIPAValueBox, 4, 5);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawAcetoneValueBox, 4, 6);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawBackGround, 3, 0);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawConcertration, 2, 0);
            this.CylinderRawEffPanel.Controls.Add(this.CylinderRawH2S, 1, 2);
            this.CylinderRawEffPanel.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderRawEffPanel.Location = new System.Drawing.Point(568, 64);
            this.CylinderRawEffPanel.Name = "CylinderRawEffPanel";
            this.CylinderRawEffPanel.RowCount = 7;
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5.882353F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 13.97059F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.95988F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.95988F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.95988F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.95988F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.95988F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CylinderRawEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.CylinderRawEffPanel.Size = new System.Drawing.Size(407, 409);
            this.CylinderRawEffPanel.TabIndex = 22;
            // 
            // label76
            // 
            this.label76.AutoSize = true;
            this.label76.Location = new System.Drawing.Point(4, 1);
            this.label76.Name = "label76";
            this.label76.Size = new System.Drawing.Size(20, 20);
            this.label76.TabIndex = 20;
            this.label76.Text = "V";
            this.label76.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // CylinderRawSO2CheckBox
            // 
            this.CylinderRawSO2CheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawSO2CheckBox.AutoSize = true;
            this.CylinderRawSO2CheckBox.Location = new System.Drawing.Point(7, 28);
            this.CylinderRawSO2CheckBox.Name = "CylinderRawSO2CheckBox";
            this.CylinderRawSO2CheckBox.Size = new System.Drawing.Size(15, 50);
            this.CylinderRawSO2CheckBox.TabIndex = 13;
            this.CylinderRawSO2CheckBox.Tag = "SO2";
            this.CylinderRawSO2CheckBox.UseVisualStyleBackColor = true;
            // 
            // CylinderRawH2SCheckBox
            // 
            this.CylinderRawH2SCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawH2SCheckBox.AutoSize = true;
            this.CylinderRawH2SCheckBox.Location = new System.Drawing.Point(7, 85);
            this.CylinderRawH2SCheckBox.Name = "CylinderRawH2SCheckBox";
            this.CylinderRawH2SCheckBox.Size = new System.Drawing.Size(15, 58);
            this.CylinderRawH2SCheckBox.TabIndex = 17;
            this.CylinderRawH2SCheckBox.Tag = "H2S";
            this.CylinderRawH2SCheckBox.UseVisualStyleBackColor = true;
            // 
            // CylinderRawNH3CheckBox
            // 
            this.CylinderRawNH3CheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawNH3CheckBox.AutoSize = true;
            this.CylinderRawNH3CheckBox.Location = new System.Drawing.Point(7, 150);
            this.CylinderRawNH3CheckBox.Name = "CylinderRawNH3CheckBox";
            this.CylinderRawNH3CheckBox.Size = new System.Drawing.Size(15, 58);
            this.CylinderRawNH3CheckBox.TabIndex = 21;
            this.CylinderRawNH3CheckBox.Tag = "NH3";
            this.CylinderRawNH3CheckBox.UseVisualStyleBackColor = true;
            // 
            // CylinderRawTolueneCheckBox
            // 
            this.CylinderRawTolueneCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawTolueneCheckBox.AutoSize = true;
            this.CylinderRawTolueneCheckBox.Location = new System.Drawing.Point(7, 215);
            this.CylinderRawTolueneCheckBox.Name = "CylinderRawTolueneCheckBox";
            this.CylinderRawTolueneCheckBox.Size = new System.Drawing.Size(15, 58);
            this.CylinderRawTolueneCheckBox.TabIndex = 25;
            this.CylinderRawTolueneCheckBox.Tag = "Toluene";
            this.CylinderRawTolueneCheckBox.UseVisualStyleBackColor = true;
            // 
            // CylinderRawIPACheckBox
            // 
            this.CylinderRawIPACheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawIPACheckBox.AutoSize = true;
            this.CylinderRawIPACheckBox.Location = new System.Drawing.Point(7, 280);
            this.CylinderRawIPACheckBox.Name = "CylinderRawIPACheckBox";
            this.CylinderRawIPACheckBox.Size = new System.Drawing.Size(15, 58);
            this.CylinderRawIPACheckBox.TabIndex = 29;
            this.CylinderRawIPACheckBox.Tag = "IPA";
            this.CylinderRawIPACheckBox.UseVisualStyleBackColor = true;
            // 
            // CylinderRawAcetoneCheckBox
            // 
            this.CylinderRawAcetoneCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawAcetoneCheckBox.AutoSize = true;
            this.CylinderRawAcetoneCheckBox.Location = new System.Drawing.Point(7, 345);
            this.CylinderRawAcetoneCheckBox.Name = "CylinderRawAcetoneCheckBox";
            this.CylinderRawAcetoneCheckBox.Size = new System.Drawing.Size(15, 60);
            this.CylinderRawAcetoneCheckBox.TabIndex = 33;
            this.CylinderRawAcetoneCheckBox.Tag = "Acetone";
            this.CylinderRawAcetoneCheckBox.UseVisualStyleBackColor = true;
            // 
            // CylinderRawValue
            // 
            this.CylinderRawValue.AutoSize = true;
            this.CylinderRawValue.Location = new System.Drawing.Point(279, 1);
            this.CylinderRawValue.Name = "CylinderRawValue";
            this.CylinderRawValue.Size = new System.Drawing.Size(41, 20);
            this.CylinderRawValue.TabIndex = 23;
            this.CylinderRawValue.Text = "讀值";
            // 
            // CylinderRawSO2
            // 
            this.CylinderRawSO2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawSO2.AutoSize = true;
            this.CylinderRawSO2.Location = new System.Drawing.Point(33, 43);
            this.CylinderRawSO2.Name = "CylinderRawSO2";
            this.CylinderRawSO2.Size = new System.Drawing.Size(75, 20);
            this.CylinderRawSO2.TabIndex = 20;
            this.CylinderRawSO2.Tag = "SO2";
            this.CylinderRawSO2.Text = "SO2";
            this.CylinderRawSO2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // CylinderRawNH3
            // 
            this.CylinderRawNH3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.CylinderRawNH3.AutoSize = true;
            this.CylinderRawNH3.Location = new System.Drawing.Point(49, 169);
            this.CylinderRawNH3.Name = "CylinderRawNH3";
            this.CylinderRawNH3.Size = new System.Drawing.Size(43, 20);
            this.CylinderRawNH3.TabIndex = 20;
            this.CylinderRawNH3.Tag = "NH3";
            this.CylinderRawNH3.Text = "NH3";
            // 
            // CylinderRawToluene
            // 
            this.CylinderRawToluene.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawToluene.AutoSize = true;
            this.CylinderRawToluene.Location = new System.Drawing.Point(33, 234);
            this.CylinderRawToluene.Name = "CylinderRawToluene";
            this.CylinderRawToluene.Size = new System.Drawing.Size(75, 20);
            this.CylinderRawToluene.TabIndex = 20;
            this.CylinderRawToluene.Tag = "Toluene";
            this.CylinderRawToluene.Text = "Toluene";
            // 
            // CylinderRawIPA
            // 
            this.CylinderRawIPA.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.CylinderRawIPA.AutoSize = true;
            this.CylinderRawIPA.ImageAlign = System.Drawing.ContentAlignment.TopRight;
            this.CylinderRawIPA.Location = new System.Drawing.Point(53, 299);
            this.CylinderRawIPA.Name = "CylinderRawIPA";
            this.CylinderRawIPA.Size = new System.Drawing.Size(35, 20);
            this.CylinderRawIPA.TabIndex = 20;
            this.CylinderRawIPA.Tag = "IPA";
            this.CylinderRawIPA.Text = "IPA";
            // 
            // CylinderRawAcetone
            // 
            this.CylinderRawAcetone.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawAcetone.AutoSize = true;
            this.CylinderRawAcetone.Location = new System.Drawing.Point(33, 365);
            this.CylinderRawAcetone.Name = "CylinderRawAcetone";
            this.CylinderRawAcetone.Size = new System.Drawing.Size(75, 20);
            this.CylinderRawAcetone.TabIndex = 20;
            this.CylinderRawAcetone.Tag = "Acetone";
            this.CylinderRawAcetone.Text = "Acetone";
            // 
            // CylinderRawSO2ConcertrationBox
            // 
            this.CylinderRawSO2ConcertrationBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawSO2ConcertrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawSO2ConcertrationBox.Location = new System.Drawing.Point(112, 25);
            this.CylinderRawSO2ConcertrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawSO2ConcertrationBox.Name = "CylinderRawSO2ConcertrationBox";
            this.CylinderRawSO2ConcertrationBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawSO2ConcertrationBox.TabIndex = 14;
            // 
            // CylinderRawSO2BackGroundBox
            // 
            this.CylinderRawSO2BackGroundBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawSO2BackGroundBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawSO2BackGroundBox.Location = new System.Drawing.Point(194, 25);
            this.CylinderRawSO2BackGroundBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawSO2BackGroundBox.Name = "CylinderRawSO2BackGroundBox";
            this.CylinderRawSO2BackGroundBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawSO2BackGroundBox.TabIndex = 15;
            // 
            // CylinderRawH2SBackGroundBox
            // 
            this.CylinderRawH2SBackGroundBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawH2SBackGroundBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawH2SBackGroundBox.Location = new System.Drawing.Point(194, 82);
            this.CylinderRawH2SBackGroundBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawH2SBackGroundBox.Name = "CylinderRawH2SBackGroundBox";
            this.CylinderRawH2SBackGroundBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawH2SBackGroundBox.TabIndex = 19;
            // 
            // CylinderRawNH3BackGroundBox
            // 
            this.CylinderRawNH3BackGroundBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawNH3BackGroundBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawNH3BackGroundBox.Location = new System.Drawing.Point(194, 147);
            this.CylinderRawNH3BackGroundBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawNH3BackGroundBox.Name = "CylinderRawNH3BackGroundBox";
            this.CylinderRawNH3BackGroundBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawNH3BackGroundBox.TabIndex = 23;
            // 
            // CylinderRawTolueneBackGroundBox
            // 
            this.CylinderRawTolueneBackGroundBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawTolueneBackGroundBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawTolueneBackGroundBox.Location = new System.Drawing.Point(194, 212);
            this.CylinderRawTolueneBackGroundBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawTolueneBackGroundBox.Name = "CylinderRawTolueneBackGroundBox";
            this.CylinderRawTolueneBackGroundBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawTolueneBackGroundBox.TabIndex = 27;
            // 
            // CylinderRawIPABackGroundBox
            // 
            this.CylinderRawIPABackGroundBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawIPABackGroundBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawIPABackGroundBox.Location = new System.Drawing.Point(194, 277);
            this.CylinderRawIPABackGroundBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawIPABackGroundBox.Name = "CylinderRawIPABackGroundBox";
            this.CylinderRawIPABackGroundBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawIPABackGroundBox.TabIndex = 31;
            // 
            // CylinderRawAcetoneBackGroundBox
            // 
            this.CylinderRawAcetoneBackGroundBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawAcetoneBackGroundBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawAcetoneBackGroundBox.Location = new System.Drawing.Point(194, 342);
            this.CylinderRawAcetoneBackGroundBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawAcetoneBackGroundBox.Name = "CylinderRawAcetoneBackGroundBox";
            this.CylinderRawAcetoneBackGroundBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawAcetoneBackGroundBox.TabIndex = 35;
            // 
            // CylinderRawAcetoneConcertrationBox
            // 
            this.CylinderRawAcetoneConcertrationBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawAcetoneConcertrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawAcetoneConcertrationBox.Location = new System.Drawing.Point(112, 342);
            this.CylinderRawAcetoneConcertrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawAcetoneConcertrationBox.Name = "CylinderRawAcetoneConcertrationBox";
            this.CylinderRawAcetoneConcertrationBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawAcetoneConcertrationBox.TabIndex = 34;
            // 
            // CylinderRawIPAConcertrationBox
            // 
            this.CylinderRawIPAConcertrationBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawIPAConcertrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawIPAConcertrationBox.Location = new System.Drawing.Point(112, 277);
            this.CylinderRawIPAConcertrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawIPAConcertrationBox.Name = "CylinderRawIPAConcertrationBox";
            this.CylinderRawIPAConcertrationBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawIPAConcertrationBox.TabIndex = 30;
            // 
            // CylinderRawTolueneConcertrationBox
            // 
            this.CylinderRawTolueneConcertrationBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawTolueneConcertrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawTolueneConcertrationBox.Location = new System.Drawing.Point(112, 212);
            this.CylinderRawTolueneConcertrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawTolueneConcertrationBox.Name = "CylinderRawTolueneConcertrationBox";
            this.CylinderRawTolueneConcertrationBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawTolueneConcertrationBox.TabIndex = 26;
            // 
            // CylinderRawNH3ConcertrationBox
            // 
            this.CylinderRawNH3ConcertrationBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawNH3ConcertrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawNH3ConcertrationBox.Location = new System.Drawing.Point(112, 147);
            this.CylinderRawNH3ConcertrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawNH3ConcertrationBox.Name = "CylinderRawNH3ConcertrationBox";
            this.CylinderRawNH3ConcertrationBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawNH3ConcertrationBox.TabIndex = 22;
            // 
            // CylinderRawH2SConcertrationBox
            // 
            this.CylinderRawH2SConcertrationBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawH2SConcertrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawH2SConcertrationBox.Location = new System.Drawing.Point(112, 82);
            this.CylinderRawH2SConcertrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawH2SConcertrationBox.Name = "CylinderRawH2SConcertrationBox";
            this.CylinderRawH2SConcertrationBox.Size = new System.Drawing.Size(81, 22);
            this.CylinderRawH2SConcertrationBox.TabIndex = 18;
            // 
            // CylinderRawSO2ValueBox
            // 
            this.CylinderRawSO2ValueBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawSO2ValueBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawSO2ValueBox.Location = new System.Drawing.Point(276, 25);
            this.CylinderRawSO2ValueBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawSO2ValueBox.Multiline = true;
            this.CylinderRawSO2ValueBox.Name = "CylinderRawSO2ValueBox";
            this.CylinderRawSO2ValueBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CylinderRawSO2ValueBox.Size = new System.Drawing.Size(130, 56);
            this.CylinderRawSO2ValueBox.TabIndex = 16;
            // 
            // CylinderRawH2SValueBox
            // 
            this.CylinderRawH2SValueBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawH2SValueBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawH2SValueBox.Location = new System.Drawing.Point(276, 82);
            this.CylinderRawH2SValueBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawH2SValueBox.Multiline = true;
            this.CylinderRawH2SValueBox.Name = "CylinderRawH2SValueBox";
            this.CylinderRawH2SValueBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CylinderRawH2SValueBox.Size = new System.Drawing.Size(130, 64);
            this.CylinderRawH2SValueBox.TabIndex = 20;
            // 
            // CylinderRawGas
            // 
            this.CylinderRawGas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.CylinderRawGas.Location = new System.Drawing.Point(50, 1);
            this.CylinderRawGas.Name = "CylinderRawGas";
            this.CylinderRawGas.Size = new System.Drawing.Size(41, 23);
            this.CylinderRawGas.TabIndex = 0;
            this.CylinderRawGas.Text = "氣體";
            // 
            // CylinderRawNH3ValueBox
            // 
            this.CylinderRawNH3ValueBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawNH3ValueBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawNH3ValueBox.Location = new System.Drawing.Point(276, 147);
            this.CylinderRawNH3ValueBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawNH3ValueBox.Multiline = true;
            this.CylinderRawNH3ValueBox.Name = "CylinderRawNH3ValueBox";
            this.CylinderRawNH3ValueBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CylinderRawNH3ValueBox.Size = new System.Drawing.Size(130, 64);
            this.CylinderRawNH3ValueBox.TabIndex = 24;
            // 
            // CylinderRawTolueneValueBox
            // 
            this.CylinderRawTolueneValueBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawTolueneValueBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawTolueneValueBox.Location = new System.Drawing.Point(276, 212);
            this.CylinderRawTolueneValueBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawTolueneValueBox.Multiline = true;
            this.CylinderRawTolueneValueBox.Name = "CylinderRawTolueneValueBox";
            this.CylinderRawTolueneValueBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CylinderRawTolueneValueBox.Size = new System.Drawing.Size(130, 64);
            this.CylinderRawTolueneValueBox.TabIndex = 28;
            // 
            // CylinderRawIPAValueBox
            // 
            this.CylinderRawIPAValueBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawIPAValueBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawIPAValueBox.Location = new System.Drawing.Point(276, 277);
            this.CylinderRawIPAValueBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawIPAValueBox.Multiline = true;
            this.CylinderRawIPAValueBox.Name = "CylinderRawIPAValueBox";
            this.CylinderRawIPAValueBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CylinderRawIPAValueBox.Size = new System.Drawing.Size(130, 64);
            this.CylinderRawIPAValueBox.TabIndex = 32;
            // 
            // CylinderRawAcetoneValueBox
            // 
            this.CylinderRawAcetoneValueBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawAcetoneValueBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.CylinderRawAcetoneValueBox.Location = new System.Drawing.Point(276, 342);
            this.CylinderRawAcetoneValueBox.Margin = new System.Windows.Forms.Padding(0);
            this.CylinderRawAcetoneValueBox.Multiline = true;
            this.CylinderRawAcetoneValueBox.Name = "CylinderRawAcetoneValueBox";
            this.CylinderRawAcetoneValueBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.CylinderRawAcetoneValueBox.Size = new System.Drawing.Size(130, 66);
            this.CylinderRawAcetoneValueBox.TabIndex = 36;
            // 
            // CylinderRawBackGround
            // 
            this.CylinderRawBackGround.AutoSize = true;
            this.CylinderRawBackGround.Location = new System.Drawing.Point(197, 1);
            this.CylinderRawBackGround.Name = "CylinderRawBackGround";
            this.CylinderRawBackGround.Size = new System.Drawing.Size(57, 20);
            this.CylinderRawBackGround.TabIndex = 21;
            this.CylinderRawBackGround.Text = "背景值";
            // 
            // CylinderRawConcertration
            // 
            this.CylinderRawConcertration.AutoSize = true;
            this.CylinderRawConcertration.Location = new System.Drawing.Point(115, 1);
            this.CylinderRawConcertration.Name = "CylinderRawConcertration";
            this.CylinderRawConcertration.Size = new System.Drawing.Size(41, 20);
            this.CylinderRawConcertration.TabIndex = 22;
            this.CylinderRawConcertration.Text = "濃度";
            // 
            // CylinderRawH2S
            // 
            this.CylinderRawH2S.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.CylinderRawH2S.AutoSize = true;
            this.CylinderRawH2S.Location = new System.Drawing.Point(33, 82);
            this.CylinderRawH2S.Name = "CylinderRawH2S";
            this.CylinderRawH2S.Size = new System.Drawing.Size(75, 64);
            this.CylinderRawH2S.TabIndex = 20;
            this.CylinderRawH2S.Tag = "H2S";
            this.CylinderRawH2S.Text = "H2S";
            this.CylinderRawH2S.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CylinderRawEff
            // 
            this.CylinderRawEff.AutoSize = true;
            this.CylinderRawEff.Location = new System.Drawing.Point(564, 31);
            this.CylinderRawEff.Name = "CylinderRawEff";
            this.CylinderRawEff.Size = new System.Drawing.Size(41, 20);
            this.CylinderRawEff.TabIndex = 21;
            this.CylinderRawEff.Text = "效率";
            // 
            // CylinderRawMeshBox
            // 
            this.CylinderRawMeshBox.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.CylinderRawMeshBox.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn12,
            this.dataGridViewTextBoxColumn13});
            this.CylinderRawMeshBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.CylinderRawMeshBox.Location = new System.Drawing.Point(314, 64);
            this.CylinderRawMeshBox.Name = "CylinderRawMeshBox";
            this.CylinderRawMeshBox.RowTemplate.Height = 24;
            this.CylinderRawMeshBox.Size = new System.Drawing.Size(237, 139);
            this.CylinderRawMeshBox.TabIndex = 51;
            // 
            // dataGridViewTextBoxColumn12
            // 
            this.dataGridViewTextBoxColumn12.HeaderText = "目數";
            this.dataGridViewTextBoxColumn12.Name = "dataGridViewTextBoxColumn12";
            // 
            // dataGridViewTextBoxColumn13
            // 
            this.dataGridViewTextBoxColumn13.HeaderText = "weight";
            this.dataGridViewTextBoxColumn13.Name = "dataGridViewTextBoxColumn13";
            // 
            // CylinderRawParticle
            // 
            this.CylinderRawParticle.AutoSize = true;
            this.CylinderRawParticle.Location = new System.Drawing.Point(310, 30);
            this.CylinderRawParticle.Name = "CylinderRawParticle";
            this.CylinderRawParticle.Size = new System.Drawing.Size(73, 20);
            this.CylinderRawParticle.TabIndex = 8;
            this.CylinderRawParticle.Text = "粒徑測試";
            // 
            // FilterPage
            // 
            this.FilterPage.BackColor = System.Drawing.Color.White;
            this.FilterPage.Controls.Add(this.FilterReportCustmorBox);
            this.FilterPage.Controls.Add(this.FilterModelBox);
            this.FilterPage.Controls.Add(this.FilterProductionBox);
            this.FilterPage.Controls.Add(this.FilterTestDateBox);
            this.FilterPage.Controls.Add(this.FilterBox);
            this.FilterPage.Controls.Add(this.FilterMaterialNumerBox);
            this.FilterPage.Controls.Add(this.FilterAlarmBox);
            this.FilterPage.Controls.Add(this.FilterCarbonLotBox);
            this.FilterPage.Controls.Add(this.ReFilterNOBox);
            this.FilterPage.Controls.Add(this.FilterPackageNOBox);
            this.FilterPage.Controls.Add(this.FilterMaterialNumer);
            this.FilterPage.Controls.Add(this.FilterAlarm);
            this.FilterPage.Controls.Add(this.label1);
            this.FilterPage.Controls.Add(this.FilterReportBox);
            this.FilterPage.Controls.Add(this.ReFilterNO);
            this.FilterPage.Controls.Add(this.FilterModel);
            this.FilterPage.Controls.Add(this.FilterReportCustmor);
            this.FilterPage.Controls.Add(this.FilterPackageNO);
            this.FilterPage.Controls.Add(this.FilterProduction);
            this.FilterPage.Controls.Add(this.FilterReport);
            this.FilterPage.Controls.Add(this.FilterTestDate);
            this.FilterPage.Location = new System.Drawing.Point(4, 29);
            this.FilterPage.Name = "FilterPage";
            this.FilterPage.Size = new System.Drawing.Size(1018, 625);
            this.FilterPage.TabIndex = 2;
            this.FilterPage.Text = "濾網成品";
            // 
            // FilterReportCustmorBox
            // 
            this.FilterReportCustmorBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterReportCustmorBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.FilterReportCustmorBox.FormattingEnabled = true;
            this.FilterReportCustmorBox.Items.AddRange(new object[] {
            "TSMC",
            "台中美光",
            "桃園美光"});
            this.FilterReportCustmorBox.Location = new System.Drawing.Point(175, 187);
            this.FilterReportCustmorBox.Name = "FilterReportCustmorBox";
            this.FilterReportCustmorBox.Size = new System.Drawing.Size(100, 28);
            this.FilterReportCustmorBox.TabIndex = 5;
            // 
            // FilterModelBox
            // 
            this.FilterModelBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterModelBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.FilterModelBox.FormattingEnabled = true;
            this.FilterModelBox.Items.AddRange(new object[] {
            "6VS",
            "6VB",
            "4VS",
            "4VB",
            "Panel",
            "蒸籠式",
            "抽取式單片",
            "抽取式整組",
            "堆疊式單片",
            "堆疊式整組"});
            this.FilterModelBox.Location = new System.Drawing.Point(175, 230);
            this.FilterModelBox.Name = "FilterModelBox";
            this.FilterModelBox.Size = new System.Drawing.Size(100, 28);
            this.FilterModelBox.TabIndex = 5;
            // 
            // FilterProductionBox
            // 
            this.FilterProductionBox.Checked = false;
            this.FilterProductionBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterProductionBox.CustomFormat = "yyyy.MM.dd";
            this.FilterProductionBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.FilterProductionBox.Location = new System.Drawing.Point(177, 29);
            this.FilterProductionBox.Name = "FilterProductionBox";
            this.FilterProductionBox.Size = new System.Drawing.Size(100, 29);
            this.FilterProductionBox.TabIndex = 1;
            this.FilterProductionBox.ValueChanged += new System.EventHandler(this.FilterProductTestDateBox_ValueChanged);
            // 
            // FilterTestDateBox
            // 
            this.FilterTestDateBox.Checked = false;
            this.FilterTestDateBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterTestDateBox.CustomFormat = "yyyy.MM.dd";
            this.FilterTestDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.FilterTestDateBox.Location = new System.Drawing.Point(177, 72);
            this.FilterTestDateBox.Name = "FilterTestDateBox";
            this.FilterTestDateBox.Size = new System.Drawing.Size(100, 29);
            this.FilterTestDateBox.TabIndex = 1;
            this.FilterTestDateBox.ValueChanged += new System.EventHandler(this.FilterProductTestDateBox_ValueChanged);
            // 
            // FilterBox
            // 
            this.FilterBox.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.FilterBox.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.生產序號,
            this.重量,
            this.length,
            this.width,
            this.height,
            this.diagonal,
            this.Particle_In,
            this.Particle_Out,
            this.IPA_In,
            this.IPA_Out,
            this.Acetone_In,
            this.Acetone_Out,
            this.Nontarget_In,
            this.Nontarget_Out,
            this.Pressure_Drop_spec,
            this.Pressure_Drop});
            this.FilterBox.Location = new System.Drawing.Point(40, 310);
            this.FilterBox.Name = "FilterBox";
            this.FilterBox.RowTemplate.Height = 24;
            this.FilterBox.Size = new System.Drawing.Size(915, 270);
            this.FilterBox.TabIndex = 11;
            // 
            // 生產序號
            // 
            this.生產序號.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.生產序號.Frozen = true;
            this.生產序號.HeaderText = "生產序號";
            this.生產序號.Name = "生產序號";
            // 
            // 重量
            // 
            this.重量.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.重量.HeaderText = "重量";
            this.重量.Name = "重量";
            this.重量.Width = 70;
            // 
            // length
            // 
            this.length.HeaderText = "長";
            this.length.Name = "length";
            // 
            // width
            // 
            this.width.HeaderText = "寬";
            this.width.Name = "width";
            // 
            // height
            // 
            this.height.HeaderText = "高";
            this.height.Name = "height";
            // 
            // diagonal
            // 
            this.diagonal.HeaderText = "對角線";
            this.diagonal.Name = "diagonal";
            // 
            // Particle_In
            // 
            this.Particle_In.HeaderText = "Particle_In";
            this.Particle_In.Name = "Particle_In";
            this.Particle_In.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // Particle_Out
            // 
            this.Particle_Out.HeaderText = "Particle_Out";
            this.Particle_Out.Name = "Particle_Out";
            this.Particle_Out.Width = 105;
            // 
            // IPA_In
            // 
            this.IPA_In.HeaderText = "IPA_In";
            this.IPA_In.Name = "IPA_In";
            this.IPA_In.Width = 65;
            // 
            // IPA_Out
            // 
            this.IPA_Out.HeaderText = "IPA_Out";
            this.IPA_Out.Name = "IPA_Out";
            this.IPA_Out.Width = 78;
            // 
            // Acetone_In
            // 
            this.Acetone_In.HeaderText = "Acetone_In";
            this.Acetone_In.Name = "Acetone_In";
            // 
            // Acetone_Out
            // 
            this.Acetone_Out.HeaderText = "Acetone_Out";
            this.Acetone_Out.Name = "Acetone_Out";
            this.Acetone_Out.Width = 120;
            // 
            // Nontarget_In
            // 
            this.Nontarget_In.HeaderText = "Nontarget_In";
            this.Nontarget_In.Name = "Nontarget_In";
            this.Nontarget_In.Width = 125;
            // 
            // Nontarget_Out
            // 
            this.Nontarget_Out.HeaderText = "Nontarget_Out";
            this.Nontarget_Out.Name = "Nontarget_Out";
            this.Nontarget_Out.Width = 130;
            // 
            // Pressure_Drop_spec
            // 
            this.Pressure_Drop_spec.HeaderText = "Pressure_Drop_Spec";
            this.Pressure_Drop_spec.Name = "Pressure_Drop_spec";
            this.Pressure_Drop_spec.Width = 175;
            // 
            // Pressure_Drop
            // 
            this.Pressure_Drop.HeaderText = "Pressure_Drop";
            this.Pressure_Drop.Name = "Pressure_Drop";
            this.Pressure_Drop.Width = 140;
            // 
            // FilterMaterialNumerBox
            // 
            this.FilterMaterialNumerBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterMaterialNumerBox.Location = new System.Drawing.Point(391, 110);
            this.FilterMaterialNumerBox.Name = "FilterMaterialNumerBox";
            this.FilterMaterialNumerBox.Size = new System.Drawing.Size(100, 29);
            this.FilterMaterialNumerBox.TabIndex = 10;
            // 
            // FilterAlarmBox
            // 
            this.FilterAlarmBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterAlarmBox.Location = new System.Drawing.Point(391, 70);
            this.FilterAlarmBox.Name = "FilterAlarmBox";
            this.FilterAlarmBox.Size = new System.Drawing.Size(100, 29);
            this.FilterAlarmBox.TabIndex = 10;
            // 
            // FilterCarbonLotBox
            // 
            this.FilterCarbonLotBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterCarbonLotBox.Location = new System.Drawing.Point(391, 25);
            this.FilterCarbonLotBox.Name = "FilterCarbonLotBox";
            this.FilterCarbonLotBox.Size = new System.Drawing.Size(100, 29);
            this.FilterCarbonLotBox.TabIndex = 8;
            // 
            // ReFilterNOBox
            // 
            this.ReFilterNOBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.ReFilterNOBox.Location = new System.Drawing.Point(175, 269);
            this.ReFilterNOBox.Name = "ReFilterNOBox";
            this.ReFilterNOBox.Size = new System.Drawing.Size(100, 29);
            this.ReFilterNOBox.TabIndex = 7;
            // 
            // FilterPackageNOBox
            // 
            this.FilterPackageNOBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterPackageNOBox.Location = new System.Drawing.Point(175, 150);
            this.FilterPackageNOBox.Name = "FilterPackageNOBox";
            this.FilterPackageNOBox.Size = new System.Drawing.Size(100, 29);
            this.FilterPackageNOBox.TabIndex = 3;
            // 
            // FilterMaterialNumer
            // 
            this.FilterMaterialNumer.AutoSize = true;
            this.FilterMaterialNumer.Location = new System.Drawing.Point(311, 113);
            this.FilterMaterialNumer.Name = "FilterMaterialNumer";
            this.FilterMaterialNumer.Size = new System.Drawing.Size(41, 20);
            this.FilterMaterialNumer.TabIndex = 4;
            this.FilterMaterialNumer.Text = "料號";
            // 
            // FilterAlarm
            // 
            this.FilterAlarm.AutoSize = true;
            this.FilterAlarm.Location = new System.Drawing.Point(311, 73);
            this.FilterAlarm.Name = "FilterAlarm";
            this.FilterAlarm.Size = new System.Drawing.Size(73, 20);
            this.FilterAlarm.TabIndex = 4;
            this.FilterAlarm.Text = "異常原因";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(311, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 20);
            this.label1.TabIndex = 4;
            this.label1.Text = "碳線單號";
            // 
            // FilterReportBox
            // 
            this.FilterReportBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterReportBox.Location = new System.Drawing.Point(175, 110);
            this.FilterReportBox.Name = "FilterReportBox";
            this.FilterReportBox.Size = new System.Drawing.Size(100, 29);
            this.FilterReportBox.TabIndex = 2;
            // 
            // ReFilterNO
            // 
            this.ReFilterNO.AutoSize = true;
            this.ReFilterNO.Location = new System.Drawing.Point(40, 270);
            this.ReFilterNO.Name = "ReFilterNO";
            this.ReFilterNO.Size = new System.Drawing.Size(73, 20);
            this.ReFilterNO.TabIndex = 4;
            this.ReFilterNO.Text = "再生次數";
            // 
            // FilterModel
            // 
            this.FilterModel.AutoSize = true;
            this.FilterModel.Location = new System.Drawing.Point(40, 230);
            this.FilterModel.Name = "FilterModel";
            this.FilterModel.Size = new System.Drawing.Size(73, 20);
            this.FilterModel.TabIndex = 4;
            this.FilterModel.Text = "產品型號";
            // 
            // FilterReportCustmor
            // 
            this.FilterReportCustmor.AutoSize = true;
            this.FilterReportCustmor.Location = new System.Drawing.Point(40, 190);
            this.FilterReportCustmor.Name = "FilterReportCustmor";
            this.FilterReportCustmor.Size = new System.Drawing.Size(41, 20);
            this.FilterReportCustmor.TabIndex = 4;
            this.FilterReportCustmor.Text = "客戶";
            // 
            // FilterPackageNO
            // 
            this.FilterPackageNO.AutoSize = true;
            this.FilterPackageNO.Location = new System.Drawing.Point(40, 150);
            this.FilterPackageNO.Name = "FilterPackageNO";
            this.FilterPackageNO.Size = new System.Drawing.Size(73, 20);
            this.FilterPackageNO.TabIndex = 4;
            this.FilterPackageNO.Text = "包裝單號";
            // 
            // FilterProduction
            // 
            this.FilterProduction.AutoSize = true;
            this.FilterProduction.Location = new System.Drawing.Point(40, 30);
            this.FilterProduction.Name = "FilterProduction";
            this.FilterProduction.Size = new System.Drawing.Size(73, 20);
            this.FilterProduction.TabIndex = 4;
            this.FilterProduction.Text = "生產日期";
            // 
            // FilterReport
            // 
            this.FilterReport.AutoSize = true;
            this.FilterReport.Location = new System.Drawing.Point(40, 110);
            this.FilterReport.Name = "FilterReport";
            this.FilterReport.Size = new System.Drawing.Size(73, 20);
            this.FilterReport.TabIndex = 4;
            this.FilterReport.Text = "報告編號";
            // 
            // FilterTestDate
            // 
            this.FilterTestDate.AutoSize = true;
            this.FilterTestDate.Location = new System.Drawing.Point(40, 70);
            this.FilterTestDate.Name = "FilterTestDate";
            this.FilterTestDate.Size = new System.Drawing.Size(73, 20);
            this.FilterTestDate.TabIndex = 4;
            this.FilterTestDate.Text = "測試日期";
            // 
            // FilterInProcessPage
            // 
            this.FilterInProcessPage.BackColor = System.Drawing.Color.White;
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessEffPanel);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessTypeBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessTestBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessProductionBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessCarbonInfoBox);
            this.FilterInProcessPage.Controls.Add(this.FilterSizeTB);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessPressureDropBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessTestGsmBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessWindBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessPressureBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessLowerBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessUpperBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessSpeedBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessGileBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessgsmBox);
            this.FilterInProcessPage.Controls.Add(this.FilterMaterialNOBox);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessCarbonOrderBox);
            this.FilterInProcessPage.Controls.Add(this.label2);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessReportNOTB);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessOrderBox);
            this.FilterInProcessPage.Controls.Add(this.FilterSize);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessWire);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessEff);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessTestGsm);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessWind);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessPressure);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessLower);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessUpper);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessSpeed);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessGile);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessType);
            this.FilterInProcessPage.Controls.Add(this.FilterMaterialNO);
            this.FilterInProcessPage.Controls.Add(this.label3);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessReportNO);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessgsm);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessOrder);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessTest);
            this.FilterInProcessPage.Controls.Add(this.FilterInProcessProduction);
            this.FilterInProcessPage.Location = new System.Drawing.Point(4, 29);
            this.FilterInProcessPage.Name = "FilterInProcessPage";
            this.FilterInProcessPage.Padding = new System.Windows.Forms.Padding(3);
            this.FilterInProcessPage.Size = new System.Drawing.Size(1018, 625);
            this.FilterInProcessPage.TabIndex = 1;
            this.FilterInProcessPage.Text = "濾網半成品";
            // 
            // FilterInProcessEffPanel
            // 
            this.FilterInProcessEffPanel.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.FilterInProcessEffPanel.ColumnCount = 5;
            this.FilterInProcessEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 5.951709F));
            this.FilterInProcessEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.57306F));
            this.FilterInProcessEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.57306F));
            this.FilterInProcessEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20.57306F));
            this.FilterInProcessEffPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 32.32911F));
            this.FilterInProcessEffPanel.Controls.Add(this.label35, 0, 0);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInGas, 1, 0);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInSO2Concertration, 2, 0);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGround, 3, 0);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValue, 4, 0);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInSO2, 1, 1);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInH2S, 1, 2);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInNH3, 1, 3);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInToluene, 1, 4);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInSO2CheckBox, 0, 1);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInH2SCheckBox, 0, 2);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInNH3CheckBox, 0, 3);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInTolueneCheckBox, 0, 4);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInSO2ConcentrationBox, 2, 1);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInH2SConcentrationBox, 2, 2);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInNH3ConcentrationBox, 2, 3);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGroundTolueneBox, 3, 4);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGroundNH3Box, 3, 3);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGroundH2SBox, 3, 2);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGroundSO2Box, 3, 1);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInTolueneConcentrationBox, 2, 4);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValueSO2Box, 4, 1);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValueH2SBox, 4, 2);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValueNH3Box, 4, 3);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValueTolueneBox, 4, 4);
            this.FilterInProcessEffPanel.Controls.Add(this.checkBox2, 0, 6);
            this.FilterInProcessEffPanel.Controls.Add(this.label4, 1, 5);
            this.FilterInProcessEffPanel.Controls.Add(this.label5, 1, 6);
            this.FilterInProcessEffPanel.Controls.Add(this.checkBox1, 0, 5);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValueIPABox, 4, 5);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInValueAcetoneBox, 4, 6);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInIPAConcentrationBox, 2, 5);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInAcetoneConcentrationBox, 2, 6);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGroundIPABox, 3, 5);
            this.FilterInProcessEffPanel.Controls.Add(this.FilterInBackGroundAcetoneBox, 3, 6);
            this.FilterInProcessEffPanel.Location = new System.Drawing.Point(538, 53);
            this.FilterInProcessEffPanel.Name = "FilterInProcessEffPanel";
            this.FilterInProcessEffPanel.RowCount = 7;
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5.378973F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.77017F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.77017F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.77017F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.77017F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.77017F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.77017F));
            this.FilterInProcessEffPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.FilterInProcessEffPanel.Size = new System.Drawing.Size(407, 518);
            this.FilterInProcessEffPanel.TabIndex = 19;
            // 
            // label35
            // 
            this.label35.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.label35.AutoSize = true;
            this.label35.Location = new System.Drawing.Point(4, 1);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(17, 27);
            this.label35.TabIndex = 20;
            this.label35.Text = "V";
            this.label35.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // FilterInGas
            // 
            this.FilterInGas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInGas.Location = new System.Drawing.Point(45, 1);
            this.FilterInGas.Name = "FilterInGas";
            this.FilterInGas.Size = new System.Drawing.Size(41, 27);
            this.FilterInGas.TabIndex = 0;
            this.FilterInGas.Text = "氣體";
            this.FilterInGas.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // FilterInSO2Concertration
            // 
            this.FilterInSO2Concertration.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInSO2Concertration.AutoSize = true;
            this.FilterInSO2Concertration.Location = new System.Drawing.Point(128, 1);
            this.FilterInSO2Concertration.Name = "FilterInSO2Concertration";
            this.FilterInSO2Concertration.Size = new System.Drawing.Size(41, 27);
            this.FilterInSO2Concertration.TabIndex = 22;
            this.FilterInSO2Concertration.Text = "濃度";
            this.FilterInSO2Concertration.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // FilterInBackGround
            // 
            this.FilterInBackGround.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInBackGround.AutoSize = true;
            this.FilterInBackGround.Location = new System.Drawing.Point(203, 1);
            this.FilterInBackGround.Name = "FilterInBackGround";
            this.FilterInBackGround.Size = new System.Drawing.Size(57, 27);
            this.FilterInBackGround.TabIndex = 21;
            this.FilterInBackGround.Text = "背景值";
            this.FilterInBackGround.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // FilterInValue
            // 
            this.FilterInValue.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FilterInValue.AutoSize = true;
            this.FilterInValue.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.FilterInValue.Location = new System.Drawing.Point(277, 1);
            this.FilterInValue.Name = "FilterInValue";
            this.FilterInValue.Size = new System.Drawing.Size(126, 27);
            this.FilterInValue.TabIndex = 23;
            this.FilterInValue.Text = "讀值";
            // 
            // FilterInSO2
            // 
            this.FilterInSO2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.FilterInSO2.AutoSize = true;
            this.FilterInSO2.Location = new System.Drawing.Point(28, 59);
            this.FilterInSO2.Name = "FilterInSO2";
            this.FilterInSO2.Size = new System.Drawing.Size(76, 20);
            this.FilterInSO2.TabIndex = 20;
            this.FilterInSO2.Text = "SO2";
            this.FilterInSO2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // FilterInH2S
            // 
            this.FilterInH2S.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.FilterInH2S.AutoSize = true;
            this.FilterInH2S.Location = new System.Drawing.Point(28, 110);
            this.FilterInH2S.Name = "FilterInH2S";
            this.FilterInH2S.Size = new System.Drawing.Size(76, 80);
            this.FilterInH2S.TabIndex = 20;
            this.FilterInH2S.Text = "H2S";
            this.FilterInH2S.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // FilterInNH3
            // 
            this.FilterInNH3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.FilterInNH3.AutoSize = true;
            this.FilterInNH3.Location = new System.Drawing.Point(44, 221);
            this.FilterInNH3.Name = "FilterInNH3";
            this.FilterInNH3.Size = new System.Drawing.Size(43, 20);
            this.FilterInNH3.TabIndex = 20;
            this.FilterInNH3.Text = "NH3";
            // 
            // FilterInToluene
            // 
            this.FilterInToluene.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.FilterInToluene.AutoSize = true;
            this.FilterInToluene.Location = new System.Drawing.Point(31, 302);
            this.FilterInToluene.Name = "FilterInToluene";
            this.FilterInToluene.Size = new System.Drawing.Size(70, 20);
            this.FilterInToluene.TabIndex = 20;
            this.FilterInToluene.Text = "Toluene";
            // 
            // FilterInSO2CheckBox
            // 
            this.FilterInSO2CheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInSO2CheckBox.AutoSize = true;
            this.FilterInSO2CheckBox.Location = new System.Drawing.Point(5, 32);
            this.FilterInSO2CheckBox.Name = "FilterInSO2CheckBox";
            this.FilterInSO2CheckBox.Size = new System.Drawing.Size(15, 74);
            this.FilterInSO2CheckBox.TabIndex = 17;
            this.FilterInSO2CheckBox.Tag = "SO2";
            this.FilterInSO2CheckBox.UseVisualStyleBackColor = true;
            // 
            // FilterInH2SCheckBox
            // 
            this.FilterInH2SCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInH2SCheckBox.AutoSize = true;
            this.FilterInH2SCheckBox.Location = new System.Drawing.Point(5, 113);
            this.FilterInH2SCheckBox.Name = "FilterInH2SCheckBox";
            this.FilterInH2SCheckBox.Size = new System.Drawing.Size(15, 74);
            this.FilterInH2SCheckBox.TabIndex = 21;
            this.FilterInH2SCheckBox.Tag = "H2S";
            this.FilterInH2SCheckBox.UseVisualStyleBackColor = true;
            // 
            // FilterInNH3CheckBox
            // 
            this.FilterInNH3CheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInNH3CheckBox.AutoSize = true;
            this.FilterInNH3CheckBox.Location = new System.Drawing.Point(5, 194);
            this.FilterInNH3CheckBox.Name = "FilterInNH3CheckBox";
            this.FilterInNH3CheckBox.Size = new System.Drawing.Size(15, 74);
            this.FilterInNH3CheckBox.TabIndex = 25;
            this.FilterInNH3CheckBox.Tag = "NH3";
            this.FilterInNH3CheckBox.UseVisualStyleBackColor = true;
            // 
            // FilterInTolueneCheckBox
            // 
            this.FilterInTolueneCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.FilterInTolueneCheckBox.AutoSize = true;
            this.FilterInTolueneCheckBox.Location = new System.Drawing.Point(5, 275);
            this.FilterInTolueneCheckBox.Name = "FilterInTolueneCheckBox";
            this.FilterInTolueneCheckBox.Size = new System.Drawing.Size(15, 74);
            this.FilterInTolueneCheckBox.TabIndex = 29;
            this.FilterInTolueneCheckBox.Tag = "Toluene";
            this.FilterInTolueneCheckBox.UseVisualStyleBackColor = true;
            // 
            // FilterInSO2ConcentrationBox
            // 
            this.FilterInSO2ConcentrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInSO2ConcentrationBox.Location = new System.Drawing.Point(108, 29);
            this.FilterInSO2ConcentrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInSO2ConcentrationBox.Name = "FilterInSO2ConcentrationBox";
            this.FilterInSO2ConcentrationBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInSO2ConcentrationBox.TabIndex = 18;
            // 
            // FilterInH2SConcentrationBox
            // 
            this.FilterInH2SConcentrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInH2SConcentrationBox.Location = new System.Drawing.Point(108, 110);
            this.FilterInH2SConcentrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInH2SConcentrationBox.Name = "FilterInH2SConcentrationBox";
            this.FilterInH2SConcentrationBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInH2SConcentrationBox.TabIndex = 22;
            // 
            // FilterInNH3ConcentrationBox
            // 
            this.FilterInNH3ConcentrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInNH3ConcentrationBox.Location = new System.Drawing.Point(108, 191);
            this.FilterInNH3ConcentrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInNH3ConcentrationBox.Name = "FilterInNH3ConcentrationBox";
            this.FilterInNH3ConcentrationBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInNH3ConcentrationBox.TabIndex = 26;
            // 
            // FilterInBackGroundTolueneBox
            // 
            this.FilterInBackGroundTolueneBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInBackGroundTolueneBox.Location = new System.Drawing.Point(191, 272);
            this.FilterInBackGroundTolueneBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInBackGroundTolueneBox.Name = "FilterInBackGroundTolueneBox";
            this.FilterInBackGroundTolueneBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInBackGroundTolueneBox.TabIndex = 31;
            // 
            // FilterInBackGroundNH3Box
            // 
            this.FilterInBackGroundNH3Box.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInBackGroundNH3Box.Location = new System.Drawing.Point(191, 191);
            this.FilterInBackGroundNH3Box.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInBackGroundNH3Box.Name = "FilterInBackGroundNH3Box";
            this.FilterInBackGroundNH3Box.Size = new System.Drawing.Size(82, 22);
            this.FilterInBackGroundNH3Box.TabIndex = 27;
            // 
            // FilterInBackGroundH2SBox
            // 
            this.FilterInBackGroundH2SBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInBackGroundH2SBox.Location = new System.Drawing.Point(191, 110);
            this.FilterInBackGroundH2SBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInBackGroundH2SBox.Name = "FilterInBackGroundH2SBox";
            this.FilterInBackGroundH2SBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInBackGroundH2SBox.TabIndex = 23;
            // 
            // FilterInBackGroundSO2Box
            // 
            this.FilterInBackGroundSO2Box.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInBackGroundSO2Box.Location = new System.Drawing.Point(191, 29);
            this.FilterInBackGroundSO2Box.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInBackGroundSO2Box.Name = "FilterInBackGroundSO2Box";
            this.FilterInBackGroundSO2Box.Size = new System.Drawing.Size(82, 22);
            this.FilterInBackGroundSO2Box.TabIndex = 19;
            // 
            // FilterInTolueneConcentrationBox
            // 
            this.FilterInTolueneConcentrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInTolueneConcentrationBox.Location = new System.Drawing.Point(108, 272);
            this.FilterInTolueneConcentrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInTolueneConcentrationBox.Name = "FilterInTolueneConcentrationBox";
            this.FilterInTolueneConcentrationBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInTolueneConcentrationBox.TabIndex = 30;
            // 
            // FilterInValueSO2Box
            // 
            this.FilterInValueSO2Box.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInValueSO2Box.Location = new System.Drawing.Point(274, 29);
            this.FilterInValueSO2Box.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInValueSO2Box.Multiline = true;
            this.FilterInValueSO2Box.Name = "FilterInValueSO2Box";
            this.FilterInValueSO2Box.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.FilterInValueSO2Box.Size = new System.Drawing.Size(132, 80);
            this.FilterInValueSO2Box.TabIndex = 20;
            // 
            // FilterInValueH2SBox
            // 
            this.FilterInValueH2SBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInValueH2SBox.Location = new System.Drawing.Point(274, 110);
            this.FilterInValueH2SBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInValueH2SBox.Multiline = true;
            this.FilterInValueH2SBox.Name = "FilterInValueH2SBox";
            this.FilterInValueH2SBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.FilterInValueH2SBox.Size = new System.Drawing.Size(132, 80);
            this.FilterInValueH2SBox.TabIndex = 24;
            // 
            // FilterInValueNH3Box
            // 
            this.FilterInValueNH3Box.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInValueNH3Box.Location = new System.Drawing.Point(274, 191);
            this.FilterInValueNH3Box.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInValueNH3Box.Multiline = true;
            this.FilterInValueNH3Box.Name = "FilterInValueNH3Box";
            this.FilterInValueNH3Box.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.FilterInValueNH3Box.Size = new System.Drawing.Size(132, 80);
            this.FilterInValueNH3Box.TabIndex = 28;
            // 
            // FilterInValueTolueneBox
            // 
            this.FilterInValueTolueneBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInValueTolueneBox.Location = new System.Drawing.Point(274, 272);
            this.FilterInValueTolueneBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInValueTolueneBox.Multiline = true;
            this.FilterInValueTolueneBox.Name = "FilterInValueTolueneBox";
            this.FilterInValueTolueneBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.FilterInValueTolueneBox.Size = new System.Drawing.Size(132, 80);
            this.FilterInValueTolueneBox.TabIndex = 32;
            // 
            // checkBox2
            // 
            this.checkBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(4, 437);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(17, 77);
            this.checkBox2.TabIndex = 37;
            this.checkBox2.Tag = "Acetone";
            this.checkBox2.Text = "checkBox1";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(48, 383);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(35, 20);
            this.label4.TabIndex = 20;
            this.label4.Text = "IPA";
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(30, 465);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 20);
            this.label5.TabIndex = 20;
            this.label5.Text = "Acetone";
            // 
            // checkBox1
            // 
            this.checkBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(4, 356);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(17, 74);
            this.checkBox1.TabIndex = 33;
            this.checkBox1.Tag = "IPA";
            this.checkBox1.Text = "checkBox1";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // FilterInValueIPABox
            // 
            this.FilterInValueIPABox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInValueIPABox.Location = new System.Drawing.Point(274, 353);
            this.FilterInValueIPABox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInValueIPABox.Multiline = true;
            this.FilterInValueIPABox.Name = "FilterInValueIPABox";
            this.FilterInValueIPABox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.FilterInValueIPABox.Size = new System.Drawing.Size(132, 80);
            this.FilterInValueIPABox.TabIndex = 36;
            // 
            // FilterInValueAcetoneBox
            // 
            this.FilterInValueAcetoneBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInValueAcetoneBox.Location = new System.Drawing.Point(274, 434);
            this.FilterInValueAcetoneBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInValueAcetoneBox.Multiline = true;
            this.FilterInValueAcetoneBox.Name = "FilterInValueAcetoneBox";
            this.FilterInValueAcetoneBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.FilterInValueAcetoneBox.Size = new System.Drawing.Size(132, 80);
            this.FilterInValueAcetoneBox.TabIndex = 40;
            // 
            // FilterInIPAConcentrationBox
            // 
            this.FilterInIPAConcentrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInIPAConcentrationBox.Location = new System.Drawing.Point(108, 353);
            this.FilterInIPAConcentrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInIPAConcentrationBox.Name = "FilterInIPAConcentrationBox";
            this.FilterInIPAConcentrationBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInIPAConcentrationBox.TabIndex = 34;
            // 
            // FilterInAcetoneConcentrationBox
            // 
            this.FilterInAcetoneConcentrationBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInAcetoneConcentrationBox.Location = new System.Drawing.Point(108, 434);
            this.FilterInAcetoneConcentrationBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInAcetoneConcentrationBox.Name = "FilterInAcetoneConcentrationBox";
            this.FilterInAcetoneConcentrationBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInAcetoneConcentrationBox.TabIndex = 38;
            // 
            // FilterInBackGroundIPABox
            // 
            this.FilterInBackGroundIPABox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInBackGroundIPABox.Location = new System.Drawing.Point(191, 353);
            this.FilterInBackGroundIPABox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInBackGroundIPABox.Name = "FilterInBackGroundIPABox";
            this.FilterInBackGroundIPABox.Size = new System.Drawing.Size(82, 22);
            this.FilterInBackGroundIPABox.TabIndex = 35;
            // 
            // FilterInBackGroundAcetoneBox
            // 
            this.FilterInBackGroundAcetoneBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FilterInBackGroundAcetoneBox.Location = new System.Drawing.Point(191, 434);
            this.FilterInBackGroundAcetoneBox.Margin = new System.Windows.Forms.Padding(0);
            this.FilterInBackGroundAcetoneBox.Name = "FilterInBackGroundAcetoneBox";
            this.FilterInBackGroundAcetoneBox.Size = new System.Drawing.Size(82, 22);
            this.FilterInBackGroundAcetoneBox.TabIndex = 39;
            // 
            // FilterInProcessTypeBox
            // 
            this.FilterInProcessTypeBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessTypeBox.FormattingEnabled = true;
            this.FilterInProcessTypeBox.Items.AddRange(new object[] {
            "SG001",
            "SG009",
            "SG032",
            "SG042",
            "IER101",
            "SI013",
            "IKP101",
            "SG001+IKP101",
            "SG009+IKP101",
            "SG032+IKP101",
            "SG042+IKP101",
            "IER101+IKP101",
            "SI013+IKP101"});
            this.FilterInProcessTypeBox.Location = new System.Drawing.Point(175, 265);
            this.FilterInProcessTypeBox.Name = "FilterInProcessTypeBox";
            this.FilterInProcessTypeBox.Size = new System.Drawing.Size(100, 28);
            this.FilterInProcessTypeBox.TabIndex = 6;
            // 
            // FilterInProcessTestBox
            // 
            this.FilterInProcessTestBox.Checked = false;
            this.FilterInProcessTestBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessTestBox.CustomFormat = "yyyy.MM.dd";
            this.FilterInProcessTestBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.FilterInProcessTestBox.Location = new System.Drawing.Point(176, 66);
            this.FilterInProcessTestBox.Name = "FilterInProcessTestBox";
            this.FilterInProcessTestBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessTestBox.TabIndex = 3;
            this.FilterInProcessTestBox.ValueChanged += new System.EventHandler(this.FilterInProcessTestBox_ValueChange);
            // 
            // FilterInProcessProductionBox
            // 
            this.FilterInProcessProductionBox.Checked = false;
            this.FilterInProcessProductionBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessProductionBox.CustomFormat = "yyyy.MM.dd";
            this.FilterInProcessProductionBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.FilterInProcessProductionBox.Location = new System.Drawing.Point(176, 26);
            this.FilterInProcessProductionBox.Name = "FilterInProcessProductionBox";
            this.FilterInProcessProductionBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessProductionBox.TabIndex = 2;
            // 
            // FilterInProcessCarbonInfoBox
            // 
            this.FilterInProcessCarbonInfoBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessCarbonInfoBox.Location = new System.Drawing.Point(390, 179);
            this.FilterInProcessCarbonInfoBox.Name = "FilterInProcessCarbonInfoBox";
            this.FilterInProcessCarbonInfoBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessCarbonInfoBox.TabIndex = 16;
            // 
            // FilterSizeTB
            // 
            this.FilterSizeTB.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterSizeTB.Location = new System.Drawing.Point(390, 141);
            this.FilterSizeTB.Name = "FilterSizeTB";
            this.FilterSizeTB.Size = new System.Drawing.Size(100, 29);
            this.FilterSizeTB.TabIndex = 15;
            // 
            // FilterInProcessPressureDropBox
            // 
            this.FilterInProcessPressureDropBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessPressureDropBox.Location = new System.Drawing.Point(390, 103);
            this.FilterInProcessPressureDropBox.Name = "FilterInProcessPressureDropBox";
            this.FilterInProcessPressureDropBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessPressureDropBox.TabIndex = 15;
            // 
            // FilterInProcessTestGsmBox
            // 
            this.FilterInProcessTestGsmBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessTestGsmBox.Location = new System.Drawing.Point(390, 61);
            this.FilterInProcessTestGsmBox.Name = "FilterInProcessTestGsmBox";
            this.FilterInProcessTestGsmBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessTestGsmBox.TabIndex = 14;
            // 
            // FilterInProcessWindBox
            // 
            this.FilterInProcessWindBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessWindBox.Location = new System.Drawing.Point(390, 27);
            this.FilterInProcessWindBox.Name = "FilterInProcessWindBox";
            this.FilterInProcessWindBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessWindBox.TabIndex = 13;
            // 
            // FilterInProcessPressureBox
            // 
            this.FilterInProcessPressureBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessPressureBox.Location = new System.Drawing.Point(175, 505);
            this.FilterInProcessPressureBox.Name = "FilterInProcessPressureBox";
            this.FilterInProcessPressureBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessPressureBox.TabIndex = 12;
            // 
            // FilterInProcessLowerBox
            // 
            this.FilterInProcessLowerBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessLowerBox.Location = new System.Drawing.Point(175, 465);
            this.FilterInProcessLowerBox.Name = "FilterInProcessLowerBox";
            this.FilterInProcessLowerBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessLowerBox.TabIndex = 11;
            // 
            // FilterInProcessUpperBox
            // 
            this.FilterInProcessUpperBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessUpperBox.Location = new System.Drawing.Point(175, 425);
            this.FilterInProcessUpperBox.Name = "FilterInProcessUpperBox";
            this.FilterInProcessUpperBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessUpperBox.TabIndex = 10;
            // 
            // FilterInProcessSpeedBox
            // 
            this.FilterInProcessSpeedBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessSpeedBox.Location = new System.Drawing.Point(175, 385);
            this.FilterInProcessSpeedBox.Name = "FilterInProcessSpeedBox";
            this.FilterInProcessSpeedBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessSpeedBox.TabIndex = 9;
            // 
            // FilterInProcessGileBox
            // 
            this.FilterInProcessGileBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessGileBox.Location = new System.Drawing.Point(175, 345);
            this.FilterInProcessGileBox.Name = "FilterInProcessGileBox";
            this.FilterInProcessGileBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessGileBox.TabIndex = 8;
            // 
            // FilterInProcessgsmBox
            // 
            this.FilterInProcessgsmBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessgsmBox.Location = new System.Drawing.Point(175, 305);
            this.FilterInProcessgsmBox.Name = "FilterInProcessgsmBox";
            this.FilterInProcessgsmBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessgsmBox.TabIndex = 7;
            // 
            // FilterMaterialNOBox
            // 
            this.FilterMaterialNOBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterMaterialNOBox.Location = new System.Drawing.Point(175, 223);
            this.FilterMaterialNOBox.Name = "FilterMaterialNOBox";
            this.FilterMaterialNOBox.Size = new System.Drawing.Size(100, 29);
            this.FilterMaterialNOBox.TabIndex = 5;
            // 
            // FilterInProcessCarbonOrderBox
            // 
            this.FilterInProcessCarbonOrderBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessCarbonOrderBox.Location = new System.Drawing.Point(175, 183);
            this.FilterInProcessCarbonOrderBox.Name = "FilterInProcessCarbonOrderBox";
            this.FilterInProcessCarbonOrderBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessCarbonOrderBox.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(310, 184);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 20);
            this.label2.TabIndex = 12;
            this.label2.Text = "碳線資訊";
            // 
            // FilterInProcessReportNOTB
            // 
            this.FilterInProcessReportNOTB.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessReportNOTB.Location = new System.Drawing.Point(175, 103);
            this.FilterInProcessReportNOTB.Name = "FilterInProcessReportNOTB";
            this.FilterInProcessReportNOTB.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessReportNOTB.TabIndex = 3;
            // 
            // FilterInProcessOrderBox
            // 
            this.FilterInProcessOrderBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterInProcessOrderBox.Location = new System.Drawing.Point(175, 143);
            this.FilterInProcessOrderBox.Name = "FilterInProcessOrderBox";
            this.FilterInProcessOrderBox.Size = new System.Drawing.Size(100, 29);
            this.FilterInProcessOrderBox.TabIndex = 3;
            // 
            // FilterSize
            // 
            this.FilterSize.AutoSize = true;
            this.FilterSize.Location = new System.Drawing.Point(310, 146);
            this.FilterSize.Name = "FilterSize";
            this.FilterSize.Size = new System.Drawing.Size(73, 20);
            this.FilterSize.TabIndex = 12;
            this.FilterSize.Text = "成品尺寸";
            // 
            // FilterInProcessWire
            // 
            this.FilterInProcessWire.AutoSize = true;
            this.FilterInProcessWire.Location = new System.Drawing.Point(310, 108);
            this.FilterInProcessWire.Name = "FilterInProcessWire";
            this.FilterInProcessWire.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessWire.TabIndex = 12;
            this.FilterInProcessWire.Text = "壓損";
            // 
            // FilterInProcessEff
            // 
            this.FilterInProcessEff.AutoSize = true;
            this.FilterInProcessEff.Location = new System.Drawing.Point(534, 30);
            this.FilterInProcessEff.Name = "FilterInProcessEff";
            this.FilterInProcessEff.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessEff.TabIndex = 2;
            this.FilterInProcessEff.Text = "效率";
            // 
            // FilterInProcessTestGsm
            // 
            this.FilterInProcessTestGsm.AutoSize = true;
            this.FilterInProcessTestGsm.Location = new System.Drawing.Point(310, 66);
            this.FilterInProcessTestGsm.Name = "FilterInProcessTestGsm";
            this.FilterInProcessTestGsm.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessTestGsm.TabIndex = 2;
            this.FilterInProcessTestGsm.Text = "重量";
            // 
            // FilterInProcessWind
            // 
            this.FilterInProcessWind.AutoSize = true;
            this.FilterInProcessWind.Location = new System.Drawing.Point(310, 30);
            this.FilterInProcessWind.Name = "FilterInProcessWind";
            this.FilterInProcessWind.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessWind.TabIndex = 2;
            this.FilterInProcessWind.Text = "風速";
            // 
            // FilterInProcessPressure
            // 
            this.FilterInProcessPressure.AutoSize = true;
            this.FilterInProcessPressure.Location = new System.Drawing.Point(40, 508);
            this.FilterInProcessPressure.Name = "FilterInProcessPressure";
            this.FilterInProcessPressure.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessPressure.TabIndex = 2;
            this.FilterInProcessPressure.Text = "壓力";
            // 
            // FilterInProcessLower
            // 
            this.FilterInProcessLower.AutoSize = true;
            this.FilterInProcessLower.Location = new System.Drawing.Point(40, 468);
            this.FilterInProcessLower.Name = "FilterInProcessLower";
            this.FilterInProcessLower.Size = new System.Drawing.Size(89, 20);
            this.FilterInProcessLower.TabIndex = 2;
            this.FilterInProcessLower.Text = "下烘爐溫度";
            // 
            // FilterInProcessUpper
            // 
            this.FilterInProcessUpper.AutoSize = true;
            this.FilterInProcessUpper.Location = new System.Drawing.Point(40, 428);
            this.FilterInProcessUpper.Name = "FilterInProcessUpper";
            this.FilterInProcessUpper.Size = new System.Drawing.Size(89, 20);
            this.FilterInProcessUpper.TabIndex = 2;
            this.FilterInProcessUpper.Text = "上烘爐溫度";
            // 
            // FilterInProcessSpeed
            // 
            this.FilterInProcessSpeed.AutoSize = true;
            this.FilterInProcessSpeed.Location = new System.Drawing.Point(40, 388);
            this.FilterInProcessSpeed.Name = "FilterInProcessSpeed";
            this.FilterInProcessSpeed.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessSpeed.TabIndex = 2;
            this.FilterInProcessSpeed.Text = "速度";
            // 
            // FilterInProcessGile
            // 
            this.FilterInProcessGile.AutoSize = true;
            this.FilterInProcessGile.Location = new System.Drawing.Point(40, 348);
            this.FilterInProcessGile.Name = "FilterInProcessGile";
            this.FilterInProcessGile.Size = new System.Drawing.Size(41, 20);
            this.FilterInProcessGile.TabIndex = 2;
            this.FilterInProcessGile.Text = "膠粉";
            // 
            // FilterInProcessType
            // 
            this.FilterInProcessType.AutoSize = true;
            this.FilterInProcessType.Location = new System.Drawing.Point(40, 268);
            this.FilterInProcessType.Name = "FilterInProcessType";
            this.FilterInProcessType.Size = new System.Drawing.Size(73, 20);
            this.FilterInProcessType.TabIndex = 2;
            this.FilterInProcessType.Text = "原料種類";
            // 
            // FilterMaterialNO
            // 
            this.FilterMaterialNO.AutoSize = true;
            this.FilterMaterialNO.Location = new System.Drawing.Point(40, 228);
            this.FilterMaterialNO.Name = "FilterMaterialNO";
            this.FilterMaterialNO.Size = new System.Drawing.Size(73, 20);
            this.FilterMaterialNO.TabIndex = 2;
            this.FilterMaterialNO.Text = "原料批號";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(40, 188);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "碳線單號";
            // 
            // FilterInProcessReportNO
            // 
            this.FilterInProcessReportNO.AutoSize = true;
            this.FilterInProcessReportNO.Location = new System.Drawing.Point(40, 108);
            this.FilterInProcessReportNO.Name = "FilterInProcessReportNO";
            this.FilterInProcessReportNO.Size = new System.Drawing.Size(73, 20);
            this.FilterInProcessReportNO.TabIndex = 2;
            this.FilterInProcessReportNO.Text = "報告編號";
            // 
            // FilterInProcessgsm
            // 
            this.FilterInProcessgsm.AutoSize = true;
            this.FilterInProcessgsm.Location = new System.Drawing.Point(40, 308);
            this.FilterInProcessgsm.Name = "FilterInProcessgsm";
            this.FilterInProcessgsm.Size = new System.Drawing.Size(73, 20);
            this.FilterInProcessgsm.TabIndex = 2;
            this.FilterInProcessgsm.Text = "需求克數";
            // 
            // FilterInProcessOrder
            // 
            this.FilterInProcessOrder.AutoSize = true;
            this.FilterInProcessOrder.Location = new System.Drawing.Point(40, 148);
            this.FilterInProcessOrder.Name = "FilterInProcessOrder";
            this.FilterInProcessOrder.Size = new System.Drawing.Size(73, 20);
            this.FilterInProcessOrder.TabIndex = 2;
            this.FilterInProcessOrder.Text = "生產單號";
            // 
            // FilterInProcessTest
            // 
            this.FilterInProcessTest.AutoSize = true;
            this.FilterInProcessTest.Location = new System.Drawing.Point(40, 70);
            this.FilterInProcessTest.Name = "FilterInProcessTest";
            this.FilterInProcessTest.Size = new System.Drawing.Size(73, 20);
            this.FilterInProcessTest.TabIndex = 2;
            this.FilterInProcessTest.Text = "測試日期";
            // 
            // FilterInProcessProduction
            // 
            this.FilterInProcessProduction.AutoSize = true;
            this.FilterInProcessProduction.Location = new System.Drawing.Point(40, 30);
            this.FilterInProcessProduction.Name = "FilterInProcessProduction";
            this.FilterInProcessProduction.Size = new System.Drawing.Size(73, 20);
            this.FilterInProcessProduction.TabIndex = 2;
            this.FilterInProcessProduction.Text = "生產日期";
            // 
            // FilterRawPage
            // 
            this.FilterRawPage.AccessibleRole = System.Windows.Forms.AccessibleRole.OutlineButton;
            this.FilterRawPage.BackColor = System.Drawing.Color.White;
            this.FilterRawPage.Controls.Add(this.FilterRawReportNoTB);
            this.FilterRawPage.Controls.Add(this.FilterRawParticleSizeBox);
            this.FilterRawPage.Controls.Add(this.FilterRawArriveDateBox);
            this.FilterRawPage.Controls.Add(this.FilterRawTestDateBox);
            this.FilterRawPage.Controls.Add(this.FilterRawTypeBox);
            this.FilterRawPage.Controls.Add(this.FilterRawEffvalue);
            this.FilterRawPage.Controls.Add(this.FilterRawBackGround);
            this.FilterRawPage.Controls.Add(this.FilterRawConcertration);
            this.FilterRawPage.Controls.Add(this.FilterRawBackGroundBox);
            this.FilterRawPage.Controls.Add(this.FilterRawEffvalueBox);
            this.FilterRawPage.Controls.Add(this.FilterRawConcertrationBox);
            this.FilterRawPage.Controls.Add(this.FilterRawPressureBox);
            this.FilterRawPage.Controls.Add(this.FilterRawVOCsOutletBox);
            this.FilterRawPage.Controls.Add(this.FilterRawVOCsInletBox);
            this.FilterRawPage.Controls.Add(this.FilterRawWeightBox);
            this.FilterRawPage.Controls.Add(this.FilterRawNumberBox);
            this.FilterRawPage.Controls.Add(this.FilterRawQuantityBox);
            this.FilterRawPage.Controls.Add(this.FilterRawQtyWeightBox);
            this.FilterRawPage.Controls.Add(this.FilterRawBatchNOBox);
            this.FilterRawPage.Controls.Add(this.FilterRawEff);
            this.FilterRawPage.Controls.Add(this.FilterRawPressure);
            this.FilterRawPage.Controls.Add(this.FilterRawVOCsOutlet);
            this.FilterRawPage.Controls.Add(this.FilterRawVOCsInlet);
            this.FilterRawPage.Controls.Add(this.FilterRawVOCs);
            this.FilterRawPage.Controls.Add(this.FilterRawWeight);
            this.FilterRawPage.Controls.Add(this.FilterRawParticleSize);
            this.FilterRawPage.Controls.Add(this.FilterRawNumber);
            this.FilterRawPage.Controls.Add(this.FilterRawQuantity);
            this.FilterRawPage.Controls.Add(this.FilterRawBatchNO);
            this.FilterRawPage.Controls.Add(this.FilterRawReportNo);
            this.FilterRawPage.Controls.Add(this.FilterRawType);
            this.FilterRawPage.Controls.Add(this.FilterRawArriveDate);
            this.FilterRawPage.Controls.Add(this.FilterRawTestDate);
            this.FilterRawPage.Cursor = System.Windows.Forms.Cursors.Default;
            this.FilterRawPage.Location = new System.Drawing.Point(4, 29);
            this.FilterRawPage.Name = "FilterRawPage";
            this.FilterRawPage.Padding = new System.Windows.Forms.Padding(3);
            this.FilterRawPage.Size = new System.Drawing.Size(1018, 625);
            this.FilterRawPage.TabIndex = 0;
            this.FilterRawPage.Text = "濾網原料";
            // 
            // FilterRawReportNoTB
            // 
            this.FilterRawReportNoTB.Location = new System.Drawing.Point(175, 100);
            this.FilterRawReportNoTB.Name = "FilterRawReportNoTB";
            this.FilterRawReportNoTB.Size = new System.Drawing.Size(100, 29);
            this.FilterRawReportNoTB.TabIndex = 3;
            // 
            // FilterRawParticleSizeBox
            // 
            this.FilterRawParticleSizeBox.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.FilterRawParticleSizeBox.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.目數,
            this.weight});
            this.FilterRawParticleSizeBox.Location = new System.Drawing.Point(315, 220);
            this.FilterRawParticleSizeBox.Name = "FilterRawParticleSizeBox";
            this.FilterRawParticleSizeBox.RowTemplate.Height = 24;
            this.FilterRawParticleSizeBox.Size = new System.Drawing.Size(237, 156);
            this.FilterRawParticleSizeBox.TabIndex = 16;
            // 
            // 目數
            // 
            this.目數.HeaderText = "目數";
            this.目數.Name = "目數";
            // 
            // weight
            // 
            this.weight.HeaderText = "weight";
            this.weight.Name = "weight";
            // 
            // FilterRawArriveDateBox
            // 
            this.FilterRawArriveDateBox.Checked = false;
            this.FilterRawArriveDateBox.CustomFormat = "yyyy.MM.dd";
            this.FilterRawArriveDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.FilterRawArriveDateBox.Location = new System.Drawing.Point(175, 65);
            this.FilterRawArriveDateBox.Name = "FilterRawArriveDateBox";
            this.FilterRawArriveDateBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawArriveDateBox.TabIndex = 2;
            this.FilterRawArriveDateBox.ValueChanged += new System.EventHandler(this.FilterRawArriveDateBox_ValueChanged);
            // 
            // FilterRawTestDateBox
            // 
            this.FilterRawTestDateBox.Checked = false;
            this.FilterRawTestDateBox.CustomFormat = "yyyy.MM.dd";
            this.FilterRawTestDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.FilterRawTestDateBox.Location = new System.Drawing.Point(175, 25);
            this.FilterRawTestDateBox.Name = "FilterRawTestDateBox";
            this.FilterRawTestDateBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawTestDateBox.TabIndex = 1;
            this.FilterRawTestDateBox.ValueChanged += new System.EventHandler(this.FilterRawTestDateBox_ValueChange);
            // 
            // FilterRawTypeBox
            // 
            this.FilterRawTypeBox.FormattingEnabled = true;
            this.FilterRawTypeBox.Items.AddRange(new object[] {
            "IKP101",
            "SI013",
            "SG001",
            "IER101"});
            this.FilterRawTypeBox.Location = new System.Drawing.Point(175, 135);
            this.FilterRawTypeBox.Name = "FilterRawTypeBox";
            this.FilterRawTypeBox.Size = new System.Drawing.Size(100, 28);
            this.FilterRawTypeBox.TabIndex = 4;
            this.FilterRawTypeBox.SelectedIndexChanged += new System.EventHandler(this.FilterRawTypeBox_SelectedIndexChanged);
            // 
            // FilterRawEffvalue
            // 
            this.FilterRawEffvalue.AutoSize = true;
            this.FilterRawEffvalue.Location = new System.Drawing.Point(612, 71);
            this.FilterRawEffvalue.Name = "FilterRawEffvalue";
            this.FilterRawEffvalue.Size = new System.Drawing.Size(41, 20);
            this.FilterRawEffvalue.TabIndex = 3;
            this.FilterRawEffvalue.Text = "讀值";
            // 
            // FilterRawBackGround
            // 
            this.FilterRawBackGround.AutoSize = true;
            this.FilterRawBackGround.Location = new System.Drawing.Point(495, 70);
            this.FilterRawBackGround.Name = "FilterRawBackGround";
            this.FilterRawBackGround.Size = new System.Drawing.Size(57, 20);
            this.FilterRawBackGround.TabIndex = 2;
            this.FilterRawBackGround.Text = "背景值";
            // 
            // FilterRawConcertration
            // 
            this.FilterRawConcertration.AutoSize = true;
            this.FilterRawConcertration.Location = new System.Drawing.Point(390, 70);
            this.FilterRawConcertration.Name = "FilterRawConcertration";
            this.FilterRawConcertration.Size = new System.Drawing.Size(41, 20);
            this.FilterRawConcertration.TabIndex = 2;
            this.FilterRawConcertration.Text = "濃度";
            // 
            // FilterRawBackGroundBox
            // 
            this.FilterRawBackGroundBox.Location = new System.Drawing.Point(499, 93);
            this.FilterRawBackGroundBox.Name = "FilterRawBackGroundBox";
            this.FilterRawBackGroundBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawBackGroundBox.TabIndex = 14;
            // 
            // FilterRawEffvalueBox
            // 
            this.FilterRawEffvalueBox.Cursor = System.Windows.Forms.Cursors.IBeam;
            this.FilterRawEffvalueBox.Location = new System.Drawing.Point(616, 93);
            this.FilterRawEffvalueBox.Multiline = true;
            this.FilterRawEffvalueBox.Name = "FilterRawEffvalueBox";
            this.FilterRawEffvalueBox.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.FilterRawEffvalueBox.Size = new System.Drawing.Size(129, 99);
            this.FilterRawEffvalueBox.TabIndex = 15;
            // 
            // FilterRawConcertrationBox
            // 
            this.FilterRawConcertrationBox.Location = new System.Drawing.Point(394, 93);
            this.FilterRawConcertrationBox.Name = "FilterRawConcertrationBox";
            this.FilterRawConcertrationBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawConcertrationBox.TabIndex = 13;
            // 
            // FilterRawPressureBox
            // 
            this.FilterRawPressureBox.Location = new System.Drawing.Point(390, 25);
            this.FilterRawPressureBox.Name = "FilterRawPressureBox";
            this.FilterRawPressureBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawPressureBox.TabIndex = 12;
            // 
            // FilterRawVOCsOutletBox
            // 
            this.FilterRawVOCsOutletBox.Location = new System.Drawing.Point(175, 376);
            this.FilterRawVOCsOutletBox.Name = "FilterRawVOCsOutletBox";
            this.FilterRawVOCsOutletBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawVOCsOutletBox.TabIndex = 11;
            // 
            // FilterRawVOCsInletBox
            // 
            this.FilterRawVOCsInletBox.Location = new System.Drawing.Point(175, 336);
            this.FilterRawVOCsInletBox.Name = "FilterRawVOCsInletBox";
            this.FilterRawVOCsInletBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawVOCsInletBox.TabIndex = 10;
            // 
            // FilterRawWeightBox
            // 
            this.FilterRawWeightBox.Location = new System.Drawing.Point(175, 297);
            this.FilterRawWeightBox.Name = "FilterRawWeightBox";
            this.FilterRawWeightBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawWeightBox.TabIndex = 9;
            // 
            // FilterRawNumberBox
            // 
            this.FilterRawNumberBox.Location = new System.Drawing.Point(175, 255);
            this.FilterRawNumberBox.Name = "FilterRawNumberBox";
            this.FilterRawNumberBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawNumberBox.TabIndex = 8;
            // 
            // FilterRawQuantityBox
            // 
            this.FilterRawQuantityBox.Location = new System.Drawing.Point(225, 215);
            this.FilterRawQuantityBox.Name = "FilterRawQuantityBox";
            this.FilterRawQuantityBox.Size = new System.Drawing.Size(50, 29);
            this.FilterRawQuantityBox.TabIndex = 7;
            // 
            // FilterRawQtyWeightBox
            // 
            this.FilterRawQtyWeightBox.Location = new System.Drawing.Point(175, 215);
            this.FilterRawQtyWeightBox.Name = "FilterRawQtyWeightBox";
            this.FilterRawQtyWeightBox.Size = new System.Drawing.Size(50, 29);
            this.FilterRawQtyWeightBox.TabIndex = 6;
            // 
            // FilterRawBatchNOBox
            // 
            this.FilterRawBatchNOBox.Location = new System.Drawing.Point(175, 175);
            this.FilterRawBatchNOBox.Name = "FilterRawBatchNOBox";
            this.FilterRawBatchNOBox.Size = new System.Drawing.Size(100, 29);
            this.FilterRawBatchNOBox.TabIndex = 5;
            // 
            // FilterRawEff
            // 
            this.FilterRawEff.AutoSize = true;
            this.FilterRawEff.Location = new System.Drawing.Point(310, 70);
            this.FilterRawEff.Name = "FilterRawEff";
            this.FilterRawEff.Size = new System.Drawing.Size(41, 20);
            this.FilterRawEff.TabIndex = 0;
            this.FilterRawEff.Text = "效率";
            // 
            // FilterRawPressure
            // 
            this.FilterRawPressure.AutoSize = true;
            this.FilterRawPressure.Location = new System.Drawing.Point(310, 30);
            this.FilterRawPressure.Name = "FilterRawPressure";
            this.FilterRawPressure.Size = new System.Drawing.Size(41, 20);
            this.FilterRawPressure.TabIndex = 0;
            this.FilterRawPressure.Text = "壓損";
            // 
            // FilterRawVOCsOutlet
            // 
            this.FilterRawVOCsOutlet.AutoSize = true;
            this.FilterRawVOCsOutlet.Location = new System.Drawing.Point(100, 379);
            this.FilterRawVOCsOutlet.Name = "FilterRawVOCsOutlet";
            this.FilterRawVOCsOutlet.Size = new System.Drawing.Size(57, 20);
            this.FilterRawVOCsOutlet.TabIndex = 0;
            this.FilterRawVOCsOutlet.Text = "Outlet";
            // 
            // FilterRawVOCsInlet
            // 
            this.FilterRawVOCsInlet.AutoSize = true;
            this.FilterRawVOCsInlet.Location = new System.Drawing.Point(100, 339);
            this.FilterRawVOCsInlet.Name = "FilterRawVOCsInlet";
            this.FilterRawVOCsInlet.Size = new System.Drawing.Size(43, 20);
            this.FilterRawVOCsInlet.TabIndex = 0;
            this.FilterRawVOCsInlet.Text = "Inlet";
            // 
            // FilterRawVOCs
            // 
            this.FilterRawVOCs.AutoSize = true;
            this.FilterRawVOCs.Location = new System.Drawing.Point(40, 361);
            this.FilterRawVOCs.Name = "FilterRawVOCs";
            this.FilterRawVOCs.Size = new System.Drawing.Size(51, 20);
            this.FilterRawVOCs.TabIndex = 0;
            this.FilterRawVOCs.Text = "VOCs";
            // 
            // FilterRawWeight
            // 
            this.FilterRawWeight.AutoSize = true;
            this.FilterRawWeight.Location = new System.Drawing.Point(40, 300);
            this.FilterRawWeight.Name = "FilterRawWeight";
            this.FilterRawWeight.Size = new System.Drawing.Size(89, 20);
            this.FilterRawWeight.TabIndex = 0;
            this.FilterRawWeight.Text = "測試品重量";
            // 
            // FilterRawParticleSize
            // 
            this.FilterRawParticleSize.AutoSize = true;
            this.FilterRawParticleSize.Location = new System.Drawing.Point(311, 190);
            this.FilterRawParticleSize.Name = "FilterRawParticleSize";
            this.FilterRawParticleSize.Size = new System.Drawing.Size(73, 20);
            this.FilterRawParticleSize.TabIndex = 0;
            this.FilterRawParticleSize.Text = "粒徑大小";
            // 
            // FilterRawNumber
            // 
            this.FilterRawNumber.AutoSize = true;
            this.FilterRawNumber.Location = new System.Drawing.Point(40, 250);
            this.FilterRawNumber.Name = "FilterRawNumber";
            this.FilterRawNumber.Size = new System.Drawing.Size(109, 40);
            this.FilterRawNumber.TabIndex = 0;
            this.FilterRawNumber.Text = "原料編號\n(輸入#後兩位)";
            // 
            // FilterRawQuantity
            // 
            this.FilterRawQuantity.AutoSize = true;
            this.FilterRawQuantity.Location = new System.Drawing.Point(40, 220);
            this.FilterRawQuantity.Name = "FilterRawQuantity";
            this.FilterRawQuantity.Size = new System.Drawing.Size(73, 20);
            this.FilterRawQuantity.TabIndex = 0;
            this.FilterRawQuantity.Text = "進貨數量";
            // 
            // FilterRawBatchNO
            // 
            this.FilterRawBatchNO.AutoSize = true;
            this.FilterRawBatchNO.Location = new System.Drawing.Point(40, 180);
            this.FilterRawBatchNO.Name = "FilterRawBatchNO";
            this.FilterRawBatchNO.Size = new System.Drawing.Size(73, 20);
            this.FilterRawBatchNO.TabIndex = 0;
            this.FilterRawBatchNO.Text = "進廠批號";
            // 
            // FilterRawReportNo
            // 
            this.FilterRawReportNo.AutoSize = true;
            this.FilterRawReportNo.Location = new System.Drawing.Point(40, 105);
            this.FilterRawReportNo.Name = "FilterRawReportNo";
            this.FilterRawReportNo.Size = new System.Drawing.Size(73, 20);
            this.FilterRawReportNo.TabIndex = 0;
            this.FilterRawReportNo.Text = "報告編號";
            // 
            // FilterRawType
            // 
            this.FilterRawType.AutoSize = true;
            this.FilterRawType.Location = new System.Drawing.Point(40, 140);
            this.FilterRawType.Name = "FilterRawType";
            this.FilterRawType.Size = new System.Drawing.Size(73, 20);
            this.FilterRawType.TabIndex = 0;
            this.FilterRawType.Text = "原料種類";
            // 
            // FilterRawArriveDate
            // 
            this.FilterRawArriveDate.AutoSize = true;
            this.FilterRawArriveDate.Location = new System.Drawing.Point(40, 70);
            this.FilterRawArriveDate.Name = "FilterRawArriveDate";
            this.FilterRawArriveDate.Size = new System.Drawing.Size(73, 20);
            this.FilterRawArriveDate.TabIndex = 0;
            this.FilterRawArriveDate.Text = "到廠日期";
            // 
            // FilterRawTestDate
            // 
            this.FilterRawTestDate.AutoSize = true;
            this.FilterRawTestDate.Location = new System.Drawing.Point(40, 30);
            this.FilterRawTestDate.Name = "FilterRawTestDate";
            this.FilterRawTestDate.Size = new System.Drawing.Size(73, 20);
            this.FilterRawTestDate.TabIndex = 0;
            this.FilterRawTestDate.Text = "測試日期";
            // 
            // tabControl1
            // 
            this.tabControl1.AllowDrop = true;
            this.tabControl1.Controls.Add(this.FilterRawPage);
            this.tabControl1.Controls.Add(this.FilterInProcessPage);
            this.tabControl1.Controls.Add(this.FilterPage);
            this.tabControl1.Controls.Add(this.CylinderRawPage);
            this.tabControl1.Controls.Add(this.CylinderPage);
            this.tabControl1.Controls.Add(this.RawMaterialPage);
            this.tabControl1.Cursor = System.Windows.Forms.Cursors.Hand;
            this.tabControl1.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tabControl1.Location = new System.Drawing.Point(1, 71);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1026, 658);
            this.tabControl1.TabIndex = 2;
            // 
            // RawMaterialPage
            // 
            this.RawMaterialPage.Controls.Add(this.RawMaterialdgv);
            this.RawMaterialPage.Controls.Add(this.MaterialTypeTB);
            this.RawMaterialPage.Controls.Add(this.label12);
            this.RawMaterialPage.Controls.Add(this.MaterialTestDateBox);
            this.RawMaterialPage.Controls.Add(this.RawMaterialNOtb);
            this.RawMaterialPage.Controls.Add(this.label23);
            this.RawMaterialPage.Controls.Add(this.MaterialReportNOTB);
            this.RawMaterialPage.Controls.Add(this.label22);
            this.RawMaterialPage.Controls.Add(this.label21);
            this.RawMaterialPage.Location = new System.Drawing.Point(4, 29);
            this.RawMaterialPage.Name = "RawMaterialPage";
            this.RawMaterialPage.Padding = new System.Windows.Forms.Padding(3);
            this.RawMaterialPage.Size = new System.Drawing.Size(1018, 625);
            this.RawMaterialPage.TabIndex = 5;
            this.RawMaterialPage.Text = "物料";
            this.RawMaterialPage.UseVisualStyleBackColor = true;
            // 
            // RawMaterialdgv
            // 
            this.RawMaterialdgv.AllowUserToResizeColumns = false;
            this.RawMaterialdgv.AllowUserToResizeRows = false;
            this.RawMaterialdgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.RawMaterialdgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11});
            this.RawMaterialdgv.Cursor = System.Windows.Forms.Cursors.Default;
            this.RawMaterialdgv.Location = new System.Drawing.Point(44, 164);
            this.RawMaterialdgv.Name = "RawMaterialdgv";
            this.RawMaterialdgv.RowTemplate.Height = 24;
            this.RawMaterialdgv.Size = new System.Drawing.Size(929, 341);
            this.RawMaterialdgv.TabIndex = 5;
            // 
            // MaterialTypeTB
            // 
            this.MaterialTypeTB.Cursor = System.Windows.Forms.Cursors.Default;
            this.MaterialTypeTB.FormattingEnabled = true;
            this.MaterialTypeTB.Items.AddRange(new object[] {
            "濾筒",
            "濾網"});
            this.MaterialTypeTB.Location = new System.Drawing.Point(175, 25);
            this.MaterialTypeTB.Name = "MaterialTypeTB";
            this.MaterialTypeTB.Size = new System.Drawing.Size(100, 28);
            this.MaterialTypeTB.TabIndex = 1;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(40, 30);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(73, 20);
            this.label12.TabIndex = 25;
            this.label12.Text = "物料種類";
            // 
            // MaterialTestDateBox
            // 
            this.MaterialTestDateBox.Checked = false;
            this.MaterialTestDateBox.Cursor = System.Windows.Forms.Cursors.Default;
            this.MaterialTestDateBox.CustomFormat = "yyyy.MM.dd";
            this.MaterialTestDateBox.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.MaterialTestDateBox.Location = new System.Drawing.Point(175, 59);
            this.MaterialTestDateBox.Name = "MaterialTestDateBox";
            this.MaterialTestDateBox.Size = new System.Drawing.Size(100, 29);
            this.MaterialTestDateBox.TabIndex = 2;
            this.MaterialTestDateBox.ValueChanged += new System.EventHandler(this.MaterialTestDateBox_ValueChanged);
            // 
            // RawMaterialNOtb
            // 
            this.RawMaterialNOtb.Cursor = System.Windows.Forms.Cursors.Default;
            this.RawMaterialNOtb.Location = new System.Drawing.Point(175, 129);
            this.RawMaterialNOtb.Name = "RawMaterialNOtb";
            this.RawMaterialNOtb.Size = new System.Drawing.Size(164, 29);
            this.RawMaterialNOtb.TabIndex = 4;
            this.RawMaterialNOtb.TextChanged += new System.EventHandler(this.CylinderRawTypeBox_SelectedIndexChanged);
            this.RawMaterialNOtb.KeyDown += new System.Windows.Forms.KeyEventHandler(this.RawMaterialNOtb_keyDown);
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(40, 132);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(41, 20);
            this.label23.TabIndex = 22;
            this.label23.Text = "料號";
            // 
            // MaterialReportNOTB
            // 
            this.MaterialReportNOTB.Cursor = System.Windows.Forms.Cursors.Default;
            this.MaterialReportNOTB.Location = new System.Drawing.Point(175, 94);
            this.MaterialReportNOTB.Name = "MaterialReportNOTB";
            this.MaterialReportNOTB.Size = new System.Drawing.Size(164, 29);
            this.MaterialReportNOTB.TabIndex = 3;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(40, 97);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(73, 20);
            this.label22.TabIndex = 22;
            this.label22.Text = "報告編號";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(40, 64);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(73, 20);
            this.label21.TabIndex = 23;
            this.label21.Text = "測試日期";
            // 
            // execute
            // 
            this.execute.Cursor = System.Windows.Forms.Cursors.Hand;
            this.execute.Font = new System.Drawing.Font("微軟正黑體", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.execute.Location = new System.Drawing.Point(814, 12);
            this.execute.Name = "execute";
            this.execute.Size = new System.Drawing.Size(121, 53);
            this.execute.TabIndex = 3;
            this.execute.Text = "執行";
            this.execute.UseVisualStyleBackColor = true;
            this.execute.Click += new System.EventHandler(this.Execute_Click);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.dataGridViewTextBoxColumn1.HeaderText = "進貨日期";
            this.dataGridViewTextBoxColumn1.MinimumWidth = 50;
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.dataGridViewTextBoxColumn2.HeaderText = "料號/物料名稱";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 250;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.HeaderText = "進貨數量";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.HeaderText = "單位";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Width = 70;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.HeaderText = "抽檢數";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Width = 65;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.HeaderText = "單位";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.Width = 70;
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.HeaderText = "外觀";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            this.dataGridViewTextBoxColumn7.Width = 70;
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.HeaderText = "規格值";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            this.dataGridViewTextBoxColumn8.Width = 150;
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.HeaderText = "測量值";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            this.dataGridViewTextBoxColumn9.Width = 150;
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.HeaderText = "合格/不合格";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            this.dataGridViewTextBoxColumn10.Width = 130;
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.HeaderText = "備註";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            this.dataGridViewTextBoxColumn11.Width = 140;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1010, 706);
            this.Controls.Add(this.execute);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "QC";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.CylinderPage.ResumeLayout(false);
            this.CylinderPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CylinderBox)).EndInit();
            this.CylinderRawPage.ResumeLayout(false);
            this.CylinderRawPage.PerformLayout();
            this.CylinderRawEffPanel.ResumeLayout(false);
            this.CylinderRawEffPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.CylinderRawMeshBox)).EndInit();
            this.FilterPage.ResumeLayout(false);
            this.FilterPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FilterBox)).EndInit();
            this.FilterInProcessPage.ResumeLayout(false);
            this.FilterInProcessPage.PerformLayout();
            this.FilterInProcessEffPanel.ResumeLayout(false);
            this.FilterInProcessEffPanel.PerformLayout();
            this.FilterRawPage.ResumeLayout(false);
            this.FilterRawPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.FilterRawParticleSizeBox)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.RawMaterialPage.ResumeLayout(false);
            this.RawMaterialPage.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RawMaterialdgv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabPage CylinderPage;
        private System.Windows.Forms.DateTimePicker CylinderTestDateBox;
        private System.Windows.Forms.DataGridView CylinderBox;
        private System.Windows.Forms.TextBox ReCylinderNOBox;
        private System.Windows.Forms.TextBox CylinderCustmorBox;
        private System.Windows.Forms.TextBox CylinderNOBox;
        private System.Windows.Forms.TextBox CylinderReportNOBox;
        private System.Windows.Forms.Label ReCylinderNO;
        private System.Windows.Forms.Label CylinderCustmor;
        private System.Windows.Forms.Label CylinderNO;
        private System.Windows.Forms.Label CylinderReportNO;
        private System.Windows.Forms.Label CylinderTestDate;
        private System.Windows.Forms.TabPage CylinderRawPage;
        private System.Windows.Forms.TableLayoutPanel CylinderRawEffPanel;
        private System.Windows.Forms.Label CylinderRawH2S;
        private System.Windows.Forms.Label CylinderRawSO2;
        private System.Windows.Forms.Label CylinderRawConcertration;
        private System.Windows.Forms.Label CylinderRawGas;
        private System.Windows.Forms.Label label76;
        private System.Windows.Forms.Label CylinderRawBackGround;
        private System.Windows.Forms.Label CylinderRawValue;
        private System.Windows.Forms.CheckBox CylinderRawSO2CheckBox;
        private System.Windows.Forms.CheckBox CylinderRawH2SCheckBox;
        private System.Windows.Forms.Label CylinderRawEff;
        private System.Windows.Forms.DataGridView CylinderRawMeshBox;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn12;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn13;
        private System.Windows.Forms.Label CylinderRawParticle;
        private System.Windows.Forms.TabPage FilterPage;
        private System.Windows.Forms.ComboBox FilterModelBox;
        private System.Windows.Forms.DateTimePicker FilterTestDateBox;
        private System.Windows.Forms.DataGridView FilterBox;
        private System.Windows.Forms.TextBox ReFilterNOBox;
        private System.Windows.Forms.TextBox FilterPackageNOBox;
        private System.Windows.Forms.TextBox FilterReportBox;
        private System.Windows.Forms.Label ReFilterNO;
        private System.Windows.Forms.Label FilterModel;
        private System.Windows.Forms.Label FilterReportCustmor;
        private System.Windows.Forms.Label FilterPackageNO;
        private System.Windows.Forms.Label FilterReport;
        private System.Windows.Forms.Label FilterTestDate;
        private System.Windows.Forms.TabPage FilterInProcessPage;
        private System.Windows.Forms.TableLayoutPanel FilterInProcessEffPanel;
        private System.Windows.Forms.Label FilterInValue;
        private System.Windows.Forms.Label label35;
        private System.Windows.Forms.Label FilterInGas;
        private System.Windows.Forms.Label FilterInSO2Concertration;
        private System.Windows.Forms.Label FilterInBackGround;
        private System.Windows.Forms.CheckBox FilterInSO2CheckBox;
        private System.Windows.Forms.CheckBox FilterInH2SCheckBox;
        private System.Windows.Forms.CheckBox FilterInNH3CheckBox;
        private System.Windows.Forms.Label FilterInToluene;
        private System.Windows.Forms.Label FilterInNH3;
        private System.Windows.Forms.Label FilterInH2S;
        private System.Windows.Forms.Label FilterInSO2;
        private System.Windows.Forms.ComboBox FilterInProcessTypeBox;
        private System.Windows.Forms.DateTimePicker FilterInProcessTestBox;
        private System.Windows.Forms.DateTimePicker FilterInProcessProductionBox;
        private System.Windows.Forms.TextBox FilterInProcessPressureDropBox;
        private System.Windows.Forms.TextBox FilterInProcessTestGsmBox;
        private System.Windows.Forms.TextBox FilterInProcessWindBox;
        private System.Windows.Forms.TextBox FilterInProcessPressureBox;
        private System.Windows.Forms.TextBox FilterInProcessLowerBox;
        private System.Windows.Forms.TextBox FilterInProcessUpperBox;
        private System.Windows.Forms.TextBox FilterInProcessSpeedBox;
        private System.Windows.Forms.TextBox FilterInProcessGileBox;
        private System.Windows.Forms.TextBox FilterInProcessgsmBox;
        private System.Windows.Forms.TextBox FilterInProcessOrderBox;
        private System.Windows.Forms.Label FilterInProcessWire;
        private System.Windows.Forms.Label FilterInProcessEff;
        private System.Windows.Forms.Label FilterInProcessTestGsm;
        private System.Windows.Forms.Label FilterInProcessWind;
        private System.Windows.Forms.Label FilterInProcessPressure;
        private System.Windows.Forms.Label FilterInProcessLower;
        private System.Windows.Forms.Label FilterInProcessUpper;
        private System.Windows.Forms.Label FilterInProcessSpeed;
        private System.Windows.Forms.Label FilterInProcessGile;
        private System.Windows.Forms.Label FilterInProcessType;
        private System.Windows.Forms.Label FilterInProcessgsm;
        private System.Windows.Forms.Label FilterInProcessOrder;
        private System.Windows.Forms.Label FilterInProcessTest;
        private System.Windows.Forms.Label FilterInProcessProduction;
        private System.Windows.Forms.TabPage FilterRawPage;
        private System.Windows.Forms.DataGridView FilterRawParticleSizeBox;
        private System.Windows.Forms.DataGridViewTextBoxColumn 目數;
        private System.Windows.Forms.DataGridViewTextBoxColumn weight;
        private System.Windows.Forms.DateTimePicker FilterRawArriveDateBox;
        private System.Windows.Forms.DateTimePicker FilterRawTestDateBox;
        private System.Windows.Forms.ComboBox FilterRawTypeBox;
        private System.Windows.Forms.Label FilterRawEffvalue;
        private System.Windows.Forms.Label FilterRawBackGround;
        private System.Windows.Forms.Label FilterRawConcertration;
        private System.Windows.Forms.TextBox FilterRawBackGroundBox;
        private System.Windows.Forms.TextBox FilterRawEffvalueBox;
        private System.Windows.Forms.TextBox FilterRawConcertrationBox;
        private System.Windows.Forms.TextBox FilterRawPressureBox;
        private System.Windows.Forms.TextBox FilterRawVOCsOutletBox;
        private System.Windows.Forms.TextBox FilterRawVOCsInletBox;
        private System.Windows.Forms.TextBox FilterRawWeightBox;
        private System.Windows.Forms.TextBox FilterRawNumberBox;
        private System.Windows.Forms.TextBox FilterRawQtyWeightBox;
        private System.Windows.Forms.TextBox FilterRawBatchNOBox;
        private System.Windows.Forms.Label FilterRawEff;
        private System.Windows.Forms.Label FilterRawPressure;
        private System.Windows.Forms.Label FilterRawVOCsOutlet;
        private System.Windows.Forms.Label FilterRawVOCsInlet;
        private System.Windows.Forms.Label FilterRawVOCs;
        private System.Windows.Forms.Label FilterRawWeight;
        private System.Windows.Forms.Label FilterRawParticleSize;
        private System.Windows.Forms.Label FilterRawNumber;
        private System.Windows.Forms.Label FilterRawQuantity;
        private System.Windows.Forms.Label FilterRawBatchNO;
        private System.Windows.Forms.Label FilterRawType;
        private System.Windows.Forms.Label FilterRawArriveDate;
        private System.Windows.Forms.Label FilterRawTestDate;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.CheckBox CylinderRawNH3CheckBox;
        private System.Windows.Forms.CheckBox CylinderRawTolueneCheckBox;
        private System.Windows.Forms.Label CylinderRawNH3;
        private System.Windows.Forms.Label CylinderRawToluene;
        private System.Windows.Forms.CheckBox CylinderRawIPACheckBox;
        private System.Windows.Forms.CheckBox CylinderRawAcetoneCheckBox;
        private System.Windows.Forms.Label CylinderRawIPA;
        private System.Windows.Forms.Label CylinderRawAcetone;
        private System.Windows.Forms.Button execute;
        private System.Windows.Forms.TextBox CylinderRawSO2ConcertrationBox;
        private System.Windows.Forms.CheckBox FilterInTolueneCheckBox;
        private System.Windows.Forms.TextBox CylinderRawSO2BackGroundBox;
        private System.Windows.Forms.TextBox CylinderRawH2SBackGroundBox;
        private System.Windows.Forms.TextBox CylinderRawNH3BackGroundBox;
        private System.Windows.Forms.TextBox CylinderRawTolueneBackGroundBox;
        private System.Windows.Forms.TextBox CylinderRawIPABackGroundBox;
        private System.Windows.Forms.TextBox CylinderRawAcetoneBackGroundBox;
        private System.Windows.Forms.TextBox CylinderRawAcetoneConcertrationBox;
        private System.Windows.Forms.TextBox CylinderRawIPAConcertrationBox;
        private System.Windows.Forms.TextBox CylinderRawTolueneConcertrationBox;
        private System.Windows.Forms.TextBox CylinderRawNH3ConcertrationBox;
        private System.Windows.Forms.TextBox CylinderRawH2SConcertrationBox;
        private System.Windows.Forms.TextBox CylinderRawSO2ValueBox;
        private System.Windows.Forms.TextBox CylinderRawH2SValueBox;
        private System.Windows.Forms.TextBox CylinderRawNH3ValueBox;
        private System.Windows.Forms.TextBox CylinderRawTolueneValueBox;
        private System.Windows.Forms.TextBox CylinderRawIPAValueBox;
        private System.Windows.Forms.TextBox CylinderRawAcetoneValueBox;
        private System.Windows.Forms.TextBox FilterInSO2ConcentrationBox;
        private System.Windows.Forms.TextBox FilterInH2SConcentrationBox;
        private System.Windows.Forms.TextBox FilterInNH3ConcentrationBox;
        private System.Windows.Forms.TextBox FilterInBackGroundTolueneBox;
        private System.Windows.Forms.TextBox FilterInBackGroundNH3Box;
        private System.Windows.Forms.TextBox FilterInBackGroundH2SBox;
        private System.Windows.Forms.TextBox FilterInBackGroundSO2Box;
        private System.Windows.Forms.TextBox FilterInTolueneConcentrationBox;
        private System.Windows.Forms.TextBox FilterInValueSO2Box;
        private System.Windows.Forms.TextBox FilterInValueH2SBox;
        private System.Windows.Forms.TextBox FilterInValueNH3Box;
        private System.Windows.Forms.TextBox FilterInValueTolueneBox;
        private System.Windows.Forms.TextBox FilterMaterialNOBox;
        private System.Windows.Forms.Label FilterMaterialNO;
        private System.Windows.Forms.Label CylinderMaterialSN;
        private System.Windows.Forms.TextBox FilterCarbonLotBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox FilterRawQuantityBox;
        private System.Windows.Forms.TextBox CYLMaterialSNBox;
        private System.Windows.Forms.Label CYLType;
        private System.Windows.Forms.TextBox FilterInProcessCarbonOrderBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker FilterProductionBox;
        private System.Windows.Forms.Label FilterProduction;
        private System.Windows.Forms.TextBox FilterAlarmBox;
        private System.Windows.Forms.Label FilterAlarm;
        private System.Windows.Forms.ComboBox CYLTypeBox;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYLSN;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYLWeight;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Particle_In;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Particle_out;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_IPA_in;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_IPA_out;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Acetone_In;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Acetone_out;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Nontarget_in;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Nontarget_out;
        private System.Windows.Forms.DataGridViewTextBoxColumn CYL_Pressure_Drop;
        private System.Windows.Forms.ComboBox FilterReportCustmorBox;
        private System.Windows.Forms.TextBox FilterMaterialNumerBox;
        private System.Windows.Forms.Label FilterMaterialNumer;
        private System.Windows.Forms.DataGridViewTextBoxColumn 生產序號;
        private System.Windows.Forms.DataGridViewTextBoxColumn 重量;
        private System.Windows.Forms.DataGridViewTextBoxColumn length;
        private System.Windows.Forms.DataGridViewTextBoxColumn width;
        private System.Windows.Forms.DataGridViewTextBoxColumn height;
        private System.Windows.Forms.DataGridViewTextBoxColumn diagonal;
        private System.Windows.Forms.DataGridViewTextBoxColumn Particle_In;
        private System.Windows.Forms.DataGridViewTextBoxColumn Particle_Out;
        private System.Windows.Forms.DataGridViewTextBoxColumn IPA_In;
        private System.Windows.Forms.DataGridViewTextBoxColumn IPA_Out;
        private System.Windows.Forms.DataGridViewTextBoxColumn Acetone_In;
        private System.Windows.Forms.DataGridViewTextBoxColumn Acetone_Out;
        private System.Windows.Forms.DataGridViewTextBoxColumn Nontarget_In;
        private System.Windows.Forms.DataGridViewTextBoxColumn Nontarget_Out;
        private System.Windows.Forms.DataGridViewTextBoxColumn Pressure_Drop_spec;
        private System.Windows.Forms.DataGridViewTextBoxColumn Pressure_Drop;
        private System.Windows.Forms.TextBox FilterInProcessCarbonInfoBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox FilterInValueIPABox;
        private System.Windows.Forms.TextBox FilterInValueAcetoneBox;
        private System.Windows.Forms.TextBox FilterInIPAConcentrationBox;
        private System.Windows.Forms.TextBox FilterInAcetoneConcentrationBox;
        private System.Windows.Forms.TextBox FilterInBackGroundIPABox;
        private System.Windows.Forms.TextBox FilterInBackGroundAcetoneBox;
        private System.Windows.Forms.TextBox CylinderRawPressureBox;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.DateTimePicker CylinderRawArriveDateBox;
        private System.Windows.Forms.DateTimePicker CylinderRawTestDateBox;
        private System.Windows.Forms.ComboBox CylinderRawTypeBox;
        private System.Windows.Forms.TextBox CylinderRawVOCsOutletBox;
        private System.Windows.Forms.TextBox CylinderRawVOCsInletBox;
        private System.Windows.Forms.TextBox CylinderRawMoistureTB;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox CylinderRawAshTB;
        private System.Windows.Forms.TextBox CylinderRawWeightBox;
        private System.Windows.Forms.TextBox CylinderRawNumberBox;
        private System.Windows.Forms.TextBox CylinderRawQtyPackBox;
        private System.Windows.Forms.TextBox CylinderRawQtyWeightBox;
        private System.Windows.Forms.TextBox CylinderRawLotBox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.CheckBox chkAsh;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox chkMoisture;
        private System.Windows.Forms.Label FilterRawReportNo;
        private System.Windows.Forms.TextBox FilterRawReportNoTB;
        private System.Windows.Forms.TextBox CylinderRawReportNoTB;
        private System.Windows.Forms.Label CylinderRawReportNo;
        private System.Windows.Forms.TextBox FilterInProcessReportNOTB;
        private System.Windows.Forms.Label FilterInProcessReportNO;
        private System.Windows.Forms.TextBox FilterSizeTB;
        private System.Windows.Forms.Label FilterSize;
        private System.Windows.Forms.TabPage RawMaterialPage;
        private System.Windows.Forms.DateTimePicker MaterialTestDateBox;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.ComboBox MaterialTypeTB;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox MaterialReportNOTB;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.TextBox RawMaterialNOtb;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.DataGridView RawMaterialdgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
    }
}

