using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Tester.
	/// </summary>
	public class Tester : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Splitter splitter2;
		private System.Windows.Forms.Panel panel3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private C1.Data.C1DataSet cdsTester;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTesterGrupo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTesterGrupo1;
		private Infragistics.Win.Misc.UltraButton btProcedimiento;
		private Infragistics.Win.Misc.UltraButton btGrabaProceso;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnSeveridad;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btReporte;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private System.Windows.Forms.CheckBox chkCobros;
		private System.ComponentModel.IContainer components;

		public Tester()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TesterGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTesterGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ErrorSalida", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idErrorSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mensaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Severidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("TesterGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTesterGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Tester", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTester");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Instrucciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ejecutar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Errores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comentario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTesterGrupo", -1, "cmbTesterGrupo1", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden", -1, null, 1, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.panel1 = new System.Windows.Forms.Panel();
			this.chkCobros = new System.Windows.Forms.CheckBox();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btReporte = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.spnSeveridad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btGrabaProceso = new Infragistics.Win.Misc.UltraButton();
			this.btProcedimiento = new Infragistics.Win.Misc.UltraButton();
			this.cmbTesterGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsTester = new C1.Data.C1DataSet();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.panel2 = new System.Windows.Forms.Panel();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.splitter2 = new System.Windows.Forms.Splitter();
			this.panel3 = new System.Windows.Forms.Panel();
			this.cmbTesterGrupo1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSeveridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTesterGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTester)).BeginInit();
			this.panel2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.panel3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTesterGrupo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.chkCobros);
			this.panel1.Controls.Add(this.cmbHasta);
			this.panel1.Controls.Add(this.cmbDesde);
			this.panel1.Controls.Add(this.btReporte);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Controls.Add(this.spnSeveridad);
			this.panel1.Controls.Add(this.ultraButton1);
			this.panel1.Controls.Add(this.btGrabaProceso);
			this.panel1.Controls.Add(this.btProcedimiento);
			this.panel1.Controls.Add(this.cmbTesterGrupo);
			this.panel1.Controls.Add(this.btGenerar);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(784, 72);
			this.panel1.TabIndex = 0;
			// 
			// chkCobros
			// 
			this.chkCobros.Location = new System.Drawing.Point(552, 40);
			this.chkCobros.Name = "chkCobros";
			this.chkCobros.Size = new System.Drawing.Size(64, 24);
			this.chkCobros.TabIndex = 12;
			this.chkCobros.Text = "Cobros";
			// 
			// cmbHasta
			// 
			dateButton1.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton1);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(552, 8);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(96, 21);
			this.cmbHasta.TabIndex = 11;
			// 
			// cmbDesde
			// 
			dateButton2.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton2);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(448, 8);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(96, 21);
			this.cmbDesde.TabIndex = 10;
			// 
			// btReporte
			// 
			this.btReporte.Location = new System.Drawing.Point(456, 40);
			this.btReporte.Name = "btReporte";
			this.btReporte.TabIndex = 9;
			this.btReporte.Text = "Reporte";
			this.btReporte.Click += new System.EventHandler(this.btReporte_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(216, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(55, 16);
			this.label1.TabIndex = 8;
			this.label1.Text = "Severidad";
			// 
			// spnSeveridad
			// 
			this.spnSeveridad.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeteo, "Seteo.Tester"));
			this.spnSeveridad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnSeveridad.Location = new System.Drawing.Point(288, 40);
			this.spnSeveridad.MaskInput = "nnn";
			this.spnSeveridad.MaxValue = 100;
			this.spnSeveridad.MinValue = 0;
			this.spnSeveridad.Name = "spnSeveridad";
			this.spnSeveridad.PromptChar = ' ';
			this.spnSeveridad.Size = new System.Drawing.Size(72, 21);
			this.spnSeveridad.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnSeveridad.TabIndex = 7;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(400, 40);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 23);
			this.ultraButton1.TabIndex = 6;
			this.ultraButton1.Text = "X";
			this.toolTip1.SetToolTip(this.ultraButton1, "Graba Proceso");
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btGrabaProceso
			// 
			this.btGrabaProceso.Location = new System.Drawing.Point(400, 8);
			this.btGrabaProceso.Name = "btGrabaProceso";
			this.btGrabaProceso.Size = new System.Drawing.Size(24, 23);
			this.btGrabaProceso.TabIndex = 5;
			this.btGrabaProceso.Text = "P";
			this.toolTip1.SetToolTip(this.btGrabaProceso, "Graba Proceso");
			this.btGrabaProceso.Click += new System.EventHandler(this.btGrabaProceso_Click);
			// 
			// btProcedimiento
			// 
			this.btProcedimiento.Location = new System.Drawing.Point(368, 8);
			this.btProcedimiento.Name = "btProcedimiento";
			this.btProcedimiento.Size = new System.Drawing.Size(24, 23);
			this.btProcedimiento.TabIndex = 4;
			this.btProcedimiento.Text = "...";
			this.btProcedimiento.Click += new System.EventHandler(this.btProcedimiento_Click);
			// 
			// cmbTesterGrupo
			// 
			this.cmbTesterGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTesterGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTesterGrupo.DataMember = "TesterGrupo";
			this.cmbTesterGrupo.DataSource = this.cdsTester;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTesterGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTesterGrupo.DisplayMember = "Grupo";
			this.cmbTesterGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTesterGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTesterGrupo.Location = new System.Drawing.Point(216, 8);
			this.cmbTesterGrupo.Name = "cmbTesterGrupo";
			this.cmbTesterGrupo.Size = new System.Drawing.Size(144, 21);
			this.cmbTesterGrupo.TabIndex = 2;
			this.cmbTesterGrupo.ValueMember = "idTesterGrupo";
			this.cmbTesterGrupo.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTesterGrupo_Validating);
			this.cmbTesterGrupo.Validated += new System.EventHandler(this.cmbTesterGrupo_Validated);
			this.cmbTesterGrupo.ValueChanged += new System.EventHandler(this.cmbTesterGrupo_ValueChanged);
			this.cmbTesterGrupo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTesterGrupo_InitializeLayout);
			// 
			// cdsTester
			// 
			this.cdsTester.BindingContextCtrl = this;
			this.cdsTester.DataLibrary = "LibPersonal";
			this.cdsTester.DataLibraryUrl = "";
			this.cdsTester.DataSetDef = "dsTester";
			this.cdsTester.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTester.Name = "cdsTester";
			this.cdsTester.SchemaClassName = "LibPersonal.DataClass";
			this.cdsTester.SchemaDef = null;
			this.cdsTester.BeforeFill += new C1.Data.FillEventHandler(this.cdsTester_BeforeFill);
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(16, 8);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.TabIndex = 0;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// splitter1
			// 
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Top;
			this.splitter1.Location = new System.Drawing.Point(0, 72);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(784, 3);
			this.splitter1.TabIndex = 1;
			this.splitter1.TabStop = false;
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.ultraGrid2);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panel2.Location = new System.Drawing.Point(0, 254);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(784, 128);
			this.panel2.TabIndex = 2;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataMember = "ErrorSalida";
			this.ultraGrid2.DataSource = this.cdsTester;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 473;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 172;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 118;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid2.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid2.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid2.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid2.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid2.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid2.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid2.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(784, 128);
			this.ultraGrid2.TabIndex = 0;
			this.ultraGrid2.Text = "Salida";
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			// 
			// splitter2
			// 
			this.splitter2.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.splitter2.Location = new System.Drawing.Point(0, 251);
			this.splitter2.Name = "splitter2";
			this.splitter2.Size = new System.Drawing.Size(784, 3);
			this.splitter2.TabIndex = 3;
			this.splitter2.TabStop = false;
			// 
			// panel3
			// 
			this.panel3.Controls.Add(this.cmbTesterGrupo1);
			this.panel3.Controls.Add(this.ultraGrid1);
			this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel3.Location = new System.Drawing.Point(0, 75);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(784, 176);
			this.panel3.TabIndex = 4;
			// 
			// cmbTesterGrupo1
			// 
			this.cmbTesterGrupo1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTesterGrupo1.DataMember = "TesterGrupo";
			this.cmbTesterGrupo1.DataSource = this.cdsTester;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 181;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbTesterGrupo1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTesterGrupo1.DisplayMember = "Grupo";
			this.cmbTesterGrupo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTesterGrupo1.Location = new System.Drawing.Point(112, 80);
			this.cmbTesterGrupo1.Name = "cmbTesterGrupo1";
			this.cmbTesterGrupo1.Size = new System.Drawing.Size(224, 72);
			this.cmbTesterGrupo1.TabIndex = 1;
			this.cmbTesterGrupo1.ValueMember = "idTesterGrupo";
			this.cmbTesterGrupo1.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Tester";
			this.ultraGrid1.DataSource = this.cdsTester;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn11.CellMultiLine = Infragistics.Win.DefaultableBoolean.True;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 427;
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Width = 225;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "Grupo";
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn16.Width = 57;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Width = 54;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 10F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Fixed;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.ultraGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 0);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(784, 176);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.Text = "Procedimiento";
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// Tester
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(784, 382);
			this.Controls.Add(this.panel3);
			this.Controls.Add(this.splitter2);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.panel1);
			this.Name = "Tester";
			this.Text = "Tester";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Tester_Closing);
			this.Load += new System.EventHandler(this.Tester_Load);
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnSeveridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTesterGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTester)).EndInit();
			this.panel2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.panel3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTesterGrupo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void FiltroProcedimiento()
		{
			if (cmbTesterGrupo.Value == null) return;
			int iGrupo = (int) cmbTesterGrupo.Value;
			string stFiltro = "idTesterGrupo > 0";
			if (cmbTesterGrupo.Text.Trim() != "(Todos)")
				stFiltro = "idTesterGrupo = " + cmbTesterGrupo.Value.ToString();

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(cdsTester, "Tester", stFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			cdsTester.Fill(fcTotal, false);
			cmbTesterGrupo.Value = iGrupo;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Selected.Rows)
			{
				string stProcedimiento = dr.Cells["Instrucciones"].Value.ToString();
				Funcion.EjecutaSQL(cdsTester, stProcedimiento, true);
			}
			cdsTester.Clear();
			cdsTester.Fill();
			FiltroProcedimiento();
			ultraGrid2.Refresh();
		}

		private void btProcedimiento_Click(object sender, System.EventArgs e)
		{
			TesterGrupo miGrupo = new TesterGrupo();
			miGrupo.MdiParent = this.MdiParent;
			miGrupo.Show();
		}

		private void btGrabaProceso_Click(object sender, System.EventArgs e)
		{
			cdsTester.Update();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsTester, "Delete ErrorSalida", true);
			cdsTester.Clear();
			cdsTester.Fill();
			ultraGrid2.Refresh();
		}

		private void cdsTester_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsTester.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbTesterGrupo_ValueChanged(object sender, System.EventArgs e)
		{
//			MessageBox.Show(cmbTesterGrupo.Text);
		}

		private void Tester_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Tester'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);

			cmbTesterGrupo.Text = "(Todos)";
		}

		private void cmbTesterGrupo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}

		private void cmbTesterGrupo_Validated(object sender, System.EventArgs e)
		{
			FiltroProcedimiento();
		
		}

		private void cmbTesterGrupo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btReporte_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaIni = (DateTime) this.cmbDesde.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaIni";
			discreteVal1.Value = dtFechaIni.ToString("yyyy-MM-dd hh:mm:ss");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);
		
			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaFin";
			discreteVal2.Value = dtFechaFin.ToString("yyyy-MM-dd hh:mm:ss");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@Signo";
			int iSigno = 1;
			if (chkCobros.Checked) iSigno = -1;
			discreteVal3.Value = iSigno;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			string stReporte = "SaldosPeriodo.rpt";
			Reporte miRep = new Reporte(stReporte);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
		}

		private void Tester_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Tester'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid2.ActiveCell == null) return;
			try
			{
				if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid2.ActiveCell.DroppedDown) return;

						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid2.ActiveCell.DroppedDown) return;
			
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid2.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid2.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid2.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid2.ActiveCell.DroppedDown == false)
							ultraGrid2.ActiveCell.DroppedDown = true;
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
			}	
		}

	}
}
