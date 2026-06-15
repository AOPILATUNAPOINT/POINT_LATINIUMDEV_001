using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmIngresCredito.
	/// </summary>
	public class frmIngresCredito : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRegion;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIngreso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCobrador;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSalida;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargo;
		private System.Windows.Forms.Label cm;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Button btnAPP;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmIngresCredito()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRegion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmIngresCredito));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbRegion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdCobrador = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFechaSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cm = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAPP = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbRegion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCobrador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			this.SuspendLayout();
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(256, 40);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(40, 16);
			this.label19.TabIndex = 319;
			this.label19.Text = "Región";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbRegion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRegion.Appearance = appearance1;
			this.cmbRegion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRegion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRegion.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 245;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbRegion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbRegion.DisplayMember = "Region";
			this.cmbRegion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRegion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRegion.Enabled = false;
			this.cmbRegion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRegion.Location = new System.Drawing.Point(312, 38);
			this.cmbRegion.Name = "cmbRegion";
			this.cmbRegion.Size = new System.Drawing.Size(250, 21);
			this.cmbRegion.TabIndex = 318;
			this.cmbRegion.ValueMember = "idRegion";
			this.cmbRegion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRegion_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 104);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(70, 16);
			this.label8.TabIndex = 798;
			this.label8.Text = "Fecha Salida";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(8, 74);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(76, 16);
			this.label5.TabIndex = 796;
			this.label5.Text = "Fecha Ingreso";
			// 
			// dtFechaIngreso
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIngreso.Appearance = appearance2;
			this.dtFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaIngreso.DateButtons.Add(dateButton1);
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaIngreso.Format = "dd/MM/yyyy";
			this.dtFechaIngreso.Location = new System.Drawing.Point(96, 72);
			this.dtFechaIngreso.Name = "dtFechaIngreso";
			this.dtFechaIngreso.NonAutoSizeHeight = 23;
			this.dtFechaIngreso.Size = new System.Drawing.Size(136, 21);
			this.dtFechaIngreso.SpinButtonsVisible = true;
			this.dtFechaIngreso.TabIndex = 794;
			this.dtFechaIngreso.Value = ((object)(resources.GetObject("dtFechaIngreso.Value")));
			this.dtFechaIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaIngreso_KeyDown);
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 793;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 16);
			this.label1.TabIndex = 795;
			this.label1.Text = "Fecha";
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(24, 152);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 799;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(104, 152);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 802;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(192, 152);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 801;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(280, 152);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 800;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(368, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 803;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(256, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 1020;
			this.label3.Text = "Nombre";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(40, 16);
			this.label2.TabIndex = 1019;
			this.label2.Text = "Cédula";
			// 
			// txtNombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(312, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(272, 21);
			this.txtNombre.TabIndex = 1022;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			// 
			// txtCedula
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance5;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(96, 38);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(136, 21);
			this.txtCedula.TabIndex = 1023;
			this.txtCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCedula_KeyDown);
			this.txtCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtCedula_Validating);
			this.txtCedula.Validated += new System.EventHandler(this.txtCedula_Validated);
			// 
			// txtIdCobrador
			// 
			this.txtIdCobrador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCobrador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCobrador.Enabled = false;
			this.txtIdCobrador.Location = new System.Drawing.Point(568, 152);
			this.txtIdCobrador.Name = "txtIdCobrador";
			this.txtIdCobrador.PromptChar = ' ';
			this.txtIdCobrador.Size = new System.Drawing.Size(16, 21);
			this.txtIdCobrador.TabIndex = 1071;
			this.txtIdCobrador.Visible = false;
			// 
			// dtFechaSalida
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSalida.Appearance = appearance6;
			this.dtFechaSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaSalida.DateButtons.Add(dateButton3);
			this.dtFechaSalida.Enabled = false;
			this.dtFechaSalida.Format = "dd/MM/yyyy";
			this.dtFechaSalida.Location = new System.Drawing.Point(96, 102);
			this.dtFechaSalida.Name = "dtFechaSalida";
			this.dtFechaSalida.NonAutoSizeHeight = 23;
			this.dtFechaSalida.Size = new System.Drawing.Size(136, 21);
			this.dtFechaSalida.SpinButtonsVisible = true;
			this.dtFechaSalida.TabIndex = 1072;
			this.dtFechaSalida.Value = ((object)(resources.GetObject("dtFechaSalida.Value")));
			this.dtFechaSalida.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaSalida_KeyDown);
			// 
			// cmbCargo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCargo.Appearance = appearance7;
			this.cmbCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataSource = this.ultraDataSource1;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 245;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCargo.Enabled = false;
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(312, 72);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(250, 21);
			this.cmbCargo.TabIndex = 1073;
			this.cmbCargo.ValueMember = "idCargo";
			this.cmbCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCargo_KeyDown);
			this.cmbCargo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCargo_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cm
			// 
			this.cm.AutoSize = true;
			this.cm.Location = new System.Drawing.Point(256, 74);
			this.cm.Name = "cm";
			this.cm.Size = new System.Drawing.Size(35, 16);
			this.cm.TabIndex = 1074;
			this.cm.Text = "Cargo";
			this.cm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-152, 136);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(928, 8);
			this.groupBox2.TabIndex = 1075;
			this.groupBox2.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(256, 104);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 16);
			this.label4.TabIndex = 1076;
			this.label4.Text = "Código";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance8;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(312, 102);
			this.txtCodigo.MaxLength = 10;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(136, 21);
			this.txtCodigo.TabIndex = 1077;
			// 
			// btnAPP
			// 
			this.btnAPP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAPP.CausesValidation = false;
			this.btnAPP.Enabled = false;
			this.btnAPP.Image = ((System.Drawing.Image)(resources.GetObject("btnAPP.Image")));
			this.btnAPP.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAPP.Location = new System.Drawing.Point(448, 152);
			this.btnAPP.Name = "btnAPP";
			this.btnAPP.Size = new System.Drawing.Size(56, 23);
			this.btnAPP.TabIndex = 1078;
			this.btnAPP.Text = "&APP";
			this.btnAPP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAPP.Click += new System.EventHandler(this.btnAPP_Click);
			// 
			// frmIngresCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(600, 190);
			this.Controls.Add(this.btnAPP);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cm);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.dtFechaSalida);
			this.Controls.Add(this.txtIdCobrador);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.dtFechaIngreso);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbRegion);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmIngresCredito";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nomina CrediSolutions";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmIngresCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmIngresCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbRegion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCobrador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void lblCliente_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
//			if (btnEditar && btnNuevo)
//			{
//				int idIdentificacion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idIdentificacion From Nomina Where NIdentificacion = '{0}') Select idIdentificacion From Nomina Where NIdentificacion = '{0}' Else Select CONVERT(Int, 0)", this.txtIdentificacion.Text.ToString()));
//				if (idIdentificacion > 0)
//				{
//					//this.Consulta(idIdentificacion, true);
//					this.btnNuevo = false;
//					this.btnEditar = false;
//				}
//			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida identificación
			if (this.txtNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese los nombres completos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//this.utcNomina.Tabs[0].Selected = true;
				this.txtNombre.Focus();
				return;
			}

			if (this.txtCedula.Text.ToString().Length == 0)
			{
				MessageBox.Show("Esciba el numero de identificacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//this.utcNomina.Tabs[0].Selected = true;
				this.txtCedula.Focus();
				return;
			}

			#endregion Valida identificación

			#region Validacion			
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaIngreso, "Seleccione la Fecha De Ingreso", false, "", cdsSeteoF)) return;
			if ((int)this.txtIdCobrador.Value > 0)
			{
				if (this.dtFechaSalida.Value != System.DBNull.Value)
					if(!Validacion.vbComparaFechas((DateTime)this.dtFechaIngreso.Value,(DateTime)this.dtFechaSalida.Value,this.dtFechaIngreso,"la fecha de salida no puedes ser mayor a la fecha ingreso",true))return;
			}
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) return;
			#endregion Validacion

			if (!Validacion.vbComboVacio(this.cmbRegion, "Seleccione la Region")) return;
			if (!Validacion.vbComboVacio(this.cmbCargo, "Seleccione el Cargo")) return;

			#region Valida Cedula

			if (!Validacion.vbTexto(this.txtCodigo, 2, 30, "Codigo")) return;
			//			if ((int)this.txtIdCobrador.Value == 0)
			//			{
			//				int idCedulaRep = 0;
			//				idCedulaRep = Funcion.iEscalarSQL(cdsSeteoF, string.Format("exec ValidaCobrador '{0}'",
			//					this.txtCedula.Text.ToString()));
			//				
			//				if (idCedulaRep == 1)
			//				{
			//					MessageBox.Show("El Cobrador y/o Operador ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					this.txtCedula.Focus();
			//					return;
			//				}
			//			}

			#endregion Valida Cedula 

			DateTime dtFechaSalida = DateTime.Parse("01/01/2000");
			if (this.dtFechaSalida.Value != System.DBNull.Value)dtFechaSalida=(DateTime)this.dtFechaSalida.Value;
	
			if (DialogResult.Yes == MessageBox.Show("¿Desea guardar el registro?, Si presiona SI ya no podra modificar.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				string sSQLDet = string.Format("Exec GrabaIngresoCobrador {0},'{1}','{2}',{3},'{4}','{5}','{6}',{7},'{8}'",
					(int)this.txtIdCobrador.Value, this.txtCedula.Text, this.txtNombre.Text, (int) this.cmbRegion.Value, 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					Convert.ToDateTime(this.dtFechaIngreso.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(dtFechaSalida).ToString("yyyyMMdd"),
					(int)this.cmbCargo.Value,
					this.txtCodigo.Text);
				this.txtIdCobrador.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLDet);
				
				#region Controles 

				this.dtFecha.Enabled = false;
				this.txtCedula.Enabled = false;
				this.dtFechaIngreso.Enabled = false;
				this.dtFechaSalida.Enabled = false;
				this.txtNombre.Enabled = false;
				this.cmbRegion.Enabled = false;
				this.cmbCargo.Enabled = false;
				this.txtCodigo.Enabled = false;

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnAPP.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = false;

				if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
				if (miAcceso.Editar)this.btnEditar.Enabled = false;

				bNuevo = false;
				bEdicion = false;

				#endregion Controles 
			}
		}
			

		private void txtCedula_Validated(object sender, System.EventArgs e)
		{
		}

		private void txtCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}				
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			this.txtNombre.Enabled = true;
			this.txtCedula.Enabled = true;
			this.cmbRegion.Enabled = true;
			this.cmbCargo.Enabled = true;
			this.dtFechaSalida.Enabled = false;
			this.dtFechaIngreso.Enabled = true;
			this.txtCodigo.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnAPP.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			
			// variables de validacion inicializadas
			bNuevo = true;
			bEdicion = true;
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;
			
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAPP.Enabled = false;

			this.dtFecha.Enabled = false;
			this.txtCedula.Enabled = false;
			this.dtFechaIngreso.Enabled = false;
			this.dtFechaSalida.Enabled = false;
			this.txtNombre.Enabled = false;
			this.cmbRegion.Enabled = false;
			this.cmbCargo.Enabled = false;
			this.txtCodigo.Enabled = false;

			this.dtFecha.Value = System.DBNull.Value;
			this.txtCedula.Text = "";
			this.dtFechaIngreso.Value = System.DBNull.Value;
			this.dtFechaSalida.Value = System.DBNull.Value;
			this.txtNombre.Text = "";
			this.cmbRegion.Value = System.DBNull.Value;
			this.cmbCargo.Value = System.DBNull.Value;
			this.txtCodigo.Text = "";
			this.txtIdCobrador.Value = 0;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			// habilitar el formulario
			this.txtNombre.Enabled = true;
			this.txtCedula.Enabled = true;
			this.cmbRegion.Enabled = true;
			this.cmbCargo.Enabled = true;
			this.dtFechaSalida.Enabled = true;
			this.dtFechaIngreso.Enabled = true;
	
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnAPP.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			// this.chkInventario.CheckedChanged(sender, e);
			bNuevo = false;
			bEdicion = true;
			this.txtNombre.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaIngresoCredito Busqueda = new frmConsultaIngresoCredito())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idIngresoCobrador"].Value);
				}
			}
		}

		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec ConsultaIndividualIngresoCobrador {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdCobrador.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.txtCedula.Text = dr.GetString(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtNombre.Text = dr.GetString(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.cmbRegion.Value = dr.GetInt32(3); 
				if (dr.GetValue(4) != System.DBNull.Value) this.dtFecha.Value= dr.GetDateTime(4); 
				if (dr.GetValue(5) != System.DBNull.Value) this.dtFechaIngreso.Value = dr.GetDateTime(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.dtFechaSalida.Value = dr.GetDateTime(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.cmbCargo.Value = dr.GetInt32(7);
				if (dr.GetValue(8) != System.DBNull.Value) this.txtCodigo.Text = dr.GetString(8);
			}
			dr.Close();
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Autorizacion) this.btnAPP.Enabled = true;
		}

		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmIngresCredito_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "082917");

			if (!Funcion.Permiso("1444", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Solicitud De Gastos De Viaje", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.cmbRegion.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idRegion, Region from Region ");
			this.dtFechaIngreso.Value = DateTime.Now.AddDays(+1);
			this.dtFechaSalida.Value = DateTime.Now.AddDays(+2 );
			this.cmbCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"exec Listas 49, 0, 0");
		}


		private void ultraTextEditor2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCargo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCedula.Focus();
		}

		private void txtCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbRegion.Focus();
		}

		private void cmbRegion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaIngreso.Focus();
		}

		private void cmbCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaSalida.Focus();
		}

		private void dtFechaIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCargo.Focus();
		}

		private void dtFechaSalida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCargo.Focus();
		}

		private void frmIngresCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdCobrador.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtIdCobrador.Value);					
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnAPP_Click(object sender, System.EventArgs e)
		{
			int iCont = Funcion.iEscalarSQL(cdsSeteoF,string.Format(" SELECT COUNT(1) FROM DispositivosAPP WHERE idCom_Estado <> 2 AND CEDULA = '{0}' AND ACTIVO = 1 AND Empresa = 33",this.txtCedula.Text.ToString()));
			if( iCont == 0)
			{
				MessageBox.Show("El Usuario no activa la aplicación Movil en su Dispositivo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("¿Desea volver Activar al Usuario?, Validar que no sea un segundo dispositivo.\n La activación es bajo su responsabilidad.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				#region Abre observacion
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{	
						#region Guardar      
						using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
						{
							oConexion.Open();

							SqlTransaction oTransaction = null;
							SqlCommand oCmdActualiza = oConexion.CreateCommand();
							oCmdActualiza.Connection = oConexion;
							oCmdActualiza.CommandType = CommandType.Text;
							oCmdActualiza.CommandTimeout = 120;

							try
							{
								oTransaction = oConexion.BeginTransaction();
								oCmdActualiza.Transaction = oTransaction;

								string sSQLAutEs = string.Format("Update DispositivosAPP set KeyDispositivo= cast(newid() AS NVARCHAR(50)), iContActivacion = (iContActivacion +1) ,  ACTIVO = 0 where CEDULA = '{0}'",
									this.txtCedula.Text);
								oCmdActualiza.CommandText = sSQLAutEs;
								oCmdActualiza.ExecuteNonQuery();					
						

								string sSQLAutEsT = string.Format("iNSERT INTO DispositivosAPPDetalleActivacíon (Cedula, Notas) values ('{0}', '{1}')",
									this.txtCedula.Text, miNota.txtNotas.Text);
								oCmdActualiza.CommandText = sSQLAutEsT;
								oCmdActualiza.ExecuteNonQuery();
		
								oTransaction.Commit();
								Cursor = Cursors.Default;	

								MessageBox.Show("Usuario Activado en la Aplicación Movil.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
								//this.btnConsultar_Click(sender,e);
				
							}
							catch(Exception ex)
							{
								#region execpciones
								try
								{ 
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
								catch (Exception ex2)
								{
									MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
								}
								#endregion execpciones
							}
							finally
							{
								oConexion.Close();
							}
						}
				
						#endregion Guardar					
					}
				}
				#endregion Abre observacion

			}
		}
	}
}
