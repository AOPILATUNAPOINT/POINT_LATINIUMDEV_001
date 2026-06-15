using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCuentaSubPlan.
	/// </summary>
	public class frmCuentaSubPlan : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		bool bNuevo = false;
		bool bModifica = false;
		bool bEdicion = false;
		public static DateTime Desde = DateTime.Now;
		public static DateTime Hasta = DateTime.Now;
		int idCuentaAuxiliarPeriodo = 0;
		public static int idCuentaAuxiliarPeriodoGenerado = 0;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtFiltro;
		private System.Windows.Forms.TextBox txtDescripcion;
		private System.Windows.Forms.Label lblPadre;
		private System.Windows.Forms.Label lblDescripcion;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox grpActualizacion;
		private System.Windows.Forms.GroupBox grpBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPadre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCuentaAuxiliar;
		private System.Windows.Forms.TextBox txtTipoCuentaAuxiliar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoCuentaAuxiliar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipoCuentaAuxiliar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnGrabar;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnAnular;
		private Infragistics.Win.UltraWinGrid.UltraGrid udgEstadosFinancieros;
		private System.Windows.Forms.Button btnNota;
		private System.Windows.Forms.Button btnGuardarPeriodo;
		private System.Windows.Forms.Button button1;
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCuentaSubPlan()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAuxiliar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCuentaAuxiliar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuentaGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuentaSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Padre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoInicial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Haber");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaAuxiliarPeriodo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCuentaSubPlan));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID_CuentaEstadosABs");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuentaAuxiliar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ID_CuentaEstadosABs");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCuentaAuxiliar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoCuenta");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAuxiliar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoCuentaAuxiliar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Padre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoInicial");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debe");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Haber");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoFinal");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuentaAuxiliarPeriodo");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoFinal", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoFinal", 10, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtFiltro = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.grpActualizacion = new System.Windows.Forms.GroupBox();
			this.btnAnular = new System.Windows.Forms.Button();
			this.cmbTipoCuentaAuxiliar = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.udsTipoCuentaAuxiliar = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDescripcion = new System.Windows.Forms.TextBox();
			this.lblPadre = new System.Windows.Forms.Label();
			this.lblDescripcion = new System.Windows.Forms.Label();
			this.btnGrabar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.button1 = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.grpBusqueda = new System.Windows.Forms.GroupBox();
			this.txtTipoCuentaAuxiliar = new System.Windows.Forms.TextBox();
			this.txtidCuentaAuxiliar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPadre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.udgEstadosFinancieros = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNota = new System.Windows.Forms.Button();
			this.btnGuardarPeriodo = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.grpActualizacion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCuentaAuxiliar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoCuentaAuxiliar)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			this.grpBusqueda.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtidCuentaAuxiliar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPadre)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.udgEstadosFinancieros)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn2.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn7.DataType = typeof(System.UInt32);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.UInt32);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
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
			// txtFiltro
			// 
			this.txtFiltro.Location = new System.Drawing.Point(56, 24);
			this.txtFiltro.Name = "txtFiltro";
			this.txtFiltro.Size = new System.Drawing.Size(184, 20);
			this.txtFiltro.TabIndex = 751;
			this.txtFiltro.Text = "";
			this.txtFiltro.TextChanged += new System.EventHandler(this.txtFiltro_TextChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(48, 24);
			this.label1.TabIndex = 750;
			this.label1.Text = "Buscar:";
			// 
			// grpActualizacion
			// 
			this.grpActualizacion.Controls.Add(this.btnAnular);
			this.grpActualizacion.Controls.Add(this.cmbTipoCuentaAuxiliar);
			this.grpActualizacion.Controls.Add(this.txtDescripcion);
			this.grpActualizacion.Controls.Add(this.lblPadre);
			this.grpActualizacion.Controls.Add(this.lblDescripcion);
			this.grpActualizacion.Controls.Add(this.btnGrabar);
			this.grpActualizacion.Controls.Add(this.btnCancelar);
			this.grpActualizacion.Controls.Add(this.btnEditar);
			this.grpActualizacion.Location = new System.Drawing.Point(528, 16);
			this.grpActualizacion.Name = "grpActualizacion";
			this.grpActualizacion.Size = new System.Drawing.Size(384, 152);
			this.grpActualizacion.TabIndex = 753;
			this.grpActualizacion.TabStop = false;
			this.grpActualizacion.Text = "Actualización de datos";
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(280, 112);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 1037;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// cmbTipoCuentaAuxiliar
			// 
			this.cmbTipoCuentaAuxiliar.CausesValidation = false;
			this.cmbTipoCuentaAuxiliar.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoCuentaAuxiliar.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCuentaAuxiliar.DataSource = this.udsTipoCuentaAuxiliar;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Descripción";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 411;
			ultraGridColumn3.Header.Caption = "Tipo cuenta";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 115;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbTipoCuentaAuxiliar.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoCuentaAuxiliar.DisplayMember = "TipoCuentaAuxiliar";
			this.cmbTipoCuentaAuxiliar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoCuentaAuxiliar.Enabled = false;
			this.cmbTipoCuentaAuxiliar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCuentaAuxiliar.Location = new System.Drawing.Point(112, 72);
			this.cmbTipoCuentaAuxiliar.MaxDropDownItems = 30;
			this.cmbTipoCuentaAuxiliar.Name = "cmbTipoCuentaAuxiliar";
			this.cmbTipoCuentaAuxiliar.Size = new System.Drawing.Size(240, 22);
			this.cmbTipoCuentaAuxiliar.TabIndex = 761;
			this.cmbTipoCuentaAuxiliar.ValueMember = "ID_CuentaEstadosABs";
			// 
			// udsTipoCuentaAuxiliar
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			ultraDataColumn14.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			this.udsTipoCuentaAuxiliar.Band.Columns.AddRange(new object[] {
																																			ultraDataColumn13,
																																			ultraDataColumn14,
																																			ultraDataColumn15});
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(112, 36);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.ReadOnly = true;
			this.txtDescripcion.Size = new System.Drawing.Size(240, 20);
			this.txtDescripcion.TabIndex = 750;
			this.txtDescripcion.Text = "";
			// 
			// lblPadre
			// 
			this.lblPadre.Location = new System.Drawing.Point(8, 72);
			this.lblPadre.Name = "lblPadre";
			this.lblPadre.Size = new System.Drawing.Size(104, 24);
			this.lblPadre.TabIndex = 752;
			this.lblPadre.Text = "* Cuenta Auxiliar:";
			// 
			// lblDescripcion
			// 
			this.lblDescripcion.Location = new System.Drawing.Point(8, 24);
			this.lblDescripcion.Name = "lblDescripcion";
			this.lblDescripcion.TabIndex = 751;
			this.lblDescripcion.Text = "* Descripción:";
			// 
			// btnGrabar
			// 
			this.btnGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGrabar.Enabled = false;
			this.btnGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btnGrabar.Image")));
			this.btnGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGrabar.Location = new System.Drawing.Point(104, 112);
			this.btnGrabar.Name = "btnGrabar";
			this.btnGrabar.Size = new System.Drawing.Size(76, 23);
			this.btnGrabar.TabIndex = 749;
			this.btnGrabar.Text = "&Guardar";
			this.btnGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGrabar.Click += new System.EventHandler(this.btnGrabar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(192, 112);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 750;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(16, 112);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 1036;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.button1);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.dtHasta);
			this.groupBox2.Controls.Add(this.dtDesde);
			this.groupBox2.Controls.Add(this.btnGenerar);
			this.groupBox2.Location = new System.Drawing.Point(8, 16);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(512, 80);
			this.groupBox2.TabIndex = 754;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Período";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.CausesValidation = false;
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(416, 48);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(80, 23);
			this.button1.TabIndex = 1052;
			this.button1.Text = "&Anular";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click_2);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(200, 16);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 16);
			this.label17.TabIndex = 1051;
			this.label17.Text = "* Hasta:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(47, 16);
			this.label18.TabIndex = 1050;
			this.label18.Text = "* Desde:";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 16);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(136, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1049;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 16);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(128, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1048;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(416, 16);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(80, 23);
			this.btnGenerar.TabIndex = 752;
			this.btnGenerar.Text = "&Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.button1_Click_1);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(-760, 168);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(2608, 8);
			this.groupBox3.TabIndex = 1032;
			this.groupBox3.TabStop = false;
			// 
			// grpBusqueda
			// 
			this.grpBusqueda.Controls.Add(this.txtTipoCuentaAuxiliar);
			this.grpBusqueda.Controls.Add(this.label1);
			this.grpBusqueda.Controls.Add(this.txtFiltro);
			this.grpBusqueda.Controls.Add(this.txtidCuentaAuxiliar);
			this.grpBusqueda.Controls.Add(this.txtPadre);
			this.grpBusqueda.Controls.Add(this.btnBuscar);
			this.grpBusqueda.Location = new System.Drawing.Point(8, 104);
			this.grpBusqueda.Name = "grpBusqueda";
			this.grpBusqueda.Size = new System.Drawing.Size(512, 64);
			this.grpBusqueda.TabIndex = 761;
			this.grpBusqueda.TabStop = false;
			this.grpBusqueda.Text = "Búsqueda";
			// 
			// txtTipoCuentaAuxiliar
			// 
			this.txtTipoCuentaAuxiliar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.txtTipoCuentaAuxiliar.Enabled = false;
			this.txtTipoCuentaAuxiliar.Location = new System.Drawing.Point(260, 40);
			this.txtTipoCuentaAuxiliar.Name = "txtTipoCuentaAuxiliar";
			this.txtTipoCuentaAuxiliar.Size = new System.Drawing.Size(136, 20);
			this.txtTipoCuentaAuxiliar.TabIndex = 1036;
			this.txtTipoCuentaAuxiliar.Text = "";
			this.txtTipoCuentaAuxiliar.Visible = false;
			// 
			// txtidCuentaAuxiliar
			// 
			this.txtidCuentaAuxiliar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCuentaAuxiliar.Enabled = false;
			this.txtidCuentaAuxiliar.Location = new System.Drawing.Point(256, 16);
			this.txtidCuentaAuxiliar.MaskInput = "nn";
			this.txtidCuentaAuxiliar.MinValue = 0;
			this.txtidCuentaAuxiliar.Name = "txtidCuentaAuxiliar";
			this.txtidCuentaAuxiliar.PromptChar = ' ';
			this.txtidCuentaAuxiliar.Size = new System.Drawing.Size(136, 21);
			this.txtidCuentaAuxiliar.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtidCuentaAuxiliar.TabIndex = 1035;
			this.txtidCuentaAuxiliar.Visible = false;
			// 
			// txtPadre
			// 
			this.txtPadre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPadre.Enabled = false;
			this.txtPadre.Location = new System.Drawing.Point(256, 24);
			this.txtPadre.MaskInput = "nn";
			this.txtPadre.MinValue = 0;
			this.txtPadre.Name = "txtPadre";
			this.txtPadre.PromptChar = ' ';
			this.txtPadre.Size = new System.Drawing.Size(80, 21);
			this.txtPadre.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPadre.TabIndex = 1034;
			this.txtPadre.Visible = false;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(416, 24);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(78, 23);
			this.btnBuscar.TabIndex = 1036;
			this.btnBuscar.Text = "&Consultar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.udgEstadosFinancieros);
			this.groupBox1.Location = new System.Drawing.Point(24, 176);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(984, 328);
			this.groupBox1.TabIndex = 1035;
			this.groupBox1.TabStop = false;
			// 
			// udgEstadosFinancieros
			// 
			this.udgEstadosFinancieros.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udgEstadosFinancieros.Cursor = System.Windows.Forms.Cursors.Default;
			this.udgEstadosFinancieros.DataSource = this.ultraDataSource3;
			appearance3.BackColor = System.Drawing.Color.White;
			this.udgEstadosFinancieros.DisplayLayout.Appearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 83;
			ultraGridColumn6.Header.Caption = "Código";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 185;
			ultraGridColumn7.Header.Caption = "Descripción";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 429;
			ultraGridColumn8.Header.Caption = "Cuenta";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 101;
			ultraGridColumn9.Header.Caption = "Grupo";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 91;
			ultraGridColumn10.Header.Caption = "Subgrupo";
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 122;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 84;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Saldo inicial ($)";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Width = 107;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Debe ($)";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 67;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Haber ($)";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 67;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance7;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Saldo final ($)";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 101;
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			ultraGridBand2.Header.Caption = "Detalle";
			ultraGridBand2.HeaderVisible = true;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance8;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.udgEstadosFinancieros.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.udgEstadosFinancieros.DisplayLayout.GroupByBox.Hidden = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgEstadosFinancieros.DisplayLayout.Override.ActiveRowAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.udgEstadosFinancieros.DisplayLayout.Override.CardAreaAppearance = appearance10;
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.udgEstadosFinancieros.DisplayLayout.Override.GroupByRowAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8.5F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udgEstadosFinancieros.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadosFinancieros.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadosFinancieros.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udgEstadosFinancieros.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.udgEstadosFinancieros.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.udgEstadosFinancieros.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udgEstadosFinancieros.Location = new System.Drawing.Point(16, 24);
			this.udgEstadosFinancieros.Name = "udgEstadosFinancieros";
			this.udgEstadosFinancieros.Size = new System.Drawing.Size(944, 288);
			this.udgEstadosFinancieros.TabIndex = 748;
			this.udgEstadosFinancieros.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.udgEstadosFinancieros_InitializeLayout_2);
			this.udgEstadosFinancieros.Click += new System.EventHandler(this.udgEstadosFinancieros_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(928, 24);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 751;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNota
			// 
			this.btnNota.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNota.CausesValidation = false;
			this.btnNota.Image = ((System.Drawing.Image)(resources.GetObject("btnNota.Image")));
			this.btnNota.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNota.Location = new System.Drawing.Point(928, 64);
			this.btnNota.Name = "btnNota";
			this.btnNota.Size = new System.Drawing.Size(76, 23);
			this.btnNota.TabIndex = 1039;
			this.btnNota.Text = "&Notas";
			this.btnNota.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNota.Click += new System.EventHandler(this.button2_Click);
			// 
			// btnGuardarPeriodo
			// 
			this.btnGuardarPeriodo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardarPeriodo.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardarPeriodo.Image")));
			this.btnGuardarPeriodo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardarPeriodo.Location = new System.Drawing.Point(928, 104);
			this.btnGuardarPeriodo.Name = "btnGuardarPeriodo";
			this.btnGuardarPeriodo.Size = new System.Drawing.Size(76, 23);
			this.btnGuardarPeriodo.TabIndex = 1040;
			this.btnGuardarPeriodo.Text = "&Guardar";
			this.btnGuardarPeriodo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardarPeriodo.Visible = false;
			this.btnGuardarPeriodo.Click += new System.EventHandler(this.btnGuardarPeriodo_Click);
			// 
			// frmCuentaSubPlan
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1016, 558);
			this.Controls.Add(this.btnGuardarPeriodo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.grpActualizacion);
			this.Controls.Add(this.grpBusqueda);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNota);
			this.Name = "frmCuentaSubPlan";
			this.Text = "Estados financieros";
			this.Load += new System.EventHandler(this.frmCuentaSubPlan_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.grpActualizacion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCuentaAuxiliar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipoCuentaAuxiliar)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			this.grpBusqueda.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtidCuentaAuxiliar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPadre)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.udgEstadosFinancieros)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			bNuevo = true;
			bModifica = false;

			//his.btnDescripcion.Enabled = true;
		}

		private void udgEstadosFinancieros_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmCuentaSubPlan_Load(object sender, System.EventArgs e)
		{
			#region Seguridad
				string ambiente = "DEV"; //PRO
				if(ambiente == "PRO")
				{
					if (!Funcion.Permiso("2009", cdsSeteoF))
					{
						MessageBox.Show("No puede ingresar a Configuración Estados Financieros", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						BeginInvoke(new MethodInvoker(UnloadMe));
						return;
					}
				}
			#endregion Seguridad

			#region Fechas por defecto
				this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				this.dtHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
				DateTime dtFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
				this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2014");
				this.dtDesde.CalendarInfo.MaxDate = dtFecha;
				this.dtHasta.CalendarInfo.MaxDate = dtFecha;
				this.dtDesde.Value = "01/" + dtFecha.Month.ToString() + "/" + dtFecha.Year.ToString();
				this.dtHasta.Value = dtFecha.AddDays(-1);
			#endregion Fechas por defecto

			//ListaCuentas();
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void ListaCuentas()
		{

			idCuentaAuxiliarPeriodoGenerado = Funcion.iEscalarSQL(cdsSeteoF , String.Format("EXEC CuentaAuxiliarPeriodoGeneradoConXPeriodo '{0}','{1}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				), false);

			this.udgEstadosFinancieros.DataSource = FuncionesProcedimientos.dtGeneral(string.Format("EXEC ListaCuentaAuxiliar '{0}', {1}",
				this.txtFiltro.Text,
				idCuentaAuxiliarPeriodoGenerado
				)); 



			if(this.udgEstadosFinancieros.Rows.Count > 0)
			{
				//			this.udgEstadosFinancieros.DisplayLayout.ViewStyleBand =Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
				this.udgEstadosFinancieros.DisplayLayout.Bands[0].SortedColumns.Add("CuentaGrupo", false, true);
				this.udgEstadosFinancieros.DisplayLayout.Bands[0].SortedColumns.Add("CuentaSubGrupo", false, true);
				this.udgEstadosFinancieros.DisplayLayout.Bands[0].SortedColumns.Add("TipoCuentaAuxiliar", false,  true);
			}
			//Al filtar por código, se podrá editar
			EdicionConsulta();
		}

		private void udgEstadosFinancieros_Click(object sender, System.EventArgs e)
		{
			bool bdescripcion = true;
			bool bidcuentaauxiliar = true;
			bool btipocuentaauxiliar = true;
			bool bpadre = true;
			bool bidCuentaAuxiliarPeriodo = true;
			idCuentaAuxiliarPeriodo = 0;

			try 
			{
				this.txtDescripcion.Text = (string)this.udgEstadosFinancieros.ActiveRow.Cells["Descripcion"].Value.ToString();
			} 
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtDescripcion.Text = "";
				bdescripcion = false;
			}

			try 
			{
				this.txtTipoCuentaAuxiliar.Text = (string)this.udgEstadosFinancieros.ActiveRow.Cells["TipoCuentaAuxiliar"].Value.ToString();
			} 
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtTipoCuentaAuxiliar.Text = "";
				btipocuentaauxiliar = false;
			}
			
			try 
			{
				this.txtidCuentaAuxiliar.Value = (int)this.udgEstadosFinancieros.ActiveRow.Cells["idCuentaAuxiliar"].Value;
			} 
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtidCuentaAuxiliar.Value = 0;
				bidcuentaauxiliar = false;
			}

			try 
			{
				this.txtPadre.Value = (int)this.udgEstadosFinancieros.ActiveRow.Cells["Padre"].Value;
			} 
			catch(Exception ex)
			{
				string sms = ex.Message;
				this.txtPadre.Value = 0;
				bpadre = false;
			}

			try 
			{
				idCuentaAuxiliarPeriodo = (int)this.udgEstadosFinancieros.ActiveRow.Cells["idCuentaAuxiliarPeriodo"].Value;
			} 
			catch(Exception ex)
			{
				string sms = ex.Message;
				bidCuentaAuxiliarPeriodo = false;
			}

			if((int)this.txtidCuentaAuxiliar.Value != 0)
			{
				//Cargar código de la tabla cuenta auxiliar 
				string Codigo = "";
				Codigo = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarModel '{0}', {1}", "Reporte", (int)txtidCuentaAuxiliar.Value));
				this.cmbTipoCuentaAuxiliar.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC CuentaAuxiliarCuentasEstadosABCon '{0}'", Codigo));
				//Cargar combo por el tipo reporte
				this.cmbTipoCuentaAuxiliar.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC CuentaAuxiliarCuentasEstadosABCon '{0}'", Codigo));
				//Cargar el combo si ya está guardado anteriormente
				if(this.txtTipoCuentaAuxiliar.Text != "")
				{
					int ID_CuentaEstadosABs = 0;
					ID_CuentaEstadosABs = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarCuentaEstadosABValueCombo '{0}'", this.txtTipoCuentaAuxiliar.Text));
					if(ID_CuentaEstadosABs != 0)
					{
						this.cmbTipoCuentaAuxiliar.Value = ID_CuentaEstadosABs;
					}
				}
			}

			this.btnCancelar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnAnular.Enabled = true;

			if(bdescripcion == false && btipocuentaauxiliar == false && bidcuentaauxiliar == false && bpadre == false && bidCuentaAuxiliarPeriodo == false)
			{
				this.btnCancelar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.txtDescripcion.Text = "";
				this.txtidCuentaAuxiliar.Value = 0;
				this.txtTipoCuentaAuxiliar.Text = "";
				this.cmbTipoCuentaAuxiliar.Value = System.DBNull.Value;
				this.txtDescripcion.Enabled = false;
				this.txtTipoCuentaAuxiliar.Enabled = false;
				this.txtPadre.Value = 0;
				this.cmbTipoCuentaAuxiliar.Enabled = false;
				this.btnAnular.Enabled = false;
				idCuentaAuxiliarPeriodo = 0;
			}

		}



		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.txtDescripcion.Enabled = true;
			this.txtTipoCuentaAuxiliar.Enabled = true;
			this.btnGrabar.Enabled = true;
			this.btnEditar.Enabled = false;
			bEdicion = true;
			bModifica = true;
			this.cmbTipoCuentaAuxiliar.Enabled = true;
			this.btnAnular.Enabled = true;
		}

		private void btnGrabar_Click(object sender, System.EventArgs e)
		{

			if (this.txtDescripcion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el nombre de la cuenta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtDescripcion.Focus();
				return;
			}

//			if (this.txtTipoCuentaAuxiliar.Text.ToString().Length == 0)
//			{
//				MessageBox.Show("Ingrese la cuenta auxiliar que pertenece.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//				this.txtTipoCuentaAuxiliar.Focus();
//				return;
//			}

			if (bModifica)
			{
				string sSQL = string.Format("Select Count(*) From	CuentaAuxiliar Where Descripcion = '{0}' And idCuentaAuxiliar <> {1} And Padre = {2}", this.txtDescripcion.Text.ToString(), (int)this.txtidCuentaAuxiliar.Value
					, (int)this.txtPadre.Value);
				int iContUsuario = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iContUsuario > 0)
				{
					MessageBox.Show(string.Format("La descripción de cuenta auxiliar '{0}' ya existe.", this.txtDescripcion.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtDescripcion.Focus();
					return;
				}

				//Cargar TipoCuentaAuxiliar de la tabla cuenta auxiliar 
				string TipoCuentaAuxiliar = "";
				if(this.cmbTipoCuentaAuxiliar.Text != "" && this.cmbTipoCuentaAuxiliar.Text != "0" )
				{
					TipoCuentaAuxiliar = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarEstadosABModel '{0}', {1}", "TipoCuentaAuxiliar", (int)this.cmbTipoCuentaAuxiliar.Value));
				}

				int idCuentaAuxiliar = 0;
				try{idCuentaAuxiliar = (int)(int)this.cmbTipoCuentaAuxiliar.Value;}catch(Exception ex){string sms = ex.Message;}

				//Actualizar la cuenta auxiliar
				sSQL = string.Format("Exec ActualizarCuentaAuxiliar '{0}', '{1}', {2}, {3}, {4}", this.txtDescripcion.Text, TipoCuentaAuxiliar, (int)this.txtidCuentaAuxiliar.Value, idCuentaAuxiliarPeriodo, idCuentaAuxiliar);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
			}

			#region Mensaje de confirmación
			string sMensaje = "";
			if (bNuevo) sMensaje = "Cuenta registrada exitosamente.";
			if (bModifica) sMensaje = "Cuenta modificado exitosamente.";
			MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			#endregion Mensaje de confirmación

			ListaCuentas();

			this.txtDescripcion.Enabled = false;
			this.txtTipoCuentaAuxiliar.Enabled = false;
			this.cmbTipoCuentaAuxiliar.Enabled = false;
			this.txtidCuentaAuxiliar.Enabled = false;
						this.txtPadre.Enabled = false;
			this.btnGrabar.Enabled = false;
			this.btnEditar.Enabled = true;
			this.udgEstadosFinancieros.Enabled = true;

			bNuevo = false;
			bEdicion = false;
			bModifica = false;

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtDescripcion.Enabled = false;
			this.txtTipoCuentaAuxiliar.Enabled = false;
			this.cmbTipoCuentaAuxiliar.Enabled = false;
			this.txtidCuentaAuxiliar.Value = 0;
			this.cmbTipoCuentaAuxiliar.Value = System.DBNull.Value;
			this.txtPadre.Value = 0;
			this.txtDescripcion.Text = "";
			this.txtTipoCuentaAuxiliar.Text = "";
			this.btnEditar.Enabled = false;
			this.btnGrabar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.udgEstadosFinancieros.Enabled = true;
			this.btnAnular.Enabled = false;
			idCuentaAuxiliarPeriodo = 0;

			bNuevo = false;
			bModifica = false;
			bEdicion = false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void txtFiltro_TextChanged(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		public void EdicionConsulta()
		{
			//Desactivar los controle
			this.txtDescripcion.Enabled = false;
			this.txtTipoCuentaAuxiliar.Enabled = false;
			this.cmbTipoCuentaAuxiliar.Enabled = false;
			this.txtidCuentaAuxiliar.Value = 0;
			this.cmbTipoCuentaAuxiliar.Value = System.DBNull.Value;
			this.txtPadre.Value = 0;
			this.txtDescripcion.Text = "";
			this.txtTipoCuentaAuxiliar.Text = "";
			this.btnEditar.Enabled = false;
			this.btnGrabar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.udgEstadosFinancieros.Enabled = true;
			this.btnAnular.Enabled = false;


			bNuevo = false;
			bModifica = false;
			bEdicion = false;

			//Si solo existe una sola cuenta, podría activar la edición
			bool bEdicionConsulta = false;
			int idCuentaAuxiliar = 0;
			if (this.txtFiltro.Text != "")
			{
				idCuentaAuxiliar = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarIdConXCodigo '{0}'", this.txtFiltro.Text));
				bEdicionConsulta = true;
			}
			if(bEdicionConsulta)
			{
				if(idCuentaAuxiliar != 0)
				{
					bool bdescripcion = true;
					bool btipocuentaauxiliar = true;
					bool bpadre = true;
					bool bidcuentaauxiliar = true;
					bool bidCuentaAuxiliarPeriodo = true;

					this.txtidCuentaAuxiliar.Value = idCuentaAuxiliar;

					try 
					{
						this.txtDescripcion.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarModel '{0}', {1}", "Descripcion", idCuentaAuxiliar),false);
					} 
					catch(Exception ex)
					{
						string sms = ex.Message;
						this.txtDescripcion.Text = "";
						bdescripcion = false;
					}

					try 
					{
						this.txtTipoCuentaAuxiliar.Text = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarModel '{0}', {1}", "TipoCuentaAuxiliar", idCuentaAuxiliar), false);
					} 
					catch(Exception ex)
					{
						string sms = ex.Message;
						this.txtTipoCuentaAuxiliar.Text = "";
						btipocuentaauxiliar = false;
					}

					try 
					{
						this.txtPadre.Value = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarModel '{0}', {1}", "Padre", idCuentaAuxiliar), false);
					} 
					catch(Exception ex)
					{
						string sms = ex.Message;
						this.txtPadre.Value = 0;
						bpadre = false;
					}

					try 
					{
						idCuentaAuxiliarPeriodo = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarPeriodoIdXIdPeriodoGenCon {0}, {1}", idCuentaAuxiliar, idCuentaAuxiliarPeriodoGenerado));//(int)this.udgEstadosFinancieros.ActiveRow.Cells["idCuentaAuxiliarPeriodo"].Value;
					} 
					catch(Exception ex)
					{
						string sms = ex.Message;
						bidCuentaAuxiliarPeriodo = false;
					}

					if(idCuentaAuxiliar != 0)
					{
						//Cargar código de la tabla cuenta auxiliar 
						string Codigo = "";
						Codigo = Funcion.sEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarModel '{0}', {1}", "Reporte", (int)txtidCuentaAuxiliar.Value));
						this.cmbTipoCuentaAuxiliar.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC CuentaAuxiliarCuentasEstadosABCon '{0}'", Codigo));

						//Cargar combo por el tipo reporte
						this.cmbTipoCuentaAuxiliar.DataSource = Funcion.dvProcedimiento(cdsSeteoF, String.Format("EXEC CuentaAuxiliarCuentasEstadosABCon '{0}'", Codigo));
						//Cargar el combo si ya está guardado anteriormente

						if(this.txtTipoCuentaAuxiliar.Text != "")
						{
							int ID_CuentaEstadosABs = 0;
							ID_CuentaEstadosABs = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarCuentaEstadosABValueCombo '{0}'", this.txtTipoCuentaAuxiliar.Text));
							if(ID_CuentaEstadosABs != 0)
							{
								this.cmbTipoCuentaAuxiliar.Value = ID_CuentaEstadosABs;
							}
						}
					}

					this.btnCancelar.Enabled = true;
					this.btnEditar.Enabled = true;
					this.btnAnular.Enabled = true;

					if(bdescripcion == false && btipocuentaauxiliar == false && bidcuentaauxiliar == false && bpadre == false && bidCuentaAuxiliarPeriodo == false)
					{
						this.btnCancelar.Enabled = false;
						this.btnEditar.Enabled = false;
						this.txtDescripcion.Text = "";
						this.txtidCuentaAuxiliar.Value = 0;
						this.txtTipoCuentaAuxiliar.Text = "";
						this.cmbTipoCuentaAuxiliar.Value = System.DBNull.Value;
						this.txtDescripcion.Enabled = false;
						this.txtTipoCuentaAuxiliar.Enabled = false;
						this.txtPadre.Value = 0;
						this.cmbTipoCuentaAuxiliar.Enabled = false;
						this.btnAnular.Enabled = false;
					}
				}
			}
		}


		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			ListaCuentas();
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void button1_Click_1(object sender, System.EventArgs e)
		{
			//Procedemos a generar el perído
			#region Validación
		  if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;
			#endregion Validación
			//Validación si el período ya se encuentra generado.
			bool bRespuestaCount = Funcion.bEjecutaSQL(cdsSeteoF , String.Format("exec CuentaAuxiliarPeriodoGeneradoCount '{0}','{1}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				));

			if (!bRespuestaCount)
			{
				ListaCuentas();
				idCuentaAuxiliarPeriodoGenerado = Funcion.iEscalarSQL(cdsSeteoF , String.Format("EXEC CuentaAuxiliarPeriodoGeneradoConXPeriodo '{0}','{1}'",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
					), false);
				MessageBox.Show("El período "+ Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy") + ", ya se encuentra generado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			idCuentaAuxiliarPeriodoGenerado = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarPeriodoGeneradoCrear '{0}', '{1}', {2},{3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				, 1,
				idCuentaAuxiliarPeriodoGenerado),
				false
				);
//			if(idCuentaAuxiliarPeriodoGenerar != 0)
//			{
//				MessageBox.Show("Se guardó correctamente el período: " + Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				this.Close();
//			}

			bool bRespuesta = false;
			bRespuesta = Funcion.bEjecutaSQL(cdsSeteoF , String.Format("CuentaAuxiliarGenerarPeriodo {0}", 
				  idCuentaAuxiliarPeriodoGenerado
				));

			//Generar periodo resultado
			CuentaAuxiliarGenerarPeriodoResultado();

			//Generación
			//Se envia un bool para verificar si se muestran estos
			if(bRespuesta)
			{

				this.grpBusqueda.Enabled = true;
				this.grpActualizacion.Enabled = true;
				ListaCuentas();
				MessageBox.Show("Se generó correctamente el período: " + Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtDescripcion.Focus();

			}
			//Generar Notas
			CuentaAuxiliarNotasCenerar();


		}

		public void CuentaAuxiliarNotasCenerar()
		{
			bool bFunciones = Funcion.bEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarNotasCenerar {0}",
				frmCuentaSubPlan.idCuentaAuxiliarPeriodoGenerado),false
				);
		}

		public void CuentaAuxiliarGenerarPeriodoResultado()
		{
			bool bFunciones = Funcion.bEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarGenerarPeriodoResultado {0}",
				frmCuentaSubPlan.idCuentaAuxiliarPeriodoGenerado),false
				);
		}

		private void udgEstadosFinancieros_InitializeLayout_1(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void udgEstadosFinancieros_InitializeLayout_2(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			int idCuentaAuxiliar = 0;
			try{idCuentaAuxiliar = (int)this.txtidCuentaAuxiliar.Value;}
			catch(Exception ex){string sms = ex.Message;}
			if(idCuentaAuxiliar != 0)
			{
				//Actualizar la cuenta auxiliar
				string sSQL = string.Format("Exec ActualizarCuentaAuxiliar '{0}', '{1}', {2}, {3}, {4}", this.txtDescripcion.Text, "", (int)this.txtidCuentaAuxiliar.Value, idCuentaAuxiliarPeriodo, 0);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
			}

			this.txtDescripcion.Enabled = false;
			this.txtTipoCuentaAuxiliar.Enabled = false;
			this.cmbTipoCuentaAuxiliar.Enabled = false;
			this.txtidCuentaAuxiliar.Value = 0;
			this.cmbTipoCuentaAuxiliar.Value = System.DBNull.Value;
			this.txtPadre.Value = 0;
			this.txtDescripcion.Text = "";
			this.txtTipoCuentaAuxiliar.Text = "";
			this.btnEditar.Enabled = false;
			this.btnGrabar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.udgEstadosFinancieros.Enabled = true;
			this.btnAnular.Enabled = false;
			idCuentaAuxiliarPeriodo = 0;
			idCuentaAuxiliar = 0;

			bNuevo = false;
			bModifica = false;
			bEdicion = false;
			ListaCuentas();
			

		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			#region Validación
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;
			#endregion Validación
			Desde = Convert.ToDateTime(this.dtDesde.Value);
			Hasta = Convert.ToDateTime(this.dtHasta.Value);

			//Validar si existe un periodo generado´para poder ingresar
			bool bRespuestaCount = Funcion.bEjecutaSQL(cdsSeteoF , String.Format("exec CuentaAuxiliarPeriodoGeneradoCount '{0}','{1}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				));

			if (bRespuestaCount)
			{
				MessageBox.Show("El período "+ Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy") + ", aun no se encuentra generado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			else{
					idCuentaAuxiliarPeriodoGenerado = Funcion.iEscalarSQL(cdsSeteoF , String.Format("EXEC CuentaAuxiliarPeriodoGeneradoConXPeriodo '{0}','{1}'",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
					), false);
			}



			#region Añadir garantía
			using (frmCuentaAuxiliarNotas frmcan = new frmCuentaAuxiliarNotas())
			{		
				if (DialogResult.OK == frmcan.ShowDialog())
				{
				}
			}
			#endregion Añadir garantía
		}

		private void btnGuardarPeriodo_Click(object sender, System.EventArgs e)
		{
			//Validar si existe un periodo generado´para poder ingresar
			bool bRespuestaCount = Funcion.bEjecutaSQL(cdsSeteoF , String.Format("exec CuentaAuxiliarPeriodoGeneradoCount '{0}','{1}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				));

			if (bRespuestaCount)
			{
				MessageBox.Show("El período "+ Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy") + ", aun no se encuentra generado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

				idCuentaAuxiliarPeriodoGenerado = Funcion.iEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarPeriodoGeneradoCrear '{0}', '{1}', {2},{3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				, 1,
				idCuentaAuxiliarPeriodoGenerado),
				false
				);
				if(idCuentaAuxiliarPeriodoGenerado != 0)
				{
					MessageBox.Show("Se guardó correctamente el período: " + Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.Close();
				}
		}

		private void button1_Click_2(object sender, System.EventArgs e)
		{

			//Procedemos a generar el perído
			#region Validación
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;
			#endregion Validación

			idCuentaAuxiliarPeriodoGenerado = Funcion.iEscalarSQL(cdsSeteoF , String.Format("EXEC CuentaAuxiliarPeriodoGeneradoConXPeriodo '{0}','{1}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				), false);

			//Validar si existe un periodo generado´para poder ingresar
			bool bRespuestaCount = Funcion.bEjecutaSQL(cdsSeteoF , String.Format("exec CuentaAuxiliarPeriodoGeneradoCount '{0}','{1}'",
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")
				));

			if (bRespuestaCount)
			{
				MessageBox.Show("El período "+ Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy") + ", aun no se encuentra generado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			bool bRespuesta = Funcion.bEscalarSQL(cdsSeteoF, String.Format("EXEC CuentaAuxiliarPeriodoGeneradoEliminar {0}", 
			idCuentaAuxiliarPeriodoGenerado),
			false
			);
			if(bRespuesta)
			{
				ListaCuentas();
				MessageBox.Show("Se anuló correctamente el período: " + Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy") + " al " + Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
		}






	}
}
