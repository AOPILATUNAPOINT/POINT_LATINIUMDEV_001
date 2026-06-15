using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDepositosPorLiquidar.
	/// </summary>
	public class frmDepositosPorLiquidar : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Button btnIr;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDepositosPorLiquidar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.ContextMenu mnuIr;
		private System.Windows.Forms.MenuItem mnuIrAsiento;
		private System.Windows.Forms.MenuItem mnuIrNotaDeCredito;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDepositosPorLiquidar()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDepositosPorLiquidar));
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.btnIr = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdDepositosPorLiquidar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.mnuIr = new System.Windows.Forms.ContextMenu();
			this.mnuIrAsiento = new System.Windows.Forms.MenuItem();
			this.mnuIrNotaDeCredito = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDepositosPorLiquidar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCuenta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance1;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 360;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(70, 40);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(360, 21);
			this.cmbCuenta.TabIndex = 2;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown);
			this.cmbCuenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCuenta_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(6, 42);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 17);
			this.label5.TabIndex = 302;
			this.label5.Text = "Cuenta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(6, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(76, 17);
			this.label4.TabIndex = 301;
			this.label4.Text = "Observaciones";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(6, 106);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(43, 17);
			this.lblIce.TabIndex = 300;
			this.lblIce.Text = "Valor $.";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(6, 74);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 17);
			this.label1.TabIndex = 299;
			this.label1.Text = "Número";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(6, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 298;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance2;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(104, 144);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(528, 72);
			this.txtObservaciones.TabIndex = 5;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// txtValor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance3;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(70, 104);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 22);
			this.txtValor.TabIndex = 4;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// txtNumero
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance4;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(70, 72);
			this.txtNumero.MaxLength = 15;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 3;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// dtFecha
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance5;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.CausesValidation = false;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(70, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(232, 71);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 304;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(232, 105);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 19);
			this.lblUsuario.TabIndex = 305;
			// 
			// btnIr
			// 
			this.btnIr.CausesValidation = false;
			this.btnIr.Enabled = false;
			this.btnIr.Image = ((System.Drawing.Image)(resources.GetObject("btnIr.Image")));
			this.btnIr.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIr.Location = new System.Drawing.Point(400, 240);
			this.btnIr.Name = "btnIr";
			this.btnIr.Size = new System.Drawing.Size(80, 23);
			this.btnIr.TabIndex = 10;
			this.btnIr.Text = "Ir";
			this.btnIr.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIr.Click += new System.EventHandler(this.btnIr_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(6, 240);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(81, 240);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 7;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(162, 240);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 8;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(241, 240);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 9;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(321, 240);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 6;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(483, 240);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 11;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(562, 240);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 12;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(568, 8);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 317;
			this.txtIdAsiento.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(584, 8);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 316;
			this.txtEstado.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(600, 8);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 315;
			this.txtIdCompra.Visible = false;
			// 
			// txtIdDepositosPorLiquidar
			// 
			this.txtIdDepositosPorLiquidar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDepositosPorLiquidar.Enabled = false;
			this.txtIdDepositosPorLiquidar.Location = new System.Drawing.Point(616, 8);
			this.txtIdDepositosPorLiquidar.Name = "txtIdDepositosPorLiquidar";
			this.txtIdDepositosPorLiquidar.PromptChar = ' ';
			this.txtIdDepositosPorLiquidar.Size = new System.Drawing.Size(16, 22);
			this.txtIdDepositosPorLiquidar.TabIndex = 314;
			this.txtIdDepositosPorLiquidar.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 224);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 8);
			this.groupBox1.TabIndex = 318;
			this.groupBox1.TabStop = false;
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
			// mnuIr
			// 
			this.mnuIr.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					this.mnuIrAsiento,
																																					this.mnuIrNotaDeCredito});
			// 
			// mnuIrAsiento
			// 
			this.mnuIrAsiento.Index = 0;
			this.mnuIrAsiento.Text = "Asiento";
			this.mnuIrAsiento.Click += new System.EventHandler(this.mnuIrAsiento_Click);
			// 
			// mnuIrNotaDeCredito
			// 
			this.mnuIrNotaDeCredito.Enabled = false;
			this.mnuIrNotaDeCredito.Index = 1;
			this.mnuIrNotaDeCredito.Text = "Nota de Credito";
			this.mnuIrNotaDeCredito.Click += new System.EventHandler(this.mnuIrNotaDeCredito_Click);
			// 
			// frmDepositosPorLiquidar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(640, 272);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdDepositosPorLiquidar);
			this.Controls.Add(this.btnIr);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDepositosPorLiquidar";
			this.Text = "Depositos Por Liquidar";
			this.Load += new System.EventHandler(this.frmDepositosPorLiquidar_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDepositosPorLiquidar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		int idBloqueaTransacciones = 0;
		#endregion Variables

		private void frmDepositosPorLiquidar_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "040808");

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;		

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Descripcion From CuentaCorriente cc Inner Join Cuenta c On c.idCuenta = cc.idCuenta Order By c.Descripcion");
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 4", idBloqueaTransacciones, (int)this.txtIdDepositosPorLiquidar.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			this.txtIdDepositosPorLiquidar.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtIdCompra.Value = 0;

			this.dtFecha.Value = System.DBNull.Value;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtValor.Value = 0.00m;
			this.txtObservaciones.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.dtFecha.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtObservaciones.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnIr.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;			

			bNuevo = false;
			bEdicion = false;
			idBloqueaTransacciones = 0;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una Fecha", true, "Bancos", cdsSeteoF)) return;

			if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion ((int)this.cmbCuenta.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;

			if (!FuncionesContabilidad.vbValidaDocumentoBancario((int)this.cmbCuenta.Value, this.txtNumero, 2, (int)this.txtIdAsiento.Value, cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione una cuenta")) return;

			if (!Validacion.vbCampoDecimalVacio(this.txtValor, "El Valor del deposito no puede ser 0.00", 0, 2100)) return;

			if (!Validacion.vbTexto(this.txtObservaciones, 10, 300, "Observaciones")) return;

			string sSQL = string.Format("Exec GuardaDepositosPorLiquidar {0}, '{1}', {2}, '{3}', {4}, '{5}'", 
				(int)this.txtIdDepositosPorLiquidar.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbCuenta.Value,
				this.txtNumero.Text.ToString(), Convert.ToDecimal(this.txtValor.Value), this.txtObservaciones.Text.ToString());
			this.txtIdDepositosPorLiquidar.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			string sSQLAsiento = string.Format("Exec AsientoDepositosPorLiquidar {0}", (int)this.txtIdDepositosPorLiquidar.Value);
			this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
			
			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 4", idBloqueaTransacciones, (int)this.txtIdDepositosPorLiquidar.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);

			this.dtFecha.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtValor.Enabled = false;
			this.txtObservaciones.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			this.btnGuardar.Enabled = false;
			if (miAcceso.CrearAsiento) this.btnIr.Enabled = true;
			if (miAcceso.Imprimir) this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = false;
			idBloqueaTransacciones = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			this.cmbCuenta.Enabled = true;
			this.txtNumero.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtObservaciones.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnIr.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;			

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			this.dtFecha.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void Consulta (int idDepositoPorLiquidar)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnIr.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.mnuIrNotaDeCredito.Enabled = true;

			string sSQL = string.Format("Exec ConsultaDepositosPorLiquidar {0}", idDepositoPorLiquidar);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.txtIdDepositosPorLiquidar.Value = dr.GetInt32(0);
				this.dtFecha.Value = dr.GetDateTime(1);
				this.cmbCuenta.Value = dr.GetInt32(2);
				this.txtNumero.Text = dr.GetString(3);
				this.txtValor.Value = dr.GetDecimal(4);
				this.txtObservaciones.Text = dr.GetString(5);
				this.txtIdAsiento.Value = dr.GetInt32(6);
				this.txtEstado.Value = dr.GetInt32(7);
				this.lblEstado.Text = dr.GetString(8);
				this.txtIdCompra.Value = dr.GetInt32(9);
				this.lblUsuario.Text = dr.GetString(10);
			}
			dr.Close();	

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			if ((int)this.txtEstado.Value != 3)
			{
				if (miAcceso.Editar) if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				this.btnGuardar.Enabled = false;
				if (miAcceso.CrearAsiento) this.btnIr.Enabled = true;
				if (miAcceso.Anular) this.btnAnular.Enabled = true;
				this.btnCancelar.Enabled = true;

				if ((int)this.txtIdCompra.Value > 0) this.mnuIrNotaDeCredito.Enabled = true;
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaDepositosPorLiquidar Busqueda = new frmConsultaDepositosPorLiquidar())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idDepositosPorLiquidar"].Value);
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdDepositosPorLiquidar.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta cruzado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
						
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una Fecha", true, "Bancos", cdsSeteoF)) return;
			
			if (!FuncionesContabilidad.ValidaAsientoEnConciliacion((int)this.txtIdAsiento.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;

			if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion ((int)this.cmbCuenta.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;

			if (Funcion.ValidaBloqueo((int)this.txtIdDepositosPorLiquidar.Value, 4, cdsSeteoF)) return;
	
			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 4", (int)this.txtIdDepositosPorLiquidar.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			this.cmbCuenta.Enabled = true;
			if (miAcceso.CambiarNumeracion) this.txtNumero.Enabled = true;
			this.txtValor.Enabled = true;
			this.txtObservaciones.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnIr.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;			

			this.dtFecha.Focus();

			bNuevo = false;
			bEdicion = true;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdDepositosPorLiquidar.Value);
			
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			string strFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.Value.ToString();			
			Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
			Reporte.MdiParent = MdiParent;
			Reporte.Show();
		}

		private void btnIr_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			this.mnuIr.Show(this.btnIr, new Point(1, 25));
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdDepositosPorLiquidar.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}			
			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "", true, "Bancos", cdsSeteoF)) return;
			
			if (!FuncionesContabilidad.ValidaAsientoEnConciliacion((int)this.txtIdAsiento.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;

			if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion ((int)this.cmbCuenta.Value, Convert.ToDateTime(this.dtFecha.Value), cdsSeteoF)) return;

			if ((int)this.txtEstado.Value == 2 && !miAcceso.BAnularPagosRealizados)
			{
				MessageBox.Show("El documento esta cruzado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
//			if (!Validacion.vbValidaAnulacionCobrosCuotasPosteriores((int)this.txtIdCompra.Value, 1, cdsSeteoF)) return;

			if (Funcion.ValidaBloqueo((int)this.txtIdDepositosPorLiquidar.Value, 4, cdsSeteoF)) return;

			using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
			{
				if (DialogResult.OK == Nota.ShowDialog())
				{
					string sSQL = string.Format("Exec AnulaDepositosPorLiquidar {0}, '{1}'", (int)this.txtIdDepositosPorLiquidar.Value, Nota.txtNotas.Text.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					this.Consulta((int)this.txtIdDepositosPorLiquidar.Value);
				}
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.txtNumero.Enabled) this.txtNumero.Focus(); else this.txtValor.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void mnuIrAsiento_Click(object sender, System.EventArgs e)
		{
			frmAsientos Asiento = new frmAsientos((int)this.txtIdAsiento.Value);
			Asiento.ShowDialog();
		}

		private void mnuIrNotaDeCredito_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("No existe nota de credito de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			frmDevolucionVenta NDC = new frmDevolucionVenta((int)this.txtIdCompra.Value);
			NDC.ShowDialog();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void cmbCuenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
