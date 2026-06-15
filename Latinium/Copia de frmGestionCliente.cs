using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGestionCliente.
	/// </summary>
	public class frmGestionCliente : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsSeguimiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSeguimiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		C1.Data.C1DataRow drSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbContactos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConvencional;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNCotizacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConvencional;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCelular;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEmail;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		C1.Data.C1DataRow drSeguimiento;
		int IdCliente = 0;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		int iEstado = 0;
		bool bEdicion = false;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		public frmGestionCliente()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteContacto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClienteContacto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contacto");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbContactos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeguimiento = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkConvencional = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNCotizacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.chkCelular = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkEmail = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtConvencional = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdSeguimiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbContactos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNCotizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConvencional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSeguimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.SuspendLayout();
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(8, 152);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(96, 16);
			this.lblVendedor.TabIndex = 310;
			this.lblVendedor.Text = "Lista de Contactos";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 40);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(38, 16);
			this.lblCliente.TabIndex = 304;
			this.lblCliente.Text = "Cliente";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(8, 88);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(90, 16);
			this.label19.TabIndex = 295;
			this.label19.Text = "Fecha de Gestión";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbContactos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbContactos.Appearance = appearance1;
			this.cmbContactos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbContactos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbContactos.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeguimiento, "SeguimientoCotizacion.idClienteContacto"));
			this.cmbContactos.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 360;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbContactos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbContactos.DisplayLayout.Override.NullText = "Seleccione un contacto";
			this.cmbContactos.DisplayMember = "Contacto";
			this.cmbContactos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbContactos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbContactos.Location = new System.Drawing.Point(144, 152);
			this.cmbContactos.Name = "cmbContactos";
			this.cmbContactos.Size = new System.Drawing.Size(360, 21);
			this.cmbContactos.TabIndex = 309;
			this.cmbContactos.ValueMember = "idClienteContacto";
			this.cmbContactos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbContactos_KeyPress);
			this.cmbContactos.Validating += new System.ComponentModel.CancelEventHandler(this.cmbContactos_Validating);
			this.cmbContactos.Validated += new System.EventHandler(this.cmbContactos_Validated);
			this.cmbContactos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbContactos_InitializeLayout);
			// 
			// cdsSeguimiento
			// 
			this.cdsSeguimiento.BindingContextCtrl = this;
			this.cdsSeguimiento.DataLibrary = "LibFacturacion";
			this.cdsSeguimiento.DataLibraryUrl = "";
			this.cdsSeguimiento.DataSetDef = "dsSeguimientoCotizacion";
			this.cdsSeguimiento.EnforceConstraints = false;
			this.cdsSeguimiento.FillOnRequest = false;
			this.cdsSeguimiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguimiento.Name = "cdsSeguimiento";
			this.cdsSeguimiento.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeguimiento.SchemaDef = null;
			this.cdsSeguimiento.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguimiento_BeforeFill);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// txtCliente
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCliente.Appearance = appearance2;
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Enabled = false;
			this.txtCliente.Location = new System.Drawing.Point(144, 37);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(360, 22);
			this.txtCliente.TabIndex = 299;
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeguimiento, "SeguimientoCotizacion.Fecha"));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy hh:mm";
			this.dtFecha.Location = new System.Drawing.Point(144, 88);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(144, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 301;
			// 
			// chkConvencional
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConvencional.Appearance = appearance4;
			this.chkConvencional.BackColor = System.Drawing.Color.Transparent;
			this.chkConvencional.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeguimiento, "SeguimientoCotizacion.Telefono"));
			this.chkConvencional.Location = new System.Drawing.Point(8, 184);
			this.chkConvencional.Name = "chkConvencional";
			this.chkConvencional.Size = new System.Drawing.Size(120, 22);
			this.chkConvencional.TabIndex = 297;
			this.chkConvencional.Text = "Convencional";
			this.chkConvencional.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.chkConvencional_KeyPress);
			this.chkConvencional.BeforeCheckStateChanged += new Infragistics.Win.CheckEditor.BeforeCheckStateChangedHandler(this.chkConvencional_BeforeCheckStateChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(123, 16);
			this.label2.TabIndex = 322;
			this.label2.Text = "Ingrese el N. Cotización";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNCotizacion
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNCotizacion.Appearance = appearance5;
			this.txtNCotizacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNCotizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNCotizacion.Location = new System.Drawing.Point(144, 5);
			this.txtNCotizacion.Name = "txtNCotizacion";
			this.txtNCotizacion.Size = new System.Drawing.Size(144, 22);
			this.txtNCotizacion.TabIndex = 321;
			this.txtNCotizacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNCotizacion_KeyPress);
			this.txtNCotizacion.Validating += new System.ComponentModel.CancelEventHandler(this.txtNCotizacion_Validating);
			this.txtNCotizacion.Validated += new System.EventHandler(this.txtNCotizacion_Validated);
			this.txtNCotizacion.ValueChanged += new System.EventHandler(this.txtNCotizacion_ValueChanged);
			// 
			// txtObservaciones
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance6;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeguimiento, "SeguimientoCotizacion.Observaciones"));
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(8, 312);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(496, 152);
			this.txtObservaciones.TabIndex = 323;
			this.txtObservaciones.Leave += new System.EventHandler(this.txtObservaciones_Leave);
			this.txtObservaciones.Validating += new System.ComponentModel.CancelEventHandler(this.txtObservaciones_Validating);
			this.txtObservaciones.Validated += new System.EventHandler(this.txtObservaciones_Validated);
			this.txtObservaciones.ValueChanged += new System.EventHandler(this.txtObservaciones_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 288);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(76, 16);
			this.label3.TabIndex = 324;
			this.label3.Text = "Observaciones";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkCelular
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCelular.Appearance = appearance7;
			this.chkCelular.BackColor = System.Drawing.Color.Transparent;
			this.chkCelular.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeguimiento, "SeguimientoCotizacion.Celular"));
			this.chkCelular.Location = new System.Drawing.Point(8, 216);
			this.chkCelular.Name = "chkCelular";
			this.chkCelular.Size = new System.Drawing.Size(120, 22);
			this.chkCelular.TabIndex = 328;
			this.chkCelular.Text = "Celular";
			this.chkCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.chkCelular_KeyPress);
			this.chkCelular.BeforeCheckStateChanged += new Infragistics.Win.CheckEditor.BeforeCheckStateChangedHandler(this.chkCelular_BeforeCheckStateChanged);
			this.chkCelular.CheckedChanged += new System.EventHandler(this.chkCelular_CheckedChanged);
			// 
			// chkEmail
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEmail.Appearance = appearance8;
			this.chkEmail.BackColor = System.Drawing.Color.Transparent;
			this.chkEmail.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsSeguimiento, "SeguimientoCotizacion.EMail"));
			this.chkEmail.Location = new System.Drawing.Point(8, 248);
			this.chkEmail.Name = "chkEmail";
			this.chkEmail.Size = new System.Drawing.Size(120, 22);
			this.chkEmail.TabIndex = 329;
			this.chkEmail.Text = "E-Mail";
			this.chkEmail.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.chkEmail_KeyPress);
			this.chkEmail.BeforeCheckStateChanged += new Infragistics.Win.CheckEditor.BeforeCheckStateChangedHandler(this.chkEmail_BeforeCheckStateChanged);
			// 
			// txtConvencional
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConvencional.Appearance = appearance9;
			this.txtConvencional.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConvencional.Enabled = false;
			this.txtConvencional.Location = new System.Drawing.Point(144, 184);
			this.txtConvencional.Name = "txtConvencional";
			this.txtConvencional.Size = new System.Drawing.Size(136, 22);
			this.txtConvencional.TabIndex = 330;
			// 
			// txtCelular
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCelular.Appearance = appearance10;
			this.txtCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelular.Enabled = false;
			this.txtCelular.Location = new System.Drawing.Point(144, 216);
			this.txtCelular.Name = "txtCelular";
			this.txtCelular.Size = new System.Drawing.Size(136, 22);
			this.txtCelular.TabIndex = 331;
			// 
			// txtEMail
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEMail.Appearance = appearance11;
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Enabled = false;
			this.txtEMail.Location = new System.Drawing.Point(144, 248);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(352, 22);
			this.txtEMail.TabIndex = 332;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(304, 3);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 333;
			// 
			// btCliente
			// 
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.CausesValidation = false;
			this.btCliente.Location = new System.Drawing.Point(112, 152);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 22);
			this.btCliente.TabIndex = 334;
			this.btCliente.Text = "...";
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.CausesValidation = false;
			this.barraDato1.DataMember = "SeguimientoCotizacion";
			this.barraDato1.DataNombreId = "idSeguimiento";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsSeguimiento;
			this.barraDato1.DataTabla = "SeguimientoCotizacion";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 472);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(256, 26);
			this.barraDato1.TabIndex = 335;
			this.barraDato1.VisibleBorrar = false;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
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
			// txtIdSeguimiento
			// 
			this.txtIdSeguimiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdSeguimiento.Appearance = appearance12;
			this.txtIdSeguimiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeguimiento, "SeguimientoCotizacion.idSeguimiento"));
			this.txtIdSeguimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSeguimiento.Enabled = false;
			this.txtIdSeguimiento.FormatString = "";
			this.txtIdSeguimiento.Location = new System.Drawing.Point(480, 472);
			this.txtIdSeguimiento.Name = "txtIdSeguimiento";
			this.txtIdSeguimiento.PromptChar = ' ';
			this.txtIdSeguimiento.Size = new System.Drawing.Size(24, 22);
			this.txtIdSeguimiento.TabIndex = 336;
			this.txtIdSeguimiento.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdCompra.Appearance = appearance13;
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeguimiento, "SeguimientoCotizacion.idCompra"));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.FormatString = "";
			this.txtIdCompra.Location = new System.Drawing.Point(448, 472);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(24, 22);
			this.txtIdCompra.TabIndex = 338;
			this.txtIdCompra.Visible = false;
			// 
			// txtNumero
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance14;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSeguimiento, "SeguimientoCotizacion.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(144, 120);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(144, 22);
			this.txtNumero.TabIndex = 339;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 120);
			this.label1.Name = "label1";
			this.label1.TabIndex = 340;
			this.label1.Text = "Número de Gestión";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(520, 8);
			this.groupBox1.TabIndex = 341;
			this.groupBox1.TabStop = false;
			// 
			// frmGestionCliente
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(512, 502);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdSeguimiento);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEMail);
			this.Controls.Add(this.txtCelular);
			this.Controls.Add(this.txtConvencional);
			this.Controls.Add(this.chkEmail);
			this.Controls.Add(this.chkCelular);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtNCotizacion);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.cmbContactos);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.chkConvencional);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmGestionCliente";
			this.Text = "Gestión Cliente";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGestionCliente_KeyDown);
			this.Load += new System.EventHandler(this.frmGestionCliente_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbContactos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNCotizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConvencional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSeguimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmGestionCliente_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{			
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Deshacer(sender, e);
				}
				else
					Close();
			}
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.B)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F12)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F5)
			{
				if (!this.barraDato1.bEditar)
				{
					barraDato1.IdRegistro = (int)drSeguimiento["idSeguimiento"];
					barraDato1.ProximoId(5);					
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cdsSeguimiento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeguimiento.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}

		private void frmGestionCliente_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0625");

			if (!Funcion.Permiso("990", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso GESTIÓN CLIENTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drSeguimiento = cdsSeguimiento.TableViews["SeguimientoCotizacion"].Rows[0];

				string sSQL = string.Format("Select Numero, idCliente, Nombre, Estado From Compra Where idTipoFactura = 3 And idCompra = '{0}'", (int)drSeguimiento["idCompra"]);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{	
					this.txtNCotizacion.Value = dr.GetValue(0).ToString();
					IdCliente = dr.GetInt32(1);
					this.txtCliente.Value = dr.GetString(2);
					iEstado = dr.GetInt32(3);
				}
				dr.Close();

				this.cmbContactos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idClienteContacto, Contacto From ClienteContacto Where idCliente = {0}", IdCliente));

				sSQL = string.Format("Select Telefono, Celular, Email From ClienteContacto Where idClienteContacto = {0}", (int)drSeguimiento["idClienteContacto"]);
				SqlDataReader dr1 = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr1.Read();
				if (dr1.HasRows)
				{
					this.txtConvencional.Text = dr1.GetValue(0).ToString();
					this.txtCelular.Text = dr1.GetValue(1).ToString();
					this.txtEMail.Text = dr1.GetValue(2).ToString();
				}
				dr1.Close();
			}
			catch{};
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();

			this.txtNCotizacion.Text = "";
			this.txtCliente.Text = "";
			this.txtConvencional.Text = "";
			this.txtCelular.Text = "";
			this.txtEMail.Text = "";

			this.txtCliente.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtConvencional.Enabled = false;
			this.txtCelular.Enabled = false;
			this.txtEMail.Enabled = false;

			drSeguimiento["Fecha"] = DateTime.Now;

			bEdicion = true;

			this.txtNCotizacion.Select();
		}

		private void txtNCotizacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.cmbContactos.Select();
		}

		private void txtNCotizacion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!this.barraDato1.bEditar) return;
			
			if (this.txtNCotizacion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de cotización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				e.Cancel = true;
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Compra Where idTipoFactura = 3 And Numero = '{0}'", this.txtNCotizacion.Text.ToString()), true) == 0)
			{
				MessageBox.Show("El número de cotización " + this.txtNCotizacion.Text.ToString() + " no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				e.Cancel = true;
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Compra Where idTipoFactura = 3 And Borrar = 1 And Numero = '{0}'", this.txtNCotizacion.Text.ToString()), true) > 0)
			{
				MessageBox.Show("El número de cotización " + this.txtNCotizacion.Text.ToString() + " esta anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				e.Cancel = true;
				return;
			}
		}

		private void txtNCotizacion_Validated(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Select idCompra, idCliente, Nombre, Estado From Compra Where idTipoFactura = 3 And Borrar = 0 And Numero = '{0}'", this.txtNCotizacion.Text.ToString());
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{	
				drSeguimiento["idCompra"] = dr.GetInt32(0);
				IdCliente = dr.GetInt32(1);
				this.txtCliente.Value = dr.GetString(2);
				iEstado = dr.GetInt32(3);
			}
			dr.Close();

			this.cmbContactos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idClienteContacto, Contacto From ClienteContacto Where idCliente = {0}", IdCliente));
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			bEdicion = false;

			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();

			this.txtNCotizacion.Select();

			bEdicion = true;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			bEdicion = false;

			#region Validacion
			if (this.txtNCotizacion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de cotización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Compra Where idTipoFactura = 3 And Numero = '{0}'", this.txtNCotizacion.Text.ToString()), true) == 0)
			{
				MessageBox.Show("El número de cotización " + this.txtNCotizacion.Text.ToString() + " no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Compra Where idTipoFactura = 3 And Borrar = 1 And Numero = '{0}'", this.txtNCotizacion.Text.ToString()), true) > 0)
			{
				MessageBox.Show("El número de cotización " + this.txtNCotizacion.Text.ToString() + " esta anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				return;
			}

			if (this.cmbContactos.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un contacto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbContactos.Select();
				return;
			}

			if (!this.chkConvencional.Checked && !this.chkCelular.Checked && !this.chkEmail.Checked)
			{
				MessageBox.Show("Seleccione un tipo de contacto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.chkConvencional.Select();
				return;
			}

			if (this.txtObservaciones.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una observación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtObservaciones.Select();
				return;
			}
			#endregion Validacion

			#region Numero
			int iCont = Funcion.iEscalarSQL(cdsSeguimiento, "Select COUNT(*) From SeguimientoCotizacion") + 1;
			drSeguimiento["Numero"] = Funcion.sEscalarSQL(cdsSeguimiento, string.Format("Select ISNULL(REPLICATE('0', 10 - LEN(ISNULL(CAST({0} As NVarchar(40)),0))), '') + CAST ({0} As NVarchar(40))", iCont));
			#endregion Numero

			this.barraDato1.GrabaRegistro();

			this.barraDato1.IdRegistro = (int)drSeguimiento["idSeguimiento"];
			barraDato1.ProximoId(5);
		}

		private void cmbContactos_Validated(object sender, System.EventArgs e)
		{
			if (this.cmbContactos.Text.ToString().Length == 0) return;

			this.chkConvencional.Checked = false;
			this.chkCelular.Checked = false;
			this.chkEmail.Checked = false;

			string sSQL = string.Format("Select Telefono, Celular, Email From ClienteContacto Where idClienteContacto = {0}", (int)this.cmbContactos.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtConvencional.Text = dr.GetValue(0).ToString();
				this.txtCelular.Text = dr.GetValue(1).ToString();
				this.txtEMail.Text = dr.GetValue(2).ToString();
			}
			dr.Close();
		}

		private void cmbContactos_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (this.cmbContactos.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un contacto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbContactos.Select();
				e.Cancel = true;
				return;
			}
		}

		private void cmbContactos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.chkConvencional.Select();
		}

		private void chkConvencional_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.chkCelular.Select();
		}

		private void chkCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.chkEmail.Select();
		}

		private void chkEmail_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtObservaciones.Select();
		}

		private void txtNCotizacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cmbContactos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtObservaciones_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!this.barraDato1.bEditar) return;

			if (this.txtObservaciones.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una observación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtObservaciones.Select();
				e.Cancel = true;
				return;
			}
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			if (IdCliente == 0)
			{
				MessageBox.Show("Ingrese el número de cotización para ingresar nuevos contactos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNCotizacion.Select();
				return;
			}

			Cliente miCliente;			
			miCliente = new Cliente(true, IdCliente);
			
			miCliente.ShowDialog();			
			miCliente.Dispose();
			miCliente = null;

			this.cmbContactos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idClienteContacto, Contacto From ClienteContacto Where idCliente = {0}", IdCliente));

			this.barraDato1.Select();
		}

		private void chkCelular_BeforeCheckStateChanged(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!this.chkCelular.Checked)
				{
					if (this.txtCelular.Text.ToString().Length == 0)
					{
						MessageBox.Show("No ha registrado el número de celular del contacto seleccionado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.chkCelular.Select();
						e.Cancel = true;
						return;
					}
				}
			}
		}

		private void chkCelular_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkConvencional_BeforeCheckStateChanged(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!this.chkConvencional.Checked)
				{
					if (this.txtConvencional.Text.ToString().Length == 0)
					{
						MessageBox.Show("No ha registrado el número convencional del contacto seleccionado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.chkConvencional.Select();
						e.Cancel = true;
						return;
					}
				}
			}
		}

		private void chkEmail_BeforeCheckStateChanged(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!this.chkEmail.Checked)
				{
					if (this.txtEMail.Text.ToString().Length == 0)
					{
						MessageBox.Show("No ha registrado la dirección de Correo Electrónico del contacto seleccionado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.chkEmail.Select();
						e.Cancel = true;
						return;
					}
				}
			}
		}

		private void txtObservaciones_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void txtObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtObservaciones_Leave(object sender, System.EventArgs e)
		{
			drSeguimiento["Observaciones"] = this.txtObservaciones.Text.ToString();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (drSeguimiento == null)
			{
				MessageBox.Show("No hay registros.", "Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			using (frmBuscaGestionCotizacion miBusqueda = new frmBuscaGestionCotizacion())
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					try
					{						
						barraDato1.IdRegistro = (int)miBusqueda.grdPedido.ActiveRow.Cells["idSeguimiento"].Value;
						barraDato1.ProximoId(5);
					}
					catch{}
				}
			}
		}
	}
}

