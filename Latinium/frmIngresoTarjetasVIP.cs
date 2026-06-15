using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmIngresoTarjetasVIP.
	/// </summary>
	public class frmIngresoTarjetasVIP : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		public System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public System.Windows.Forms.Label lblCliente;
		public System.Windows.Forms.Label lblIdentificacion;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		public System.Windows.Forms.Label label7;
		public System.Windows.Forms.Label label4;
		public System.Windows.Forms.Label lblIce;
		public System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAbono;
		public System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodTarjeta;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnCancelar;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTurno;
		private System.Windows.Forms.Button btnAnular;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public System.Windows.Forms.Label lblBodega;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredefinida;
		public System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		public System.Windows.Forms.Button btnCliente;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		public System.Windows.Forms.Label label5;
		public System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefonoOtroga;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreRecibe;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefonoRecibe;
		private System.Windows.Forms.Button btnAsiento;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoRecibe;
		public System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreOtorga;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAnticipo;
		private System.Windows.Forms.Button btnImprimir;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdHistorialTarjetasVIP;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmIngresoTarjetasVIP()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmIngresoTarjetasVIP));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombreOtorga = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtTelefonoOtroga = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAbono = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtCodTarjeta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.txtTurno = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtBodegaPredefinida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnCliente = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label9 = new System.Windows.Forms.Label();
			this.txtApellidoRecibe = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombreRecibe = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtTelefonoRecibe = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.txtIdAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtIdHistorialTarjetasVIP = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreOtorga)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefonoOtroga)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAbono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefonoRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdHistorialTarjetasVIP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			this.SuspendLayout();
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
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.SystemColors.Control;
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 400);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 34;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(80, 400);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 35;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.SystemColors.Control;
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(157, 400);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 37;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.SystemColors.Control;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(234, 400);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 36;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 310;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(96, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 309;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(512, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 312;
			this.label1.Text = "Numero";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(568, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 21);
			this.txtNumero.TabIndex = 311;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 50);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(40, 16);
			this.lblCliente.TabIndex = 325;
			this.lblCliente.Text = "Cliente";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 26);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(52, 16);
			this.lblIdentificacion.TabIndex = 324;
			this.lblIdentificacion.Text = "RUC/C.I. ";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipoRuc
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance3;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 233;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(88, 24);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 326;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// txtNombreOtorga
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreOtorga.Appearance = appearance4;
			this.txtNombreOtorga.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreOtorga.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreOtorga.Enabled = false;
			this.txtNombreOtorga.Location = new System.Drawing.Point(88, 48);
			this.txtNombreOtorga.Name = "txtNombreOtorga";
			this.txtNombreOtorga.Size = new System.Drawing.Size(584, 21);
			this.txtNombreOtorga.TabIndex = 323;
			// 
			// txtRuc
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance5;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(240, 24);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 21);
			this.txtRuc.TabIndex = 322;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 74);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 328;
			this.label7.Text = "Teléfono";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTelefonoOtroga
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefonoOtroga.Appearance = appearance6;
			this.txtTelefonoOtroga.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefonoOtroga.Enabled = false;
			this.txtTelefonoOtroga.Location = new System.Drawing.Point(88, 72);
			this.txtTelefonoOtroga.Name = "txtTelefonoOtroga";
			this.txtTelefonoOtroga.Size = new System.Drawing.Size(144, 21);
			this.txtTelefonoOtroga.TabIndex = 327;
			this.txtTelefonoOtroga.ValueChanged += new System.EventHandler(this.txtTelefono_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 288);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 16);
			this.label4.TabIndex = 333;
			this.label4.Text = "Detalle";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(8, 359);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(49, 16);
			this.lblIce.TabIndex = 332;
			this.lblIce.Text = "Abono $.";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 266);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(81, 16);
			this.label2.TabIndex = 331;
			this.label2.Text = "Forma de pago";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance7;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 224;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 264);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(224, 21);
			this.cmbFormaPago.TabIndex = 334;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			// 
			// txtDetalle
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance8;
			this.txtDetalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Enabled = false;
			this.txtDetalle.Location = new System.Drawing.Point(96, 288);
			this.txtDetalle.Multiline = true;
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(592, 64);
			this.txtDetalle.TabIndex = 329;
			// 
			// txtAbono
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAbono.Appearance = appearance9;
			this.txtAbono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAbono.Enabled = false;
			this.txtAbono.FormatString = "#,##0.00";
			this.txtAbono.Location = new System.Drawing.Point(96, 357);
			this.txtAbono.Name = "txtAbono";
			this.txtAbono.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAbono.PromptChar = ' ';
			this.txtAbono.Size = new System.Drawing.Size(136, 21);
			this.txtAbono.TabIndex = 330;
			this.txtAbono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAbono_KeyDown);
			this.txtAbono.Click += new System.EventHandler(this.txtAbono_Click);
			this.txtAbono.Validating += new System.ComponentModel.CancelEventHandler(this.txtAbono_Validating);
			this.txtAbono.Enter += new System.EventHandler(this.txtAbono_Enter);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(488, 34);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(78, 16);
			this.label3.TabIndex = 335;
			this.label3.Text = "Codigo Tarjeta";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodTarjeta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodTarjeta.Appearance = appearance10;
			this.txtCodTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodTarjeta.Enabled = false;
			this.txtCodTarjeta.Location = new System.Drawing.Point(568, 32);
			this.txtCodTarjeta.Name = "txtCodTarjeta";
			this.txtCodTarjeta.Size = new System.Drawing.Size(120, 21);
			this.txtCodTarjeta.TabIndex = 336;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(-32, 384);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(720, 8);
			this.groupBox1.TabIndex = 340;
			this.groupBox1.TabStop = false;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(616, 400);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 345;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(464, 400);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 344;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// txtTurno
			// 
			this.txtTurno.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTurno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTurno.Enabled = false;
			this.txtTurno.Location = new System.Drawing.Point(672, 357);
			this.txtTurno.Name = "txtTurno";
			this.txtTurno.PromptChar = ' ';
			this.txtTurno.Size = new System.Drawing.Size(16, 21);
			this.txtTurno.TabIndex = 425;
			this.txtTurno.Visible = false;
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(536, 400);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 426;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(656, 357);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 427;
			this.txtEstado.Visible = false;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 34);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 428;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(96, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(280, 21);
			this.cmbBodega.TabIndex = 429;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// txtBodegaPredefinida
			// 
			this.txtBodegaPredefinida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredefinida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredefinida.Enabled = false;
			this.txtBodegaPredefinida.Location = new System.Drawing.Point(640, 357);
			this.txtBodegaPredefinida.Name = "txtBodegaPredefinida";
			this.txtBodegaPredefinida.PromptChar = ' ';
			this.txtBodegaPredefinida.Size = new System.Drawing.Size(16, 21);
			this.txtBodegaPredefinida.TabIndex = 430;
			this.txtBodegaPredefinida.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(528, 354);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 431;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(624, 357);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdCliente.TabIndex = 432;
			this.txtIdCliente.Visible = false;
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(376, 23);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 433;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.btnCliente);
			this.groupBox2.Controls.Add(this.lblIdentificacion);
			this.groupBox2.Controls.Add(this.lblCliente);
			this.groupBox2.Controls.Add(this.cmbTipoRuc);
			this.groupBox2.Controls.Add(this.txtNombreOtorga);
			this.groupBox2.Controls.Add(this.txtRuc);
			this.groupBox2.Controls.Add(this.label7);
			this.groupBox2.Controls.Add(this.txtTelefonoOtroga);
			this.groupBox2.Location = new System.Drawing.Point(8, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(680, 104);
			this.groupBox2.TabIndex = 434;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Cliente Otorga";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.label9);
			this.groupBox3.Controls.Add(this.txtApellidoRecibe);
			this.groupBox3.Controls.Add(this.txtNombreRecibe);
			this.groupBox3.Controls.Add(this.label5);
			this.groupBox3.Controls.Add(this.label8);
			this.groupBox3.Controls.Add(this.txtTelefonoRecibe);
			this.groupBox3.Location = new System.Drawing.Point(8, 176);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(680, 80);
			this.groupBox3.TabIndex = 435;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Cliente Recibe";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(360, 26);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(48, 16);
			this.label9.TabIndex = 439;
			this.label9.Text = "Apellido:";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtApellidoRecibe
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoRecibe.Appearance = appearance11;
			this.txtApellidoRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoRecibe.Enabled = false;
			this.txtApellidoRecibe.Location = new System.Drawing.Point(408, 24);
			this.txtApellidoRecibe.Name = "txtApellidoRecibe";
			this.txtApellidoRecibe.Size = new System.Drawing.Size(264, 21);
			this.txtApellidoRecibe.TabIndex = 438;
			// 
			// txtNombreRecibe
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreRecibe.Appearance = appearance12;
			this.txtNombreRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreRecibe.Enabled = false;
			this.txtNombreRecibe.Location = new System.Drawing.Point(88, 24);
			this.txtNombreRecibe.Name = "txtNombreRecibe";
			this.txtNombreRecibe.Size = new System.Drawing.Size(264, 21);
			this.txtNombreRecibe.TabIndex = 436;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 26);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(48, 16);
			this.label5.TabIndex = 437;
			this.label5.Text = "Nombre:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(8, 50);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(48, 16);
			this.label8.TabIndex = 437;
			this.label8.Text = "Teléfono";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTelefonoRecibe
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefonoRecibe.Appearance = appearance13;
			this.txtTelefonoRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefonoRecibe.Enabled = false;
			this.txtTelefonoRecibe.Location = new System.Drawing.Point(88, 48);
			this.txtTelefonoRecibe.Name = "txtTelefonoRecibe";
			this.txtTelefonoRecibe.Size = new System.Drawing.Size(144, 21);
			this.txtTelefonoRecibe.TabIndex = 436;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.Location = new System.Drawing.Point(312, 400);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(72, 23);
			this.btnAsiento.TabIndex = 436;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// txtIdAnticipo
			// 
			this.txtIdAnticipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAnticipo.Enabled = false;
			this.txtIdAnticipo.Location = new System.Drawing.Point(608, 357);
			this.txtIdAnticipo.Name = "txtIdAnticipo";
			this.txtIdAnticipo.PromptChar = ' ';
			this.txtIdAnticipo.Size = new System.Drawing.Size(16, 21);
			this.txtIdAnticipo.TabIndex = 437;
			this.txtIdAnticipo.Visible = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(384, 400);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 438;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtIdHistorialTarjetasVIP
			// 
			this.txtIdHistorialTarjetasVIP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdHistorialTarjetasVIP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdHistorialTarjetasVIP.Enabled = false;
			this.txtIdHistorialTarjetasVIP.Location = new System.Drawing.Point(592, 357);
			this.txtIdHistorialTarjetasVIP.Name = "txtIdHistorialTarjetasVIP";
			this.txtIdHistorialTarjetasVIP.PromptChar = ' ';
			this.txtIdHistorialTarjetasVIP.Size = new System.Drawing.Size(16, 21);
			this.txtIdHistorialTarjetasVIP.TabIndex = 439;
			this.txtIdHistorialTarjetasVIP.Visible = false;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(576, 357);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 21);
			this.txtIdAsiento.TabIndex = 440;
			this.txtIdAsiento.Visible = false;
			// 
			// frmIngresoTarjetasVIP
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(696, 430);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdHistorialTarjetasVIP);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.txtIdAnticipo);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtBodegaPredefinida);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.txtTurno);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtCodTarjeta);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.txtAbono);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmIngresoTarjetasVIP";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Ingreso Tarjetas V.I.P";
			this.Load += new System.EventHandler(this.frmIngresoTarjetasVIP_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreOtorga)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefonoOtroga)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAbono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			this.groupBox2.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefonoRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdHistorialTarjetasVIP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodega = 0;
		int iHistorialAnticipoTarjetaVIP = 0;
		string sRuc = "";
		decimal dCobro = 0.00m;
		decimal dValor = 0.00m;
		decimal dSaldo = 0.00m;

		bool bNuevo = false;
		bool bEdicion = false;
		
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmIngresoTarjetasVIP_Load(object sender, System.EventArgs e)
		{
			#region Seguridad
			miAcceso = new Acceso(cdsSeteoF, "1130");

			if (!Funcion.Permiso("2036", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Tarjetas V.I.P", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}

			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;
			#endregion Seguridad
			
			#region Fecha
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			#endregion Fecha

			#region Formas de Pago
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 3)"));
			#endregion Formas de Pago

			#region Turno Caja
			string sSQLT = string.Format("Exec RetornaTurno '{0}'", dtHoy.ToString("yyyyMMdd HH:mm"));		
			this.txtTurno.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLT);
			#endregion Turno Caja

			#region Bodega
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredefinida.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			#endregion Bodega
		}

		private void txtTelefono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdCliente.Value = 0;
			this.txtIdHistorialTarjetasVIP.Value = 0;
			this.txtIdAnticipo.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtCodTarjeta.Text = "";
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtNombreOtorga.Text = "";
			this.txtApellidoRecibe.Text = "";
			this.txtNombreRecibe.Text = "";
			this.txtTelefonoOtroga.Text = "";
			this.txtTelefonoRecibe.Text = "";
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.txtDetalle.Text = "";
			this.txtAbono.Value = 0.00m;
			this.lblEstado.Text = "";

			this.dtFecha.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.txtDetalle.Enabled = false;
			this.txtAbono.Enabled = false;
			this.txtNombreRecibe.Enabled = false;
			this.txtApellidoRecibe.Enabled = false;
			this.txtTelefonoRecibe.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnCliente.Enabled = false;
			this.btnAsiento.Enabled = false;
			
			dValor = 0.00m;
			dSaldo = 0.00m;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Value = 1;
			else this.cmbBodega.Value = (int)this.txtBodegaPredefinida.Value;

			#region Fecha
			this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");
			#endregion Fecha

			this.cmbTipoRuc.Value = 5;
			this.cmbFormaPago.Value = 1;			
			this.txtEstado.Value = 1;
			this.lblEstado.Text = "PENDIENTE";

			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Enabled = true;
			else this.cmbBodega.Enabled = false;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtDetalle.Enabled = true;
			this.txtAbono.Enabled = true;
			this.txtNombreRecibe.Enabled = true;
			this.txtApellidoRecibe.Enabled = true;
			this.txtTelefonoRecibe.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnCliente.Enabled = true;			

			if (this.cmbBodega.Enabled) this.cmbBodega.Focus();
			else if (this.dtFecha.Enabled) this.dtFecha.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null) this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 24, cdsSeteoF);
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text, 24, 1, (int)this.cmbTipoRuc.Value, 0, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtRuc.Text = miBusqueda.Ruc;
						this.cmbFormaPago.Focus();
					}
				}
			}

			if (e.KeyCode == Keys.Enter) this.cmbFormaPago.Focus();
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
						
			FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true, false);
				
			CargaCliente();
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtRuc.Text.ToString(), cdsSeteoF);

			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.cmbTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombreOtorga.Text = sParametros[4];
			if (sParametros[8] != null) this.txtTelefonoOtroga.Text = sParametros[8];		
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtIdCliente.Value = 0;
				this.txtNombreOtorga.Text = "";
				
				CargaCliente();

				if ((int)this.txtIdCliente.Value == 0)
				{					
					FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true, false);
					CargaCliente();				
				}

				if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
				if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;				
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (txtRuc.Text.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese el Número de Identificación del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					e.Cancel = true;
				}
				else
				{					
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", txtRuc.Text,(int)this.cmbTipoRuc.Value);
					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show("Número de Identificación Incorrecta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						e.Cancel = true;
					}							
				}
			}
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if ((int)this.cmbTipoRuc.Value == 4) this.txtRuc.MaxLength = 13;
				if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;
			}
		}

		private void txtAbono_Click(object sender, System.EventArgs e)
		{
			this.txtAbono.SelectAll();
		}

		private void txtAbono_Enter(object sender, System.EventArgs e)
		{
			this.txtAbono.SelectAll();
		}

		private void txtAbono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtAbono_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				if (this.txtAbono.Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
				else if (Convert.ToDecimal(this.txtAbono.Value) == 0.00m)
				{
					MessageBox.Show("El Valor debe ser Mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iOrigen = 10;
			string sCodAnticipo = "";

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Cobros", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
			
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, 1, cdsSeteoF)) return;

			if (!Validacion.vbTexto(this.txtNombreRecibe, 3, 300, "Nombre Cliente Recibe")) return;

			if (!Validacion.vbTexto(this.txtApellidoRecibe, 3, 300, "Apellido Cliente Recibe")) return;

			if (!Validacion.vbTexto(this.txtTelefonoRecibe, 10, 300, "Telefono Cliente Recibe")) return;

			if (!Validacion.vbTexto(this.txtDetalle, 10, 300, "Detalle")) return;			
			
			if (Convert.ToDecimal(this.txtAbono.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el valor del Abono", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAbono.Focus();
				return;
			}

			if (!Validacion.vbCampoDecimalVacio(this.txtAbono, "Ingrese el Valor", dCobro, 100000.00m)) return;
						
			if (!bNuevo && (Convert.ToDecimal(this.txtAbono.Value) < (Convert.ToDecimal(this.txtAbono.Value) - dSaldo)))
			{
				MessageBox.Show("El Abono no puede ser menor al Valor ya Utilizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAbono.Focus();
				return;
			}
			#endregion Validacion

			using (VentaFPagoLot miCobro = new VentaFPagoLot ((int)this.txtIdHistorialTarjetasVIP.Value, (int)this.cmbFormaPago.Value, Convert.ToDecimal(this.txtAbono.Value), iOrigen, 10))
			{
				if (DialogResult.OK == miCobro.ShowDialog())
				{
					#region Guarda Anticipo
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

							#region Numeracion
							if (bNuevo)
							{
								string sSQLNumero = string.Format("Exec NumeracionLocales 90, 0, 0");
								oCmdActualiza.CommandText = sSQLNumero;
								this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
							}
							#endregion Numeracion

							#region Codigo Tarjeta
							string sSQLNumeroAnt = string.Format("Exec AnticipoGeneraCodigoTarjetaVIP '{0}', '{1}', '{2}'", this.txtNombreRecibe.Text.ToString(), 
								this.txtApellidoRecibe.Text.ToString(), this.txtTelefonoRecibe.Text.ToString());
							oCmdActualiza.CommandText = sSQLNumeroAnt;
							sCodAnticipo = oCmdActualiza.ExecuteScalar().ToString();
							#endregion Codigo Tarjeta
							
							#region Historial Tarjetas VIP
							DateTime dtFecha = (DateTime) this.dtFecha.Value;

							string sSQLHistorial = string.Format("Exec HistorialAnticiposTarjetasVIP {0}, {1}, '{2}', {3}, '{4}', '{5}', '{6}', '{7}', {8}, '{9}', {10}, '{11}', {12}, {13}, {14}", 
								(int)this.txtIdHistorialTarjetasVIP.Value, (int)this.txtIdCliente.Value, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, sCodAnticipo, this.txtNombreRecibe.Text.ToString(),
								this.txtApellidoRecibe.Text.ToString(), this.txtTelefonoRecibe.Text.ToString(), Convert.ToDecimal(this.txtAbono.Value), this.txtNumero.Text, 
								(int)this.cmbFormaPago.Value, this.txtDetalle.Text.ToString(), (int)this.txtIdAsiento.Value, Convert.ToDecimal(this.txtAbono.Value), this.txtTurno.Value);
							oCmdActualiza.CommandText = sSQLHistorial;
							this.txtIdHistorialTarjetasVIP.Value = (int)oCmdActualiza.ExecuteScalar();
							#endregion Historial Tarjetas VIP

							#region Guarda Cobro
							if (miCobro.ultraGrid1.Rows.Count > 0)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
								{
									int idDetFormaPago = 0;
									int idBanco = 0;
									int idTarjeta = 0;
									int idCuenta = 0;
									int idPlan = 0;
									int idPlazo = 0;
									decimal dComTarjeta = 0.00m;
									DateTime dtFecCaducidad = DateTime.Today;
							
									if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
									if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
									if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
									if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
									if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
									if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
									if ((int)dr.Cells["idFormaPago"].Value == 3 || (int)dr.Cells["idFormaPago"].Value == 11) if (dr.Cells["FechaCaducidad"].Value != System.DBNull.Value) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
									if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal )dr.Cells["ValorUsoTarjeta"].Value;

									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
										idDetFormaPago, 
										(int)this.txtIdHistorialTarjetasVIP.Value, iOrigen,
										(int)dr.Cells["idFormaPago"].Value, (decimal )dr.Cells["Valor"].Value,
										idBanco, idTarjeta,
										dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"),								
										dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
										idCuenta, 
										dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(),
										idPlan, idPlazo, dComTarjeta);
									oCmdActualiza.CommandText = sSQLCobro;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
							#endregion Guarda Cobro

							oTransaction.Commit();
							
							oConexion.Close();
		
							#region Controles
							this.cmbBodega.Enabled = false;
							this.dtFecha.Enabled = false;
							this.cmbTipoRuc.Enabled = false;
							this.txtRuc.Enabled = false;
							this.cmbFormaPago.Enabled = false;
							this.txtDetalle.Enabled = false;
							this.txtAbono.Enabled = false;
							this.txtApellidoRecibe.Enabled = false;
							this.txtNombreRecibe.Enabled = false;
							this.txtTelefonoRecibe.Enabled = false;

							if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
							if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
							if (miAcceso.Editar) this.btnEditar.Enabled = true;
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							if (miAcceso.Anular) this.btnAnular.Enabled = true;
							this.btnGuardar.Enabled = false;
							this.btnCliente.Enabled = false;

							bNuevo = false;
							bEdicion = false;
							#endregion Controles
								
							#region Impresion
							using (frmNumeracionRecibo Numeracion = new frmNumeracionRecibo (this.txtNumero.Text.ToString()))
							{
								if (DialogResult.OK == Numeracion.ShowDialog())
									this.btnImprimir_Click(sender, e);
							}
							#endregion Impresion																
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						finally
						{
							oConexion.Close();
						}
					}
					#endregion Guarda Anticipo

					this.Consulta((int)this.txtIdHistorialTarjetasVIP.Value);

				}
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From HistorialTarjetasVIP Where idHistorialTarjetasVIP = {0}", (int)this.txtIdHistorialTarjetasVIP.Value));
			if (iEstado == 2) 
			{
				MessageBox.Show("No puede crear un asiento, el documento esta UTILIZADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.Consulta((int)this.txtIdAnticipo.Value);
			}			
			else if (iEstado == 3) 
			{
				MessageBox.Show("No puede crear un asiento, el documento esta ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.Consulta((int)this.txtIdAnticipo.Value);
			}			
			else
			{				
				string sSQLExisteAsiento = string.Format("Exec ValidaAsientoTarjetaVIP {0}", (int)this.txtIdHistorialTarjetasVIP.Value);
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLExisteAsiento, true);

				if (!stMensaje.StartsWith("No Existe"))
				{			
					if (DialogResult.Yes == MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{					
						this.Cursor = Cursors.WaitCursor;

						#region Crea Asiento
						int iOrigen = 10;

						string sSQLAsiento = string.Format("Exec CreaAsientoTarjetasVIP {0}, {1}, {2}", (int)this.txtIdAsiento.Value, (int)this.txtIdHistorialTarjetasVIP.Value, iOrigen);
						this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

						#endregion Crea Asiento

						Cursor = Cursors.Default;
					}
				}
				else
				{
					this.Cursor = Cursors.WaitCursor;

					#region Crea Asiento
					int iOrigen = 10;

					string sSQLAsiento = string.Format("Exec CreaAsientoTarjetasVIP {0}, {1}, {2}", (int)this.txtIdAsiento.Value, (int)this.txtIdHistorialTarjetasVIP.Value, iOrigen);
					this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

					#endregion Crea Asiento

					Cursor = Cursors.Default;
				}
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idCompraRet = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idHistorialTarjetasVIP";
			idCompraRet.Value = (int)this.txtIdHistorialTarjetasVIP.Value;
			pfRet.CurrentValues.Add (idCompraRet);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("ImpresionTarjetasVIP.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("");
			miRepor.crVista.ParameterFieldInfo = paramFields;

			miRepor.Show();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmConsultaAnticiposTarjetas Busqueda = new frmConsultaAnticiposTarjetas())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridInformacion.ActiveRow.Cells["idHistorialTarjetasVIP"].Value);
				}
			}
		}

		private void Consulta(int idHistorialTarjetasVIP)
		{
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;

				int IBodega = 0;

				string sSQLBusqueda = string.Format("Exec ConsultaIndividualAnticiposTarjetasVIP {0}", idHistorialTarjetasVIP);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdHistorialTarjetasVIP.Value = dr.GetInt32(0);
					this.dtFecha.Value = dr.GetDateTime(1);
					IBodega = dr.GetInt32(2);
					this.cmbTipoRuc.Value = dr.GetInt32(3);
					this.txtNumero.Text = dr.GetValue(4).ToString();
					this.txtCodTarjeta.Text = dr.GetValue(5).ToString();
					this.txtNombreOtorga.Text = dr.GetValue(6).ToString();
					this.txtTelefonoOtroga.Text = dr.GetValue(7).ToString();
					this.txtRuc.Text = dr.GetValue(8).ToString();
					this.txtNombreRecibe.Text = dr.GetValue(9).ToString();
					this.txtApellidoRecibe.Text = dr.GetValue(10).ToString();
					this.txtTelefonoRecibe.Text = dr.GetValue(11).ToString();
					this.txtAbono.Value = dr.GetDecimal(12);
					this.cmbFormaPago.Value = dr.GetInt32(13);
					this.txtDetalle.Text = dr.GetValue(14).ToString();
					this.txtEstado.Value = dr.GetInt32(15);
					this.lblEstado.Text = dr.GetValue(16).ToString();
					this.txtIdCliente.Value = dr.GetInt32(17);
					if (dr.GetValue(19) != System.DBNull.Value)this.txtIdAsiento.Value = dr.GetInt32(18);
					this.txtTurno.Value = dr.GetInt32(19);
				}
				dr.Close();
			
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
				if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "UTILIZADO";
				if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "ANULADO";
						
				if ((int)this.txtEstado.Value != 2)
				{
					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Cobros"))
					{					
						if (miAcceso.Editar) this.btnEditar.Enabled = true;

						if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						if (miAcceso.Anular) this.btnAnular.Enabled = true;
					}

					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;

					dValor = Convert.ToDecimal(this.txtAbono.Value);			
				}
				else if ((int)this.txtEstado.Value == 1)
				{
					this.btnAnular.Enabled = true;
					this.btnAnular.Text = "Eliminar";
				}
								
				this.cmbBodega.Value = IBodega;
				this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 24, cdsSeteoF);
				
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right) 
			{
				if ((int)this.txtIdAsiento.Value == 0)
				{
					int idAsiento = 0;
					idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From HistorialTarjetasVIP Where idHistorialTarjetasVIP = {0}", (int)this.txtIdHistorialTarjetasVIP.Value));
					if (idAsiento > 0) this.txtIdAsiento.Value = idAsiento;
				}
				
				if ((int)this.txtIdAsiento.Value == 0)
				{
					MessageBox.Show("No existe un asiento creado de esta transacción");				
				}
				else
				{
					Cursor = Cursors.WaitCursor;
					frmAsientos Asiento = new frmAsientos(((int)this.txtIdAsiento.Value));
					Asiento.MdiParent = MdiParent;
					Asiento.Show();
					Cursor = Cursors.Default;			
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El Documento esta UTILIZADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El Documento esta ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtIdHistorialTarjetasVIP.Value, 90, cdsSeteoF)) return;

			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtDetalle.Enabled = true;
			this.txtNombreRecibe.Enabled = true;
			this.txtApellidoRecibe.Enabled = true;
      this.txtTelefonoRecibe.Enabled = true;			
			this.txtAbono.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnCliente.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (!bNuevo)
			{
				this.Consulta((int)this.txtIdHistorialTarjetasVIP.Value);

				if ((int)this.txtEstado.Value == 3)
				{
					MessageBox.Show("El Documento ya esta Anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
					return;
				}

				if ((int)this.txtEstado.Value == 2)
				{
					MessageBox.Show("No puede Anular el Documento Esta UTILIZADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
					return;
				}

				string sMensaje = "Anular";

				if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de '{0}'?", sMensaje), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					if ((int)this.txtEstado.Value == 1) 
					{	
						using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
						{
							if (DialogResult.OK == miNota.ShowDialog())
							{							
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulaTarjetasVIP {0}, {1}, {2}, '{3}'", (int)this.txtIdHistorialTarjetasVIP.Value, (int)this.txtIdAsiento.Value, 10, miNota.txtNotas.Text.ToString()), true);
					
								this.Consulta((int)this.txtIdHistorialTarjetasVIP.Value);							
							}
						}
					}
				}
			}
		}
	}
}
