using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAnticiposProveedor.
	/// </summary>
	public class frmAnticiposProveedor : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label lblIce;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAnticipo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.ComponentModel.IContainer components;
		
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredefinida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private System.Windows.Forms.ImageList imageList1;

		bool bNuevo = false;
		bool bEdicion = false;
		decimal dValor = 0.00m;
		int idAnticipo = 0;
		private Acceso miAcceso;

		public frmAnticiposProveedor()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmAnticiposProveedor(int IdAnticipo)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			idAnticipo = IdAnticipo;

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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAnticiposProveedor));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.label4 = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtBodegaPredefinida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 208);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 269;
			this.label4.Text = "Detalle";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(8, 275);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(43, 16);
			this.lblIce.TabIndex = 268;
			this.lblIce.Text = "Valor $.";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 171);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 16);
			this.label2.TabIndex = 267;
			this.label2.Text = "Forma de pago";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 139);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 266;
			this.label7.Text = "Teléfono";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 107);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 265;
			this.label3.Text = "Dirección";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 75);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(55, 16);
			this.lblCliente.TabIndex = 264;
			this.lblCliente.Text = "Proveedor";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 43);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(48, 16);
			this.lblIdentificacion.TabIndex = 263;
			this.lblIdentificacion.Text = "RUC/C.I.";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(464, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 262;
			this.label1.Text = "Número";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(248, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 261;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 260;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance1;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 208;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(104, 168);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(208, 21);
			this.cmbFormaPago.TabIndex = 277;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			this.cmbFormaPago.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbFormaPago_InitializeLayout);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 180;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Codigo";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(296, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(120, 21);
			this.cmbBodega.TabIndex = 274;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// cmbTipoRuc
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance2;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 233;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(104, 40);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(160, 21);
			this.cmbTipoRuc.TabIndex = 273;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(528, 72);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdCliente.TabIndex = 271;
			this.txtIdCliente.Visible = false;
			this.txtIdCliente.ValueChanged += new System.EventHandler(this.txtIdCliente_ValueChanged);
			// 
			// txtIdAnticipo
			// 
			this.txtIdAnticipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAnticipo.Enabled = false;
			this.txtIdAnticipo.Location = new System.Drawing.Point(544, 72);
			this.txtIdAnticipo.Name = "txtIdAnticipo";
			this.txtIdAnticipo.PromptChar = ' ';
			this.txtIdAnticipo.Size = new System.Drawing.Size(16, 22);
			this.txtIdAnticipo.TabIndex = 270;
			this.txtIdAnticipo.Visible = false;
			this.txtIdAnticipo.ValueChanged += new System.EventHandler(this.txtIdAnticipo_ValueChanged);
			// 
			// txtDetalle
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance3;
			this.txtDetalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Enabled = false;
			this.txtDetalle.Location = new System.Drawing.Point(104, 200);
			this.txtDetalle.Multiline = true;
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(536, 64);
			this.txtDetalle.TabIndex = 258;
			this.txtDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDetalle_KeyDown);
			// 
			// txtValor
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance4;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(104, 274);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 22);
			this.txtValor.TabIndex = 259;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Leave += new System.EventHandler(this.txtValor_Leave);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// txtTelefono
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance5;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(104, 136);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 257;
			// 
			// txtDireccion
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance6;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(104, 104);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(536, 22);
			this.txtDireccion.TabIndex = 256;
			// 
			// txtNombre
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance7;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(104, 72);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(360, 22);
			this.txtNombre.TabIndex = 255;
			// 
			// txtRuc
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance8;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(352, 40);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 254;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNumero
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance9;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(528, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 253;
			// 
			// dtFecha
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance10;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.CausesValidation = false;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(104, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 252;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(404, 328);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 283;
			this.btnAsiento.Text = "Egreso";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(10, 328);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 278;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(85, 328);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 279;
			this.btnConsultar.Text = "Consultar";
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
			this.btnEditar.Location = new System.Drawing.Point(166, 328);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 282;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(245, 328);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 281;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(325, 328);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 280;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(487, 328);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 284;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(566, 328);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 285;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 312);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(648, 8);
			this.groupBox1.TabIndex = 286;
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
			// txtBodegaPredefinida
			// 
			this.txtBodegaPredefinida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredefinida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredefinida.Enabled = false;
			this.txtBodegaPredefinida.Location = new System.Drawing.Point(512, 72);
			this.txtBodegaPredefinida.Name = "txtBodegaPredefinida";
			this.txtBodegaPredefinida.PromptChar = ' ';
			this.txtBodegaPredefinida.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredefinida.TabIndex = 287;
			this.txtBodegaPredefinida.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(328, 272);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 288;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(496, 72);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 289;
			this.txtEstado.Visible = false;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(480, 72);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 290;
			this.txtIdAsiento.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(368, 171);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 16);
			this.label5.TabIndex = 292;
			this.label5.Text = "Cuenta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbCuenta
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance11;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource1;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 300;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(432, 169);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(208, 21);
			this.cmbCuenta.TabIndex = 293;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown_1);
			// 
			// frmAnticiposProveedor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(648, 358);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtBodegaPredefinida);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.txtIdAnticipo);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAnticiposProveedor";
			this.Text = "Anticipos a Proveedores";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmAnticiposProveedor_Closing);
			this.Load += new System.EventHandler(this.frmAnticiposProveedor_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void txtIdCliente_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtIdAnticipo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void Consulta (int idAnticipo)
		{			
			try
			{						
				string sSQL = string.Format("Exec ConsultaIndividualAnticipos {0}", idAnticipo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{		
					this.txtIdAnticipo.Value = dr.GetInt32(0);
					this.dtFecha.Value = dr.GetDateTime(1);
					this.txtNumero.Text = dr.GetString(2);
					this.cmbBodega.Value = dr.GetInt32(5);					
					this.txtIdCliente.Value = dr.GetInt32(7);
					this.cmbTipoRuc.Value = dr.GetInt32(8);
					this.txtRuc.Text = dr.GetString(9);
					this.txtNombre.Text = dr.GetString(10);
					this.txtDireccion.Text = dr.GetValue(11).ToString();
					this.txtTelefono.Text = dr.GetValue(12).ToString();
					this.cmbFormaPago.Value = dr.GetInt32(13);
					if (dr.GetValue(14) != System.DBNull.Value) this.cmbCuenta.Value = dr.GetInt32(14);
					this.txtDetalle.Text = dr.GetValue(15).ToString();
					this.txtEstado.Value = dr.GetInt32(16);
					this.txtValor.Value = dr.GetDecimal(17);
					if (dr.GetValue(18) != System.DBNull.Value) this.txtIdAsiento.Value = dr.GetInt32(18);					
				}
				dr.Close();
	
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
				if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "ANULADO";
				if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "LIQUIDADO";
				if ((int)this.txtEstado.Value == 4) this.lblEstado.Text = "PAGO AUTORIZADO";
				if ((int)this.txtEstado.Value == 5) this.lblEstado.Text = "POR LIQUIDAR";

				if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
				{
					if (miAcceso.Editar) if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;											
					if (miAcceso.Anular) if ((int)this.txtEstado.Value == 1) this.btnAnular.Enabled = true;
				}

				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				if ((int)this.txtIdAsiento.Value > 0) if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;

				this.btnCancelar.Enabled = true;

				dValor = (decimal)this.txtValor.Value;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void frmAnticiposProveedor_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0436");

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			this.txtBodegaPredefinida.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
						
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 25", MenuLatinium.IdUsuario));
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 0");
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 7)"));
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasCierreCaja");

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;

 			if (idAnticipo > 0) this.Consulta(idAnticipo);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdAnticipo.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtDireccion.Text = "";
			this.txtTelefono.Text = "";
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.txtDetalle.Text = "";
			this.txtValor.Value = 0.00m;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.lblEstado.Text = "";
			this.dtFecha.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.txtDetalle.Enabled = false;
			this.txtValor.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.cmbCuenta.Enabled = false;

			dValor = 0.00m;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.dtFecha.Value = DateTime.Today;

			if ((int)this.txtBodegaPredefinida.Value > 0) this.cmbBodega.Value = (int)this.txtBodegaPredefinida.Value;
			else
			{
				this.cmbBodega.Value = 1;
				this.cmbBodega.Enabled = true;
			}

			if (miAcceso.CambiarFecha)this.dtFecha.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.txtDetalle.Enabled = true;
			this.txtValor.Enabled = true;
			
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			this.cmbTipoRuc.Value = 1;
			this.cmbFormaPago.Value = 2;
			this.txtEstado.Value = 1;

			bNuevo = true;
			bEdicion = true;

			this.txtRuc.Focus();	
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if ((int) this.txtBodegaPredefinida.Value > 0) this.cmbTipoRuc.Focus(); else this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoRuc.Focus();
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtRuc.Text.ToString(), 25, 2, 1, 0, false))
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

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void txtDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{				
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.cmbTipoRuc.Value);
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;				
				}
				else
				{
					string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', 25, 0, 0", this.txtRuc.Text.ToString());
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
					drC.Read();
					if (!drC.HasRows)
					{
						drC.Close();		
						MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();
						e.Cancel = true;					
					}			
					drC.Close();
				}
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Proveedor
				string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', 25, 0, 0", this.txtRuc.Text.ToString());
				SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
				drC.Read();
				if (drC.HasRows)
				{
					this.txtIdCliente.Value = (int) drC.GetValue(0);
					this.txtRuc.Text = (string) drC.GetValue(1);
					this.txtNombre.Text = (string) drC.GetValue(2);
					this.txtDireccion.Text = (string) drC.GetValue(3);
					this.txtTelefono.Text = (string) drC.GetValue(4);
				}
				drC.Close();		
				#endregion Busca Proveedor							
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta ((int)this.txtIdAnticipo.Value);

			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene permisos para anular.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El registro esta anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
			}
			else if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede anular el documento esta cancelado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
			}
			else if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede anular el documento esta Autorizado el Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 5)
			{
				MessageBox.Show("No puede anular el documento ya se Pago al Proveedor, para Anular este Anticipo Anule Primero el Egreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de anular?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{			
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("AnulaAnticiposProveedor {0}, {1}, {2}, {3}, '{4}'", 
							(int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.cmbFormaPago.Value, (int)this.txtEstado.Value, miNota.txtNotas.Text.ToString()), true);
					
						this.lblEstado.Text = "ANULADO";

						this.btnAnular.Enabled = false;
					}
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta ((int)this.txtIdAnticipo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede editar el documento esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede editar el documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("No puede editar el documento esta Autorizado su Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else if ((int)this.txtEstado.Value == 5)
			{
				MessageBox.Show("No puede editar el documento ya se Pago al Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
			}
			else
			{
				
				if (miAcceso.CambiarFecha)this.dtFecha.Enabled = true;

				if ((int)this.txtBodegaPredefinida.Value > 0) this.cmbBodega.Enabled = true;

				this.cmbTipoRuc.Enabled = true;
				this.txtRuc.Enabled = true;
				this.cmbFormaPago.Enabled = true;
				this.cmbCuenta.Enabled = true;
				this.txtDetalle.Enabled = true;
				this.txtValor.Enabled = true;

				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = true;

				bNuevo = false;
				bEdicion = true;
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.dtFecha.Focus();
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
			{
				DateTime dtFechaVP = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();			
			}
			else
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtRuc.Text.ToString(), (int)this.cmbTipoRuc.Value);
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();			
				}
				else
				{
					string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', 25, 0, 0", this.txtRuc.Text.ToString());
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				
					drC.Read();
					if (!drC.HasRows)
					{
						drC.Close();		
						MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();			
						drC.Close();
					}			
					else
					{
						drC.Close();

						if (this.txtDetalle.Text.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese el detalle del Anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.txtDetalle.Focus();
						}
						else if (Convert.ToDecimal(this.txtValor.Value) == 0.00m)
						{
							MessageBox.Show("Ingrese el Valor del Anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.txtValor.Focus();				
						}
						else
						{
							DateTime dtFecha = (DateTime) this.dtFecha.Value;

							if (bNuevo) this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 25, 1, 0"));

							int idCuenta = 0;
							if (this.cmbCuenta.ActiveRow != null) idCuenta = (int)this.cmbCuenta.Value;

							string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}", 
								(int)this.txtIdAnticipo.Value, 1, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), (int)this.txtIdCliente.Value, (int)this.cmbFormaPago.Value, idCuenta, this.txtDetalle.Text.ToString(), (int)this.txtEstado.Value, Convert.ToDecimal(this.txtValor.Value));
							this.txtIdAnticipo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGraba);

							#region Autorizacion de Pago
							if ((int)this.cmbFormaPago.Value != 1)
							{
								string sSQLAutMod = string.Format("Exec CreaAutorizacionesDePago 5, {0}, '{1}', '{2}', {3}, {4}, '1 DE 1', '{5}', {6}", 
									(int)this.txtIdAnticipo.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDecimal(this.txtValor.Value),
									(int)this.cmbFormaPago.Value, this.txtDetalle.Text.ToString(), idCuenta);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAutMod, true);			
							}			
							#endregion Autorizacion de Pago
			
							this.lblEstado.Text = "PENDIENTE";
							
							this.dtFecha.Enabled = false;
							this.cmbBodega.Enabled = false;
							this.cmbTipoRuc.Enabled = false;
							this.txtRuc.Enabled = false;
							this.cmbFormaPago.Enabled = false;
							this.cmbCuenta.Enabled = false;
							this.txtDetalle.Enabled = false;
							this.txtValor.Enabled = false;

							if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
							if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
							if ((int)this.txtEstado.Value == 1 && miAcceso.Editar) this.btnEditar.Enabled = true;
							if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
							this.btnGuardar.Enabled = false;
							if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
							if (miAcceso.Anular || MenuLatinium.iNivel == 1) this.btnAnular.Enabled = true;
			
							bNuevo = false;
							bEdicion = false;
						}
					}
				}
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtValor_Leave(object sender, System.EventArgs e)
		{
			if (this.txtValor.Value == System.DBNull.Value) this.txtValor.Value = 0;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaAnticipos Busqueda = new frmBuscaAnticipos((int)this.txtBodegaPredefinida.Value, 1))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta ((int)Busqueda.ultraGrid1.ActiveRow.Cells["idAnticipo"].Value);
				}
			}
		}

		private void cmbFormaPago_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDetalle.Focus();
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{	
			this.Consulta ((int)this.txtIdAnticipo.Value);

			if (!miAcceso.CrearAsiento)
			{
				MessageBox.Show("No tiene Acceso a Ver Asientos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);		
			}
			else if (bEdicion) 
			{
				MessageBox.Show("No puede ver un asiento mientras edita el Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if ((int)this.txtEstado.Value == 1) 
			{
				MessageBox.Show("El Documento no esta pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if ((int)this.txtEstado.Value == 2) 
			{
				MessageBox.Show("No puede ver un asiento de un Documento Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else if ((int)this.txtEstado.Value == 4) 
			{
				MessageBox.Show("El Documento no esta pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
			else
			{
				this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Anticipos Where idAnticipo = {0}", (int)this.txtIdAnticipo.Value));

				if ((int)this.txtIdAsiento.Value == 0)
				{
					MessageBox.Show("No existe un asiento creado de esta transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
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

		private void cmbCuenta_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDetalle.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Consulta ((int)this.txtIdAnticipo.Value);

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("No puede imprimir el documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);			
				return;
			}

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfidAsiento = new ParameterField ();
			ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
			pfidAsiento.ParameterFieldName = "@idAnticipo";
			ValIdAsiento.Value = (int)this.txtIdAnticipo.Value;
			pfidAsiento.CurrentValues.Add (ValIdAsiento);
			paramFields.Add (pfidAsiento);
							
			Reporte Reporte = new Reporte("AnticiposProveedores.rpt", "");
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Anticipo");
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();	
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void frmAnticiposProveedor_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un recibo de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}

