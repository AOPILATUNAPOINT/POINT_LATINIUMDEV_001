using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNotaDeDebitoCliente.
	/// </summary>
	public class frmDevolucionAnticipo : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.GroupBox groupBox2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAnticipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtId;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAnticipo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCuenta;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.ComponentModel.IContainer components;

		public frmDevolucionAnticipo()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int id = 0;

		public frmDevolucionAnticipo(int Id)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			
			id = Id;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDevolucionAnticipo));
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Manual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SerieL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Signo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			this.label1 = new System.Windows.Forms.Label();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtId = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtAnticipo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCuenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAnular = new System.Windows.Forms.Button();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtId)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 123);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(26, 16);
			this.label1.TabIndex = 265;
			this.label1.Text = "Tipo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMotivos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivos.Appearance = appearance1;
			this.cmbMotivos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMotivos.Enabled = false;
			this.cmbMotivos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(64, 120);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(304, 21);
			this.cmbMotivos.TabIndex = 266;
			this.cmbMotivos.ValueMember = "idMotivo";
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(600, 280);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 263;
			this.txtIdAsiento.Visible = false;
			// 
			// txtId
			// 
			this.txtId.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtId.Enabled = false;
			this.txtId.Location = new System.Drawing.Point(576, 280);
			this.txtId.Name = "txtId";
			this.txtId.PromptChar = ' ';
			this.txtId.Size = new System.Drawing.Size(16, 22);
			this.txtId.TabIndex = 262;
			this.txtId.Visible = false;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(8, 288);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(31, 16);
			this.label3.TabIndex = 261;
			this.label3.Text = "Valor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance2;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(64, 288);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(104, 22);
			this.txtValor.TabIndex = 259;
			this.txtValor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValor_KeyDown);
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Validating += new System.ComponentModel.CancelEventHandler(this.txtValor_Validating);
			this.txtValor.ValueChanged += new System.EventHandler(this.txtValor_ValueChanged);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(288, 288);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 254;
			// 
			// txtAnticipo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnticipo.Appearance = appearance3;
			this.txtAnticipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnticipo.Enabled = false;
			this.txtAnticipo.Location = new System.Drawing.Point(72, 8);
			this.txtAnticipo.Name = "txtAnticipo";
			this.txtAnticipo.Size = new System.Drawing.Size(112, 22);
			this.txtAnticipo.TabIndex = 248;
			this.txtAnticipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAnticipo_KeyDown);
			this.txtAnticipo.Validating += new System.ComponentModel.CancelEventHandler(this.txtAnticipo_Validating);
			this.txtAnticipo.ValueChanged += new System.EventHandler(this.txtAnticipo_ValueChanged);
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(8, 11);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(44, 16);
			this.lblComprobante.TabIndex = 247;
			this.lblComprobante.Text = "Anticipo";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 184);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 16);
			this.label2.TabIndex = 246;
			this.label2.Text = "Concepto";
			// 
			// txtConcepto
			// 
			this.txtConcepto.AcceptsReturn = true;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConcepto.Appearance = appearance4;
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(64, 184);
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(632, 80);
			this.txtConcepto.TabIndex = 245;
			this.txtConcepto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConcepto_KeyDown);
			this.txtConcepto.Validating += new System.ComponentModel.CancelEventHandler(this.txtConcepto_Validating);
			this.txtConcepto.ValueChanged += new System.EventHandler(this.txtConcepto_ValueChanged);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(8, 91);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(38, 16);
			this.lblCliente.TabIndex = 242;
			this.lblCliente.Text = "Cliente";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 59);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 241;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance5;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(64, 88);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(136, 22);
			this.txtRuc.TabIndex = 239;
			// 
			// dtFecha
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance6;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 56);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 240;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			// 
			// txtNombre
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance7;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(208, 88);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(368, 22);
			this.txtNombre.TabIndex = 238;
			// 
			// txtNumero
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance8;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(464, 56);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 235;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(192, 59);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 233;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(248, 56);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 234;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(376, 59);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 231;
			this.label4.Text = "Número";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.ImageList = this.imageList1;
			this.btnAsiento.Location = new System.Drawing.Point(407, 336);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 272;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
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
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 336);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 267;
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
			this.btnConsultar.Location = new System.Drawing.Point(84, 336);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 268;
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
			this.btnEditar.Location = new System.Drawing.Point(166, 336);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 271;
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
			this.btnImprimir.Location = new System.Drawing.Point(246, 336);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 270;
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
			this.btnGuardar.Location = new System.Drawing.Point(327, 336);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 269;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(575, 336);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 273;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(651, 336);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 274;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdAnticipo
			// 
			this.txtIdAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAnticipo.Enabled = false;
			this.txtIdAnticipo.Location = new System.Drawing.Point(624, 280);
			this.txtIdAnticipo.Name = "txtIdAnticipo";
			this.txtIdAnticipo.PromptChar = ' ';
			this.txtIdAnticipo.Size = new System.Drawing.Size(16, 22);
			this.txtIdAnticipo.TabIndex = 275;
			this.txtIdAnticipo.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Double);
			ultraDataColumn9.DataType = typeof(System.Double);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(System.Double);
			ultraDataColumn10.DefaultValue = 0;
			ultraDataColumn11.DataType = typeof(System.Double);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn19.DefaultValue = false;
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox1.Location = new System.Drawing.Point(0, 320);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(736, 8);
			this.groupBox1.TabIndex = 277;
			this.groupBox1.TabStop = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 155);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 16);
			this.label5.TabIndex = 278;
			this.label5.Text = "Cuenta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(736, 8);
			this.groupBox2.TabIndex = 279;
			this.groupBox2.TabStop = false;
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
			// txtCuenta
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuenta.Appearance = appearance9;
			this.txtCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuenta.Enabled = false;
			this.txtCuenta.Location = new System.Drawing.Point(384, 120);
			this.txtCuenta.Name = "txtCuenta";
			this.txtCuenta.Size = new System.Drawing.Size(304, 22);
			this.txtCuenta.TabIndex = 280;
			this.txtCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuenta_KeyDown);
			this.txtCuenta.Validating += new System.ComponentModel.CancelEventHandler(this.txtCuenta_Validating);
			this.txtCuenta.ValueChanged += new System.EventHandler(this.txtCuenta_ValueChanged);
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(552, 280);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 281;
			this.txtEstado.Visible = false;
			// 
			// txtIdCuenta
			// 
			this.txtIdCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCuenta.Enabled = false;
			this.txtIdCuenta.Location = new System.Drawing.Point(528, 280);
			this.txtIdCuenta.Name = "txtIdCuenta";
			this.txtIdCuenta.PromptChar = ' ';
			this.txtIdCuenta.Size = new System.Drawing.Size(16, 22);
			this.txtIdCuenta.TabIndex = 282;
			this.txtIdCuenta.Visible = false;
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.ImageIndex = 55;
			this.btnAnular.ImageList = this.imageList1;
			this.btnAnular.Location = new System.Drawing.Point(491, 336);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(80, 23);
			this.btnAnular.TabIndex = 283;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// cmbCuenta
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuenta.Appearance = appearance10;
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "Cuenta";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 210;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 90;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(64, 152);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(304, 21);
			this.cmbCuenta.TabIndex = 284;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn22.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// frmDevolucionAnticipo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(736, 366);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.txtIdCuenta);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtCuenta);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIdAnticipo);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbMotivos);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtId);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtAnticipo);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbBodega);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDevolucionAnticipo";
			this.Text = "Devolución Anticipo";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDevolucionAnticipo_Closing);
			this.Load += new System.EventHandler(this.frmDevolucionAnticipo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtId)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		bool bNuevo = false;
		bool bEdicion = false;
		int idBloqueaTransacciones = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmDevolucionAnticipo_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("553", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Devolución De Anticipos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Codigo From Bodega Where Activo = 1 And Factura = 1 Order By Bodega");
				this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 27");
				this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "EXEC [ConsultaCuentaDevoAnticipos] ");

				this.dtFecha.CalendarInfo.MaxDate = DateTime.Now;

				if (id > 0) this.Consulta(id);
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 11", idBloqueaTransacciones, (int)this.txtId.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);				
			}

			bNuevo = false;
			bEdicion = false;

			this.txtId.Value = 0;
			this.txtIdAnticipo.Value = 0;
			this.cmbBodega.Enabled = false;
			this.txtIdAsiento.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdCuenta.Value = 0;

			this.txtAnticipo.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.cmbMotivos.Value = System.DBNull.Value;
			this.cmbCuenta.Value = System.DBNull.Value;

			this.txtConcepto.Text = "";
			this.txtValor.Value = 0.00m;

			this.txtAnticipo.Enabled = false;
			this.dtFecha.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.txtConcepto.Enabled = false;
			this.txtValor.Enabled = false;
			this.cmbMotivos.Value = false;
			this.cmbBodega.Value = false;

			this.lblEstado.Text = "";

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;

			idBloqueaTransacciones = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtAnticipo.Enabled = true;
			//this.dtFecha.Enabled = true;
			this.txtConcepto.Enabled = true;
			this.txtValor.Enabled = true;
			this.cmbCuenta.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;						
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
			this.dtFecha.Value = DateTime.Today;		

			this.txtAnticipo.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{	
			this.Consulta((int)this.txtId.Value);

			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From DevolucionAnticipos Where idDevolucionAnticipos = {0}", (int)this.txtId.Value));

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede EDITAR este Documento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}
			if (iEstado == 1 && (int)this.txtIdCuenta.Value != 4)
			{
				MessageBox.Show("El Documento ya fue Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if (iEstado == 2)
			{
				MessageBox.Show("El Registro Esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtId.Value, 56, cdsSeteoF)) return;

			if (Funcion.ValidaBloqueo((int)this.txtId.Value, 11, cdsSeteoF)) return;
			
			//this.dtFecha.Enabled = true;
			this.cmbBodega.Enabled = false;
			this.cmbCuenta.Enabled = true;
			this.txtConcepto.Enabled = true;
			this.txtValor.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;						
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;			

			this.dtFecha.Value = DateTime.Today;
			this.cmbMotivos.Value = 12;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtId.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			bNuevo = false;
			bEdicion = true;

			this.txtAnticipo.Focus();			
		}

		private void txtAnticipo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.txtAnticipo.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Número de Anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAnticipo.Focus();
				e.Cancel = true;				
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;				
			}
		}

		private void txtCuenta_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.txtCuenta.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Número de Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCuenta.Focus();
				e.Cancel = true;				
			}
		}

		private void txtConcepto_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.txtConcepto.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Concepto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtConcepto.Focus();
				e.Cancel = true;				
			}
		}

		private void txtValor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (this.txtValor.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
			else if (Convert.ToDecimal(this.txtValor.Value) == 0.00m)
			{
				MessageBox.Show("El Valor debe ser mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
			else 
			{
				#region Saldo
				string sSQLSaldo = string.Format("Exec ActualizaSaldoAnticipos {0}, 1", (int)this.txtIdAnticipo.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo, true);

				decimal dSaldo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Saldo, 0) From Anticipos Where idAnticipo = {0}", (int)this.txtIdAnticipo.Value));

				if (Convert.ToDecimal(this.txtValor.Value) > dSaldo)
				{
					MessageBox.Show(string.Format("El Valor a Devolver no puede ser Mayor al Saldo del Anticipo {0}", dSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
				#endregion Saldo
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.txtAnticipo.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Número de Anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtAnticipo.Focus();				
				return;
			}
		
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();				
				return;
			}
		
//			if (this.txtCuenta.Text.ToString().Length == 0)
//			{
//				MessageBox.Show("Ingrese el Número de Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.txtCuenta.Focus();				
//				return;
//			}
		
			if (this.cmbCuenta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Cuenta Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuenta.Focus();				
				return;
			}
			if ((int)this.cmbCuenta.Value == 0)
			{
				MessageBox.Show("Seleccione la Cuenta Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCuenta.Focus();				
				return;
			}

			if (this.txtConcepto.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Concepto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtConcepto.Focus();					
				return;
			}
		
			if (this.txtValor.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtValor.Focus();
			}
			else if (Convert.ToDecimal(this.txtValor.Value) == 0.00m)
			{
				MessageBox.Show("El Valor debe ser mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtValor.Focus();
			}

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
			{
				DateTime dtFechaVal = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede EDITAR este Documento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFechaVal.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}
			#endregion Validacion

			#region Numeracion
			int iCont = 0;
			if (bNuevo) 
			{
				iCont = Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From DevolucionAnticipos");
				iCont ++;
				char Pad = '0';
				this.txtNumero.Text = "DA-" + iCont.ToString().PadLeft(6, Pad);
			}
			#endregion Numeracion

//			if ((int)this.txtIdCuenta.Value == 4) this.txtEstado.Value = 1; else this.txtEstado.Value = 0;
      this.txtEstado.Value = 1;
			#region Guardar
			DateTime dtFecha = (DateTime)this.dtFecha.Value;			
			string sSQL = string.Format("Exec GuardaDevolucionAnticipos {0}, {1}, {2}, '{3}', {4}, '{5}', {6}, {7}, '{8}', {9}",
				(int)this.txtId.Value, this.txtIdAnticipo.Value, this.txtEstado.Value, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), 
				(int)this.cmbMotivos.Value, (int)this.cmbCuenta.Value, this.txtConcepto.Text.ToString(), Convert.ToDecimal(this.txtValor.Value));
			this.txtId.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			#endregion Guardar

			#region Guarda Detalle Forma Pago
//			if ((int)this.txtIdCuenta.Value == 4)
//			{
//				if (!bNuevo) Funcion.EjecutaSQL(cdsSeteoF, string.Format("If Exists(Select * From AutorizacionDePagos Where idOrigen = {0} And idTipo = 16) Delete From AutorizacionDePagos Where idOrigen = {0} And idTipo = 16", 
//											 (int)this.txtId.Value));
//
//				int iDFP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDFPNotaDebito {0}", (int)this.txtId.Value));
//				string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, 5, 1, {2}, 0, 0, '', '', '', '', '', 0, '', '', 0, 0, 0, 0", iDFP, (int)this.txtId.Value, Convert.ToDecimal(this.txtValor.Value));							
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
//			}
			#endregion Guarda Detalle Forma Pago

			#region Autorizacion de Pagos
//			if ((int)this.txtIdCuenta.Value != 4)
//			{
//				if (!bNuevo) Funcion.EjecutaSQL(cdsSeteoF, string.Format("If Exists (Select idDetFormaPago From DetalleFormaPago Where Origen = 5 And idCompra = {0}) Delete From DetalleFormaPago Where Origen = 5 And idCompra = {0}", 
//											 (int)this.txtId.Value));

				string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 16, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'DEVOLUCION ANTICIPOS', {4}", 
					(int)this.txtId.Value, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"), (decimal )this.txtValor.Value, (int)this.cmbCuenta.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAP);
//			}
			#endregion Autorizacion de Pagos

			#region Asiento
//			string sSQLAsiento = string.Format("Exec AsientoDevolucionAnticipos {0}, {1}", (int)this.txtIdAsiento.Value, (int)this.txtId.Value);	
//			this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
			#endregion Asiento
			
			#region Saldo
			string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 6", (int)this.txtIdAnticipo.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo, true);
			#endregion Saldo

			if (!bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 11", idBloqueaTransacciones, (int)this.txtId.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);				
			}

			if ((int)this.txtEstado.Value == 0) this.lblEstado.Text = "PENDIENTE"; else if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "CANCELADO"; 

			#region Controles
			this.txtAnticipo.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.dtFecha.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.txtConcepto.Enabled = false;
			this.txtValor.Enabled = false;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnImprimir.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnAsiento.Enabled = true;
			this.btnAnular.Enabled = false;
			#endregion Controles

			bNuevo = false;
			bEdicion = false;
		}

		private void txtAnticipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();

			if (e.KeyCode == Keys.F3)
			{
				using (frmBuscaAnticiposND miBusqueda = new frmBuscaAnticiposND())
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						try
						{						
							this.txtAnticipo.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Secuencial"].Value.ToString();
							this.txtIdAnticipo.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idAnticipo"].Value;
							this.cmbBodega.Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["Bodega"].Value;
							this.txtRuc.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Ruc"].Value.ToString();
							this.txtNombre.Text = miBusqueda.ultraGrid1.ActiveRow.Cells["Nombre"].Value.ToString();
							this.txtValor.Value = (decimal)miBusqueda.ultraGrid1.ActiveRow.Cells["Valor"].Value;

							this.cmbMotivos.Value = 12;
							this.txtIdCuenta.Value = 4;
//							this.txtCuenta.Text = "CAJA";

							this.dtFecha.Focus();
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void txtCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConcepto.Focus();

			if (e.KeyCode == Keys.F3)
			{
				using (frmBuscaCuentas Busqueda = new frmBuscaCuentas())
				{
					if (DialogResult.OK == Busqueda.ShowDialog())
					{
						this.txtIdCuenta.Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
						this.txtCuenta.Text = Busqueda.ultraGrid1.ActiveRow.Cells["Descripcion"].Value.ToString();
					}
				}
			}
		}

		private void txtConcepto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtValor.Focus();
		}

		private void txtValor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void Consulta (int Id)
		{
			/*idDevolucionAnticipos, da.idAnticipo, da.idAsiento, da.Estado, da.Fecha, da.Bodega, da.Numero, a.Secuencial,
				cl.Ruc, cl.Nombre, idMotivo, da.idCuenta, c.Descripcion, da.Concepto, da.Valor,
				Case da.Estado When 0 Then 'CANCELADO' When 1 Then 'ANULADO' End As EstadoM*/
			try
			{
				string sSQL = string.Format("Exec ConsultaIndividualDevolucionAnticipos {0}", Id);
				SqlDataReader drDA = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drDA.Read();
				if (drDA.HasRows)
				{
					this.txtId.Value = Id;
					this.txtIdAnticipo.Value = drDA.GetInt32(1);
					this.txtIdAsiento.Value = drDA.GetInt32(2);
					this.txtEstado.Value = drDA.GetInt32(3);
					this.dtFecha.Value = drDA.GetDateTime(4);
					this.cmbBodega.Value = drDA.GetInt32(5);
					this.txtNumero.Text = drDA.GetString(6);
					this.txtAnticipo.Text = drDA.GetString(7);					
					this.txtRuc.Text = drDA.GetString(8);
					this.txtNombre.Text = drDA.GetString(9);
					this.cmbMotivos.Value = drDA.GetInt32(10);
					this.cmbCuenta.Value = drDA.GetInt32(11);
//					this.txtCuenta.Text = drDA.GetString(12);
					this.txtConcepto.Text = drDA.GetString(13);
					this.txtValor.Value = drDA.GetDecimal(14);
					this.lblEstado.Text = drDA.GetString(15);
				}
				drDA.Close();

				if ((int)this.txtEstado.Value != 2)
				{
					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
					{						
						//if ((int)this.txtIdCuenta.Value == 4) this.btnEditar.Enabled = true; else if ((int)this.txtEstado.Value == 0) this.btnEditar.Enabled = true;
								
						this.btnEditar.Enabled = false;
						this.btnAsiento.Enabled = true;
						this.btnAnular.Enabled = false; 				
					}

					this.btnImprimir.Enabled = true;
				}

				this.btnCancelar.Enabled = true;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaDA Busqueda = new frmBuscaDA())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					try
					{							
						this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idDevolucionAnticipos"].Value);
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
		}

		private void txtCuenta_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAnticipo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value > 0)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Desea Sobreescribir el Asiento?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					#region Asiento
					string sSQLAsiento = string.Format("Exec AsientoDevolucionAnticipos {0}, {1}", (int)this.txtIdAsiento.Value, (int)this.txtId.Value);	
					this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
					#endregion Asiento
				}
			}
			else
			{
				#region Asiento
				string sSQLAsiento = string.Format("Exec AsientoDevolucionAnticipos {0}, {1}", (int)this.txtIdAsiento.Value, (int)this.txtId.Value);	
				this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
				#endregion Asiento
			}
		}

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right) 
			{						
				if ((int)this.txtIdAsiento.Value == 0)
				{
					MessageBox.Show("No existe un Asiento creado de esta Transacción");				
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

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From DevolucionAnticipos Where idDevolucionAnticipos = {0}", (int)this.txtId.Value));

			if (iEstado == 2)
			{				
				MessageBox.Show("No Puede Imprimir este Documento fue Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				this.Consulta((int)this.txtEstado.Value);
			}
			else
			{
				this.Cursor = Cursors.WaitCursor;
			
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldBodega = new ParameterField ();
				ParameterDiscreteValue discreteValBodega = new ParameterDiscreteValue ();
				paramFieldBodega.ParameterFieldName = "@idDA";
				discreteValBodega.Value = (int)this.txtId.Value; 
				paramFieldBodega.CurrentValues.Add (discreteValBodega);
				paramFields.Add (paramFieldBodega);

				Reporte miRep = new Reporte("DevolucionAnticipos.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Devolución Anticipos");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
		
				this.Cursor = Cursors.Default;
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtId.Value);

			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From DevolucionAnticipos Where idDevolucionAnticipos = {0}", (int)this.txtId.Value));
			if (iEstado == 2)
			{
				MessageBox.Show("El Registro ya esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Pagos"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No Puede ANULAR este Documento, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return;
			}

			if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtId.Value, 56, cdsSeteoF)) return;

			if (Funcion.ValidaBloqueo((int)this.txtId.Value, 11, cdsSeteoF)) return;
									
			if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Anular este Documento", "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				string sSQL = string.Format("Exec AnulaDevolucionAnticipos {0}", (int)this.txtId.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				#region Saldo
				string sSQLSaldo = string.Format("Exec ActualizaSaldoAnticipos {0}, 1", (int)this.txtIdAnticipo.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo, true);
				#endregion Saldo

				this.txtEstado.Value = 2;
				this.lblEstado.Text = "ANULADO";
				this.btnAnular.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnEditar.Enabled = false;
			}			
		}

		private void txtValor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtConcepto_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmDevolucionAnticipo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
			}
		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConcepto.Focus();
		}
	}
}

