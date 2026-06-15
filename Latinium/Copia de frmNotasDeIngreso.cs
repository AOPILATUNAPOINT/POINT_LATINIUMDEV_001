using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmOrdenesDeCompra.
	/// </summary>
	public class frmNotasDeIngreso : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkConsignacion;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCredTributario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdProyecto;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		int IdOC = 0;
		DateTime hl;
		int iOrigen = 0;
		string SColumna="";

		public frmNotasDeIngreso()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public frmNotasDeIngreso(int idOC)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
		
			IdOC = idOC;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmNotasDeIngreso(int idOC, int IOrigen, DateTime Hl, string sColumna)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
		
			InitializeComponent();

			IdOC = idOC;
			iOrigen = IOrigen;
			hl=Hl ;
			SColumna= sColumna;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNotasDeIngreso));
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkConsignacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCredTributario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(708, 456);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 280;
			this.txtIdCompra.Visible = false;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(668, 360);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 26);
			this.lblUsuario.TabIndex = 279;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// chkConsignacion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkConsignacion.Appearance = appearance1;
			this.chkConsignacion.BackColor = System.Drawing.Color.Transparent;
			this.chkConsignacion.CausesValidation = false;
			this.chkConsignacion.Enabled = false;
			this.chkConsignacion.Location = new System.Drawing.Point(552, 40);
			this.chkConsignacion.Name = "chkConsignacion";
			this.chkConsignacion.Size = new System.Drawing.Size(96, 22);
			this.chkConsignacion.TabIndex = 276;
			this.chkConsignacion.Text = "Consignación";
			this.chkConsignacion.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(664, 37);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 275;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(664, 10);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 273;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(604, 414);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 272;
			this.label1.Text = "TOTAL";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 10);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(36, 16);
			this.lblIdentificacion.TabIndex = 269;
			this.lblIdentificacion.Text = "R.U.C.";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(200, 10);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(55, 16);
			this.lblCliente.TabIndex = 258;
			this.lblCliente.Text = "Proveedor";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 42);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 256;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(200, 42);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 257;
			this.lblBodega.Text = "Bodega";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRuc
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance2;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(64, 7);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 249;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
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
			this.cmbBodega.Location = new System.Drawing.Point(264, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(288, 21);
			this.cmbBodega.TabIndex = 267;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(264, 7);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(288, 22);
			this.txtNombre.TabIndex = 248;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance4;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(668, 408);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 28);
			this.txtTotal.TabIndex = 266;
			// 
			// txtNumero
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance5;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(728, 7);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 247;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance6;
			ultraGridBand2.AddButtonCaption = "_DetCompra";
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Artículo";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn6.Width = 194;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 160;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 380;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.MaskInput = "";
			ultraGridColumn9.NullText = "0";
			ultraGridColumn9.PromptChar = ' ';
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn10.PromptChar = ' ';
			ultraGridColumn10.Width = 95;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Width = 100;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8.5F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance13;
			appearance14.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance14.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance16;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 72);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(836, 264);
			this.ultraGrid1.TabIndex = 251;
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(bool);
			ultraDataColumn11.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance17;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 368);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(480, 66);
			this.txtNotas.TabIndex = 246;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 344);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(33, 16);
			this.label3.TabIndex = 283;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 440);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(852, 8);
			this.gbBotones.TabIndex = 356;
			this.gbBotones.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 456);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 345;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(83, 456);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 354;
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
			this.btnEditar.Location = new System.Drawing.Point(164, 456);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 353;
			this.btnEditar.Text = "&Editar";
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
			this.btnImprimir.Location = new System.Drawing.Point(243, 456);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 347;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(323, 456);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 346;
			this.btnGuardar.Text = "&Guardar";
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
			this.btnAnular.Location = new System.Drawing.Point(402, 456);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 355;
			this.btnAnular.Text = "&Anular";
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
			this.btnCancelar.Location = new System.Drawing.Point(481, 456);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 348;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ForeColor = System.Drawing.Color.Black;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(556, 456);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 349;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(724, 456);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 358;
			this.txtEstado.Visible = false;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(740, 456);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdCliente.TabIndex = 357;
			this.txtIdCliente.Visible = false;
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(772, 456);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdGrupoCliente.TabIndex = 362;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(756, 456);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoRuc.TabIndex = 361;
			this.txtIdTipoRuc.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(788, 456);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 360;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtContadoCredito
			// 
			this.txtContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContadoCredito.Enabled = false;
			this.txtContadoCredito.Location = new System.Drawing.Point(804, 456);
			this.txtContadoCredito.Name = "txtContadoCredito";
			this.txtContadoCredito.PromptChar = ' ';
			this.txtContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtContadoCredito.TabIndex = 363;
			this.txtContadoCredito.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(820, 456);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 364;
			this.txtIdComprobante.Visible = false;
			// 
			// txtIdCredTributario
			// 
			this.txtIdCredTributario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCredTributario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCredTributario.Enabled = false;
			this.txtIdCredTributario.Location = new System.Drawing.Point(836, 456);
			this.txtIdCredTributario.Name = "txtIdCredTributario";
			this.txtIdCredTributario.PromptChar = ' ';
			this.txtIdCredTributario.Size = new System.Drawing.Size(16, 22);
			this.txtIdCredTributario.TabIndex = 365;
			this.txtIdCredTributario.Visible = false;
			// 
			// dtFecha
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance18;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 40);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 370;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtIdProyecto
			// 
			this.txtIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdProyecto.Enabled = false;
			this.txtIdProyecto.Location = new System.Drawing.Point(692, 456);
			this.txtIdProyecto.Name = "txtIdProyecto";
			this.txtIdProyecto.PromptChar = ' ';
			this.txtIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtIdProyecto.TabIndex = 371;
			this.txtIdProyecto.Visible = false;
			// 
			// frmNotasDeIngreso
			// 
			this.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Appearance.Options.UseFont = true;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(852, 486);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtIdProyecto);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtIdCredTributario);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtContadoCredito);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.chkConsignacion);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtNotas);
			this.KeyPreview = true;
			this.Name = "frmNotasDeIngreso";
			this.Text = "Notas De Ingreso";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmOrdenesDeCompra_KeyDown);
			this.Load += new System.EventHandler(this.frmOrdenesDeCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private Acceso miAcceso;
		bool bRastreo = false;
		int IdSigno = 1;
		private C1DataRow drSeteoF;
		private bool bActualiza = true;
		decimal dPorcIva = 0.00m;
		int idBloqueaTransacciones = 0;
		int idProyecto = 0;
		bool bCarga;
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
		
		private void frmOrdenesDeCompra_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0629");

			bCarga=true;

//			if (!Funcion.Permiso("1187", cdsSeteoF))
//			{				
//				MessageBox.Show("No Puede Ingresar a Notas De Ingreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				BeginInvoke(new MethodInvoker(this.Close));				
//
//				return;
//			}
						
			Funcion.EjecutaSQL(cdsSeteoF, "Exec AuditaCrear 53, 6, 'NOTAS DE INGRESO'", true);

			bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");

			IdSigno = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 2");
				
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, "Exec ProyectoPredefinido 2, 3", true);
									
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
						
			bCarga = false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bNuevo = false;
			bEdicion = false;

			this.txtIdCompra.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdProyecto.Value = 0;
			this.txtIdGrupoCliente.Value = 0;
			this.txtIdTipoRuc.Value = 0;
			this.txtContadoCredito.Value = 0;
			this.txtIdComprobante.Value = 0;
			this.txtIdCredTributario.Value = 0;
			
			this.chkConsignacion.Checked = false;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.cmbBodega.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");			
			this.txtNumero.Text = "";			
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleNotaIngreso 0");

			this.txtNotas.Text = "";
			
			this.txtTotal.Value = 0.00m;

			bAnulado = false;
			bEliminaAlValidar = false;

			this.chkConsignacion.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.dtFecha.Enabled = false;			
			this.txtNumero.Enabled = false;
						
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			this.txtNotas.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;						
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Text = "Anular";

			idBloqueaTransacciones = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			this.txtEstado.Value = 1;						
			this.chkConsignacion.Checked = true;
						
			this.txtRuc.Enabled = true;						
			this.cmbBodega.Enabled = true;
			this.txtNotas.Enabled = true;

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.txtRuc.Focus();

			bNuevo = true;
			bEdicion = true;
		}
		
		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(true, 36);

				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();

					if (miAcceso.CambiarFecha) this.dtFecha.Focus();						
				}				
			}

			if (e.KeyCode == Keys.Enter)
				this.cmbBodega.Focus();
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(true, 36, this.txtRuc.Text.ToString(), cdsSeteoF);

			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.txtIdTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[2] != null) this.txtIdGrupoCliente.Value = sParametros[2];
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];
			if (sParametros[6] != null) this.txtContadoCredito.Value = sParametros[6];
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				CargaCliente();
	
				if ((int)this.txtIdCliente.Value > 0)
				{
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
															
					this.chkConsignacion.Enabled = true;										
				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!Validacion.vbClienteExiste(this.txtRuc, true, 2, 6, cdsSeteoF)) 
				{
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor '{0}'", this.txtRuc.Text.ToString())) == 0)
				{
					MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual Pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtRuc.Focus();
					e.Cancel = true;
					return;
				}
				
				if (this.chkConsignacion.Checked)
				{
					if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(AplicaConsignacion, 0) From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtRuc.Text.ToString()), true))
					{
						MessageBox.Show("El Proveedor no esta Asignado para Consignación. Revise la Ficha del Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtRuc.Focus();
						e.Cancel = true;
						return;
					}						
				}				
			}
		}
				
		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bEdicion)
			{
				if (ultraGrid1.ActiveCell != null) 
				{					
					FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);	

					#region Articulos
					if (e.KeyCode == Keys.F3)
					{		
						try
						{	
							if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO")
							{									
								int idDetCompra = 0;
								if (this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
								int idArticulo = 0;
								if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;

								string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
								int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
										
								if (!bNuevo && iConteo > 0)
								{
									MessageBox.Show(string.Format("Este Articulo tiene {0} Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
									return;
								}
									
								FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, (int)this.cmbBodega.Value, 36, 6, Convert.ToDateTime(this.dtFecha.Value), false);									
							}						
						}
						catch(System.Exception ex)
						{
							MessageBox.Show(ex.Message, "Error en Ingreso de Productos", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					#endregion Articulos
				}
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					if (!bNuevo)
					{	
						for (int i=0; i<e.Rows.Length; i++)
						{							
							int idDetCompra = 0;							
							int idArticulo = 0;

							if (e.Rows[i].Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Rows[i].Cells["idDetCompra"].Value;
							if (e.Rows[i].Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Rows[i].Cells["idArticulo"].Value;

							string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
							int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
							if (iConteo > 0)
							{
								MessageBox.Show(string.Format("Este Articulo ya tiene {0} Seriales Registradas no puede Eliminarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;																
							}
						}
					}
					else e.DisplayPromptMsg = false;
				}
				else e.DisplayPromptMsg = false;
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid1, 36, (int)this.cmbBodega.Value);
		}

		private void cmdTotal()
		{			
			#region Desplazamiento de grilla
			this.txtTotal.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value);

					this.txtTotal.Value = Convert.ToDecimal(this.txtTotal.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
				}
			} 
			#endregion Desplazamiento de grilla
		}

		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}
		
		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato;

			this.txtTotal.FormatString = stFormato;
			this.txtTotal.MaskInput = stInput;
			#endregion Decimales para el total
			
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);			
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion)
			{
				if (bActualiza)
				{
					bActualiza = false;
					#region Verifica Cambio
					#region Codigo o Articulo
					if (e.Cell.Column.ToString() == "Codigo")// || e.Cell.Column.ToString() == "Articulo")
					{
						bActualiza = false;						
						FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid1, 0, 0, (int)this.cmbBodega.Value, 0, 0, 0, 0, 0, Convert.ToDateTime(this.dtFecha.Value), 36, false, false, false, false, 0, 0, 0, 0, 0, false, this.chkConsignacion.Checked, 0, cdsSeteoF);
						bActualiza = true;
						this.cmdTotal();
					} 
					#endregion Codigo o articulo
						
					#region Cantidad
					if (e.Cell.Column.ToString() == "Cantidad")
					{
						if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((int)e.Cell.Row.Cells["Cantidad"].Value > 1)
							{
								e.Cell.Value = 0;
								MessageBox.Show("Seleccione un Producto antes de Ingresar Cantidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}					

						#region Valida Seriales ingresadas
						if (!bNuevo)
						{
							int idDetCompra = 0;
							if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
							int idArticulo = 0;
							if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

							string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
							int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

							int iCantidadM = (int)e.Cell.Row.Cells["Cantidad"].Value;
 
							if (iCantidadM < iConteo)
							{
								e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
								MessageBox.Show(string.Format("El Artículo '{0}' tiene {1} seriales registradas", e.Cell.Row.Cells["Codigo"].Value.ToString().Trim(), iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
							}								
						}
						#endregion Valida Seriales ingresadas
						this.cmdTotal();						
					}
					#endregion Cantidad
					
					#endregion Fin de Verificacion de Cambio
					
					#region Precio
					if (e.Cell.Column.ToString() == "Precio")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["Precio"].Value == System.DBNull.Value) e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;

						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["Precio"].Value > 0.00m)
							{								
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}		
						
						cmdTotal();
						bActualiza = true;
					}
					#endregion Precio										
				}
				bActualiza = true;
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{	
			this.ConsultaMaestro((int)this.txtIdCompra.Value);

			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede Anular una Nota De Ingreso PROCESADA \n\n Consulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.txtIdCompra.Value, cdsSeteoF)) return;
			
			if (Funcion.ValidaBloqueo((int)this.txtIdCompra.Value, 1, cdsSeteoF)) return;
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSerialesIngresadasParaAnular {0}, 36", (int)this.txtIdCompra.Value), true) > 0 && !miAcceso.BAnularConSeriales)
			{
				MessageBox.Show("No puede Anular Existen Seriales Ingresadas\n\nConsulte Con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return;
			}			
			
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de Anular Esta Nota De Ingreso '{0}'", this.lblEstado.Text.ToString()), "Point Technoligy", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				if (Facturacion.AnulacionCompraVenta ((int)this.txtIdCompra.Value, 36, cdsSeteoF))
					this.ConsultaMaestro((int)this.txtIdCompra.Value);							
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{	
			this.ConsultaMaestro((int)this.txtIdCompra.Value);
			
			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("No puede editar una Nota De Ingreso PROCESADA, Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
									
			if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.txtIdCompra.Value, cdsSeteoF)) return;
			
			if (Funcion.ValidaBloqueo((int)this.txtIdCompra.Value, 1, cdsSeteoF)) return;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			#region Controles
			this.txtRuc.Enabled = true;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;			
			 
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);

			this.chkConsignacion.Enabled = true;							
						
			this.txtNotas.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;			
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;
			#endregion Controles			
		}

		public void ConsultaMaestro(int IdCompra)
		{
			try
			{
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;			
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;							
				this.btnAnular.Enabled = false;				
				
				int IdBodegaC = 0;
				//c.idCompra, c.Consignacion, c.idCliente, cl.idTipoRuc, cl.idGrupoCliente, cl.Ruc, cl.Nombre, c.Bodega, c.Fecha, c.Usuario, c.Numero, c.Estado, c.Notas, c.Total, c.Borrar
				string sSQL = string.Format("Exec ConsultaMaestroNotaDeIngreso {0}", IdCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdCompra.Value = IdCompra;
					this.chkConsignacion.Checked = dr.GetBoolean(1);
					this.txtIdCliente.Value = dr.GetInt32(2);
					this.txtIdTipoRuc.Value = dr.GetInt32(3);
					this.txtIdGrupoCliente.Value = dr.GetInt32(4);
					this.txtRuc.Text = dr.GetString(5);
					this.txtNombre.Text = dr.GetString(6);
					IdBodegaC = dr.GetInt32(7);
					this.dtFecha.Value = dr.GetDateTime(8);															
					this.lblUsuario.Text = dr.GetString(9);
					this.txtNumero.Text = dr.GetString(10);					
					this.txtEstado.Value = dr.GetInt32(11);
					this.txtNotas.Text = dr.GetValue(12).ToString();					
					this.txtTotal.Value = dr.GetDecimal(13);
					bAnulado = dr.GetBoolean(14);
				}
				dr.Close();
				this.cmbBodega.Value= IdBodegaC;

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleNotaIngreso {0}", IdCompra));

				if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bodega"))
				{
					if (bAnulado) this.lblEstado.Text = "ANULADO";
					else
					{
						if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";				
						if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
					}

					if (!bAnulado)
					{
						if (miAcceso.Editar)
							if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;

						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

						if (miAcceso.Anular && !bAnulado)
						{
							if ((int)this.txtEstado.Value == 1) this.btnAnular.Enabled = true;							
						}						
					}
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
			
			using (frmBuscaTransacciones Busqueda = new frmBuscaTransacciones(36, 0))//(int)this.cmbBodega.Value))
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.ConsultaMaestro((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
		}

		private void Imprimir(int iTipo)
		{			
			string stFiltro = "{Compra.idCompra} = " + this.txtIdCompra.Value.ToString();
						
			using (Reporte miReporte = new Reporte("OrdenCompra.Rpt", stFiltro))
			{			
				if (iTipo == 0) miReporte.ShowDialog(); // Pantalla				
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias
			
			this.cmdTotal();
								
			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;
			
			if (!Validacion.vbClienteExiste(this.txtRuc, true, 2, 6, cdsSeteoF)) return;
							
			if (this.chkConsignacion.Checked && !Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AplicaConsignacion From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtRuc.Text.ToString()), true))
			{
				MessageBox.Show("El proveedor no esta asignado para Consignación. Revise la ficha de Cliente.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();						
				return;
			}

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Bodega", cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Bodega")) return;
						
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 36)) return;			
												
			if (DialogResult.Yes == MessageBox.Show("Antes de Guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n INFORMACIÓN DEL PROVEEDOR \n PRODUCTOS \n PRECIOS CANTIDADES \n\n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						#region Numeracion
						if (bNuevo)
						{
							string sSQLNumero = "Exec FacturaNumeroEstacion 36";
							oCmdActualiza.CommandText = sSQLNumero;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
						}
						#endregion Numeracion

						#region Maestro
						string sSQLMaestro = string.Format("Exec GrabaMaestroNotaDeIngreso {0}, {1}, {2}, {3}, '{4}', '{5}', {6}, '{7}', {8}, 0", 
							(int)this.txtIdCompra.Value, this.chkConsignacion.Checked, (int)this.txtIdCliente.Value, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtNumero.Text.ToString(), 
							(int)this.txtEstado.Value, this.txtNotas.Text.ToString(), Convert.ToDecimal(this.txtTotal.Value));
						oCmdActualiza.CommandText = sSQLMaestro;
						this.txtIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();
						#endregion Maestro

						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							string sSQLDetalle = string.Format("Exec GuardaDetalleNotaDeIngreso {0}, {1}, {2}, {3}, {4}, {5}",
								(int)dr.Cells["idDetCompra"].Value, (int)this.txtIdCompra.Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Subtotal"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
												 
						string sSQLActDetalle = string.Format("Exec ConsultaDetalleNotaIngreso {0}", (int)this.txtIdCompra.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);
						this.ultraGrid1.DataSource = oDTDetalle;											
						#endregion Detalle	
												
						if (!bNuevo)
						{
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
							oCmdActualiza.CommandText = sSQLBloqueo;
							oCmdActualiza.ExecuteNonQuery();
						}

						oTransaction.Commit();
						
						if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";				
						if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
												
						#region Controles
						bNuevo = false;
						bEdicion = false;
						bEliminaAlValidar = false;

						this.chkConsignacion.Enabled = false;
						this.txtRuc.Enabled = false;
						this.dtFecha.Enabled = false;
						this.cmbBodega.Enabled = false;
														
						FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
																	
						this.txtNotas.Enabled = false;
						
						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;						
						this.btnGuardar.Enabled = false;												
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;																																																
						#endregion Controles						
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.cmdTotal();
		}

		private void frmOrdenesDeCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);
				if (e.KeyCode == Keys.X) if (miAcceso.Anular && bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);		
				if (e.KeyCode == Keys.F5) 
				{
					miAcceso = new Acceso(cdsSeteoF, "0602");
					this.ConsultaMaestro((int)this.txtIdCompra.Value);
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

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Bodega", cdsSeteoF)) e.Cancel = true;			
		}	

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{	
				int idDetCompra = 0;
				if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				int idArticulo = 0;
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

				string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
				int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
										
				if (!bNuevo && iConteo > 0)
				{
					MessageBox.Show(string.Format("Este Articulo tiene {0} Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
					e.Cancel = true;
				}
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 36, (DateTime)this.dtFecha.Value, this.chkConsignacion.Checked, cdsSeteoF);				
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{			
			if (this.cmbBodega.ActiveRow != null) 
				this.txtIdProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

