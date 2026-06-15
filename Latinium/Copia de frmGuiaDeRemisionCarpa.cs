using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGuiaDeRemisionCarpa.
	/// </summary>
	public class frmGuiaDeRemisionCarpa : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label7;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigen;
		private System.Windows.Forms.Label label18;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGuiaRemision;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTraslado;
		private System.Windows.Forms.GroupBox gbBotones;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label17;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label16;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtPlaca;
		private System.Windows.Forms.Label label14;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRucDestinatario;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRazonSocialDestinatario;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDestino;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoFactura;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFinTraslado;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIniTraslado;
		private System.Windows.Forms.Label lblNumero;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridArticulos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.Windows.Forms.Label lblUsuario;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmGuiaDeRemisionCarpa()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		bool bNuevo = false;
		bool bEdicion = false;
		int iBodega = 0;
		bool bImpreso = false;
		bool bAutomatico = false;
		private Acceso miAcceso;

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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivoTraslado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGuiaDeRemisionCarpa));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleGuiaDeRemision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMaestroGuiaRemision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleGuiaDeRemision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMaestroGuiaRemision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivoTraslado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			this.label7 = new System.Windows.Forms.Label();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGuiaRemision = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTraslado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtPlaca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label13 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtRucDestinatario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRazonSocialDestinatario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDestino = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbTipoFactura = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFinTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtIniTraslado = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.uGridArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGuiaRemision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucDestinatario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRazonSocialDestinatario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			this.SuspendLayout();
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(8, 136);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 16);
			this.label7.TabIndex = 331;
			this.label7.Text = "Origen";
			// 
			// txtOrigen
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrigen.Appearance = appearance1;
			this.txtOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigen.Enabled = false;
			this.txtOrigen.Location = new System.Drawing.Point(56, 136);
			this.txtOrigen.Multiline = true;
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.Size = new System.Drawing.Size(264, 22);
			this.txtOrigen.TabIndex = 330;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 8);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(30, 16);
			this.label18.TabIndex = 329;
			this.label18.Text = "Local";
			// 
			// cmbBodega
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance2;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 352;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
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
			this.cmbBodega.Location = new System.Drawing.Point(96, 8);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 21);
			this.cmbBodega.TabIndex = 328;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(504, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 327;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(592, 168);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 326;
			this.txtEstado.Visible = false;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(616, 168);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 325;
			this.txtIdCompra.Visible = false;
			// 
			// txtIdGuiaRemision
			// 
			this.txtIdGuiaRemision.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGuiaRemision.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGuiaRemision.Enabled = false;
			this.txtIdGuiaRemision.Location = new System.Drawing.Point(640, 168);
			this.txtIdGuiaRemision.Name = "txtIdGuiaRemision";
			this.txtIdGuiaRemision.PromptChar = ' ';
			this.txtIdGuiaRemision.Size = new System.Drawing.Size(16, 22);
			this.txtIdGuiaRemision.TabIndex = 324;
			this.txtIdGuiaRemision.Visible = false;
			// 
			// cmbTraslado
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTraslado.Appearance = appearance3;
			this.cmbTraslado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTraslado.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 352;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTraslado.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTraslado.DisplayMember = "Motivo";
			this.cmbTraslado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTraslado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTraslado.Enabled = false;
			this.cmbTraslado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTraslado.Location = new System.Drawing.Point(392, 104);
			this.cmbTraslado.Name = "cmbTraslado";
			this.cmbTraslado.Size = new System.Drawing.Size(264, 21);
			this.cmbTraslado.TabIndex = 323;
			this.cmbTraslado.ValueMember = "idMotivoTraslado";
			// 
			// gbBotones
			// 
			this.gbBotones.Location = new System.Drawing.Point(0, 384);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(672, 8);
			this.gbBotones.TabIndex = 322;
			this.gbBotones.TabStop = false;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(322, 616);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 301;
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
			this.btnGuardar.Location = new System.Drawing.Point(243, 616);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 300;
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
			this.btnAnular.Location = new System.Drawing.Point(402, 616);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 302;
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
			this.btnCancelar.Location = new System.Drawing.Point(481, 616);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 303;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label17.Location = new System.Drawing.Point(8, 288);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(103, 17);
			this.label17.TabIndex = 321;
			this.label17.Text = "OBSERVACIONES";
			// 
			// txtNotas
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance4;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 312);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(576, 64);
			this.txtNotas.TabIndex = 299;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(408, 256);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(30, 16);
			this.label16.TabIndex = 320;
			this.label16.Text = "Placa";
			// 
			// txtPlaca
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPlaca.Appearance = appearance5;
			this.txtPlaca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPlaca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPlaca.Enabled = false;
			this.txtPlaca.Location = new System.Drawing.Point(480, 256);
			this.txtPlaca.MaxLength = 7;
			this.txtPlaca.Name = "txtPlaca";
			this.txtPlaca.Size = new System.Drawing.Size(104, 22);
			this.txtPlaca.TabIndex = 298;
			this.txtPlaca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPlaca_KeyDown);
			this.txtPlaca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPlaca_KeyPress);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(8, 256);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 16);
			this.label14.TabIndex = 318;
			this.label14.Text = "Nombre";
			// 
			// cmbResponsable
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance6;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 376;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(88, 256);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(304, 21);
			this.cmbResponsable.TabIndex = 296;
			this.cmbResponsable.ValueMember = "idPersonal";
			this.cmbResponsable.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbResponsable_KeyDown);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label13.Location = new System.Drawing.Point(8, 232);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(230, 17);
			this.label13.TabIndex = 317;
			this.label13.Text = "IDENTIFICACIÓN DEL TRANSPORTISTA";
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(408, 200);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(58, 16);
			this.label11.TabIndex = 316;
			this.label11.Text = "R.U.C./C.I.";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 200);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(67, 16);
			this.label12.TabIndex = 315;
			this.label12.Text = "Razón Social";
			// 
			// txtRucDestinatario
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRucDestinatario.Appearance = appearance7;
			this.txtRucDestinatario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRucDestinatario.Enabled = false;
			this.txtRucDestinatario.Location = new System.Drawing.Point(480, 200);
			this.txtRucDestinatario.Name = "txtRucDestinatario";
			this.txtRucDestinatario.Size = new System.Drawing.Size(104, 22);
			this.txtRucDestinatario.TabIndex = 295;
			// 
			// txtRazonSocialDestinatario
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRazonSocialDestinatario.Appearance = appearance8;
			this.txtRazonSocialDestinatario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRazonSocialDestinatario.Enabled = false;
			this.txtRazonSocialDestinatario.Location = new System.Drawing.Point(88, 200);
			this.txtRazonSocialDestinatario.Name = "txtRazonSocialDestinatario";
			this.txtRazonSocialDestinatario.Size = new System.Drawing.Size(304, 22);
			this.txtRazonSocialDestinatario.TabIndex = 294;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label10.Location = new System.Drawing.Point(8, 176);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(222, 17);
			this.label10.TabIndex = 314;
			this.label10.Text = "IDENTIFICACIÓN DEL DESTINATARIO";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(336, 136);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(42, 16);
			this.label9.TabIndex = 313;
			this.label9.Text = "Destino";
			// 
			// txtDestino
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDestino.Appearance = appearance9;
			this.txtDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDestino.Enabled = false;
			this.txtDestino.Location = new System.Drawing.Point(392, 136);
			this.txtDestino.Multiline = true;
			this.txtDestino.Name = "txtDestino";
			this.txtDestino.Size = new System.Drawing.Size(264, 22);
			this.txtDestino.TabIndex = 293;
			this.txtDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDestino_KeyDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(272, 104);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(102, 16);
			this.label5.TabIndex = 312;
			this.label5.Text = "Motivo del Traslado";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 104);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(26, 16);
			this.label8.TabIndex = 308;
			this.label8.Text = "Tipo";
			// 
			// cmbTipoFactura
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoFactura.Appearance = appearance10;
			this.cmbTipoFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoFactura.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbTipoFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoFactura.DisplayMember = "Nombre";
			this.cmbTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoFactura.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoFactura.Enabled = false;
			this.cmbTipoFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoFactura.Location = new System.Drawing.Point(56, 104);
			this.cmbTipoFactura.Name = "cmbTipoFactura";
			this.cmbTipoFactura.Size = new System.Drawing.Size(200, 21);
			this.cmbTipoFactura.TabIndex = 289;
			this.cmbTipoFactura.ValueMember = "idTipoFactura";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(196, 17);
			this.label2.TabIndex = 307;
			this.label2.Text = "DATOS COMPROBANTE DE VENTA";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(272, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(91, 16);
			this.label1.TabIndex = 306;
			this.label1.Text = "Termina Traslado";
			// 
			// dtFinTraslado
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFinTraslado.Appearance = appearance11;
			this.dtFinTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFinTraslado.DateButtons.Add(dateButton1);
			this.dtFinTraslado.Enabled = false;
			this.dtFinTraslado.Format = "dd/MM/yyyy";
			this.dtFinTraslado.Location = new System.Drawing.Point(376, 40);
			this.dtFinTraslado.Name = "dtFinTraslado";
			this.dtFinTraslado.NonAutoSizeHeight = 23;
			this.dtFinTraslado.Size = new System.Drawing.Size(112, 21);
			this.dtFinTraslado.SpinButtonsVisible = true;
			this.dtFinTraslado.TabIndex = 287;
			this.dtFinTraslado.Value = ((object)(resources.GetObject("dtFinTraslado.Value")));
			this.dtFinTraslado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFinTraslado_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(77, 16);
			this.label6.TabIndex = 305;
			this.label6.Text = "Inicio Traslado";
			// 
			// dtIniTraslado
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtIniTraslado.Appearance = appearance12;
			this.dtIniTraslado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtIniTraslado.DateButtons.Add(dateButton2);
			this.dtIniTraslado.Enabled = false;
			this.dtIniTraslado.Format = "dd/MM/yyyy";
			this.dtIniTraslado.Location = new System.Drawing.Point(96, 40);
			this.dtIniTraslado.Name = "dtIniTraslado";
			this.dtIniTraslado.NonAutoSizeHeight = 23;
			this.dtIniTraslado.Size = new System.Drawing.Size(112, 21);
			this.dtIniTraslado.SpinButtonsVisible = true;
			this.dtIniTraslado.TabIndex = 286;
			this.dtIniTraslado.Value = ((object)(resources.GetObject("dtIniTraslado.Value")));
			this.dtIniTraslado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtIniTraslado_KeyDown);
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(488, 8);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 304;
			this.lblNumero.Text = "Número";
			// 
			// txtNumero
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance13;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(552, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(104, 22);
			this.txtNumero.TabIndex = 288;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.Validated += new System.EventHandler(this.txtNumero_Validated);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 616);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 332;
			this.btnNuevo.Text = "Nuevo";
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
			this.btnConsultar.Location = new System.Drawing.Point(83, 616);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 333;
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
			this.btnEditar.Location = new System.Drawing.Point(164, 616);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 334;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(556, 616);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 335;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// uGridArticulos
			// 
			this.uGridArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridArticulos.DataSource = this.ultraDataSource5;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridArticulos.DisplayLayout.Appearance = appearance14;
			this.uGridArticulos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 98;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 92;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 149;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 179;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 305;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 59;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 39;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Width = 151;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			ultraGridBand5.Header.Caption = "";
			this.uGridArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridArticulos.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 10F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridArticulos.DisplayLayout.Override.HeaderAppearance = appearance17;
			this.uGridArticulos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridArticulos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridArticulos.Location = new System.Drawing.Point(8, 432);
			this.uGridArticulos.Name = "uGridArticulos";
			this.uGridArticulos.Size = new System.Drawing.Size(656, 176);
			this.uGridArticulos.TabIndex = 336;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8});
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
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn14,
																																 ultraDataColumn15});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 400);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(32, 16);
			this.label3.TabIndex = 338;
			this.label3.Text = "Serial";
			// 
			// txtSerial
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance21;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Location = new System.Drawing.Point(56, 400);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(216, 22);
			this.txtSerial.TabIndex = 337;
			this.txtSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerial_KeyDown);
			this.txtSerial.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerial_KeyPress);
			this.txtSerial.ValueChanged += new System.EventHandler(this.txtSerial_ValueChanged);
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(528, 400);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 19);
			this.lblUsuario.TabIndex = 339;
			// 
			// frmGuiaDeRemisionCarpa
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(672, 646);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.uGridArticulos);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtIdGuiaRemision);
			this.Controls.Add(this.cmbTraslado);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtPlaca);
			this.Controls.Add(this.cmbResponsable);
			this.Controls.Add(this.txtRucDestinatario);
			this.Controls.Add(this.txtRazonSocialDestinatario);
			this.Controls.Add(this.txtDestino);
			this.Controls.Add(this.cmbTipoFactura);
			this.Controls.Add(this.dtFinTraslado);
			this.Controls.Add(this.dtIniTraslado);
			this.Controls.Add(this.txtNumero);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmGuiaDeRemisionCarpa";
			this.Text = "Guia De Remision Carpa";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGuiaDeRemisionCarpa_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmGuiaDeRemisionCarpa_Closing);
			this.Load += new System.EventHandler(this.frmGuiaDeRemisionCarpa_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGuiaRemision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPlaca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRucDestinatario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRazonSocialDestinatario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFinTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIniTraslado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ModificaGrilla(bool bModifica)
		{
			this.uGridArticulos.Enabled = true;
			if (bModifica)
			{
				this.uGridArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				this.uGridArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				this.uGridArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				this.uGridArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmGuiaDeRemisionCarpa_Load(object sender, System.EventArgs e)
		{
			string stTipo = "0305";
			string stPermiso = "39";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (!Funcion.Permiso(stPermiso,cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Guias de Remisión de Carpa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
				return;
			}			

			#region Listas
			this.cmbTipoFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoFactura, Nombre From CompraNumero Where GuiaRemision = 1 Order By idTipoFactura Asc");			
			this.cmbTraslado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivoTraslado, Motivo From MotivosDeTraslado Where Activo = 1");		
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 1", MenuLatinium.IdUsuario));
			//this.cmbResponsable.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Order By Nombre Asc");	//Where idCargo = 10 					
			#endregion Listas

			ModificaGrilla(false);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;							
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdGuiaRemision.Value = 0;
			this.txtIdCompra.Value = 0;
			this.txtEstado.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.dtIniTraslado.Value = System.DBNull.Value;
			this.dtFinTraslado.Value = System.DBNull.Value;
			this.cmbTipoFactura.Value = System.DBNull.Value;
			this.cmbTraslado.Value = System.DBNull.Value;
			this.txtOrigen.Text = "";
			this.txtDestino.Text = "";
			this.txtRazonSocialDestinatario.Text = "";
			this.txtRucDestinatario.Text = "";
			this.cmbResponsable.Value = System.DBNull.Value;
			this.txtPlaca.Text = "";
			this.txtNotas.Text = "";
			this.txtSerial.Text = "";
			this.lblUsuario.Text = "";

			this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleGuiaDeRemision 0");

			this.cmbBodega.Enabled = false;
			this.txtNumero.Enabled = false;
			this.dtIniTraslado.Enabled = false;
			this.dtFinTraslado.Enabled = false;
			this.cmbTipoFactura.Enabled = false;
			this.cmbTraslado.Enabled = false;
			this.txtOrigen.Enabled = false;
			this.txtDestino.Enabled = false;
			this.cmbResponsable.Enabled = false;
			this.txtPlaca.Enabled = false;
			this.txtNotas.Enabled = false;
			this.txtSerial.Enabled = false;

			ModificaGrilla(false);

			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			bNuevo = true;
			bEdicion = true;
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtIniTraslado, cdsSeteoF, false);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFinTraslado, cdsSeteoF, false);
				
			this.dtFinTraslado.CalendarInfo.MinDate = (DateTime)this.dtIniTraslado.Value;
			this.dtFinTraslado.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtIniTraslado.Value).AddDays(7);

			if (iBodega == 0) 
			{
				this.cmbBodega.Value = 1; 
				this.cmbBodega.Enabled = true;
			}
			else this.cmbBodega.Value = iBodega;

			this.cmbTipoFactura.Value = 42;
			this.cmbTraslado.Value = 5;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "SELECT Ruc, Empresa FROM Seteo", true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtRucDestinatario.Text = dr.GetString(0);
				this.txtRazonSocialDestinatario.Text = dr.GetString(1);
			}
			dr.Close();

			if (bAutomatico) this.txtNumero.Enabled = true;
			this.dtIniTraslado.Enabled = true;
			this.dtFinTraslado.Enabled = true;
			this.txtDestino.Enabled = true;
			this.cmbResponsable.Enabled = true;
			this.txtPlaca.Enabled = true;
			this.txtNotas.Enabled = true;
			this.txtSerial.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			ModificaGrilla(true);		

			if (bAutomatico) this.txtNumero.Focus(); else this.dtIniTraslado.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			#region Bodega
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}
			#endregion Bodega
			#region Numero
			if (!bAutomatico)
			{
				if (this.txtNumero.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Numero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}

				char Pad = '0';
				this.txtNumero.Text = this.txtNumero.Value.ToString().PadLeft(9, Pad);

				string sSQLVal = string.Format("Exec ValidaNumeracionGuiaDeRemision '{0}', {1}", this.txtNumero.Text.ToString(), (int)this.cmbBodega.Value);
				bool bRes = Funcion.bEscalarSQL(cdsSeteoF, sSQLVal, true);
				if (bRes)
				{
					MessageBox.Show("El Numero " + this.txtNumero.Text.ToString() + " ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			}
			#endregion Numero
			#region Destino
			if (this.txtDestino.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Destino", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDestino.Focus();
				return;
			}
			#endregion Destino
			#region Transportista
			if (this.cmbResponsable.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Transportista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbResponsable.Focus();
				return;
			}
			#endregion Transportista
			#region Placa
			if (!Validacion.vbValidaPlaca(this.txtPlaca)) return;
			#endregion Placa
			#region Observaciones
			if (this.txtNotas.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNotas.Focus();
				return;
			}
			#endregion Observaciones
			#region Productos
			if (this.uGridArticulos.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los productos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			#endregion Productos
			#endregion Validacion

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

					#region Guarda Guia Remision
					DateTime dtInicio = (DateTime)this.dtIniTraslado.Value;
					DateTime dtFin = (DateTime)this.dtFinTraslado.Value;

					#region Numeración
					if (bNuevo) 
					{
						if (bAutomatico) 
						{
							string sSQLNumero = string.Format("Exec NumeracionLocales 30, {0}, 0", (int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
						}
					}
					#endregion Numeración

					string sSQL = string.Format("Exec GGuiaRemision {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, '{9}', '{10}', {11}, '{12}', {13}", 
						this.txtIdGuiaRemision.Value, this.txtIdCompra.Value, dtInicio.ToString("yyyyMMdd"), dtFin.ToString("yyyyMMdd"),
						(int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), (int)this.cmbTipoFactura.Value, (int)this.cmbTraslado.Value,
						(int)this.cmbResponsable.Value, this.txtPlaca.Text.ToString(), this.txtNotas.Text.ToString().Trim(), 0, 
						this.txtOrigen.Text.ToString(), true);
					oCmdActualiza.CommandText = sSQL;
					this.txtIdGuiaRemision.Value = oCmdActualiza.ExecuteScalar();			

					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
					{	
						string sSQLDetalle = string.Format("Exec GuardaGuiaDeRemision {0}, {1}, {2}, {3}",
							(int)dr.Cells["idDetalleGuiaDeRemision"].Value, (int)this.txtIdGuiaRemision.Value,	
							(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idSerial"].Value);
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Detalle	

					string sSQLActDetalle = string.Format("Exec ConsultaDetalleGuiaDeRemision {0}", (int)this.txtIdGuiaRemision.Value);
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.uGridArticulos.DataSource = oDTDetalle;
					#endregion Guarda Guia Remision

					if (bNuevo)
					{
						oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 6", (int)this.txtIdGuiaRemision.Value);;
						string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

						string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
						string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
						string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

						oCmdActualiza.CommandText = string.Format("Update MaestroGuiaRemision Set Autorizacion = '{0}' Where idMaestroGuiaRemision = {1}", sClaveSRI, (int)this.txtIdGuiaRemision.Value);
						oCmdActualiza.ExecuteNonQuery();
					}

					oTransaction.Commit();

					#region Controles
					this.cmbBodega.Enabled = false;
					this.txtNumero.Enabled = false;
					this.dtIniTraslado.Enabled = false;
					this.dtFinTraslado.Enabled = false;
					this.cmbTipoFactura.Enabled = false;
					this.cmbTraslado.Enabled = false;
					this.txtOrigen.Enabled = false;
					this.txtDestino.Enabled = false;
					this.cmbResponsable.Enabled = false;
					this.txtPlaca.Enabled = false;
					this.txtNotas.Enabled = false;
					this.txtSerial.Enabled = false;

					ModificaGrilla(false);

					if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = true;
					
					this.btnGuardar.Enabled = false;
					
					if (!bImpreso) 
					{
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					}
					else if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;

					if (miAcceso.Anular)this.btnAnular.Enabled = false;					

					bNuevo = false;
					bEdicion = false;
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
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.txtNumero.Enabled = true;
			this.dtIniTraslado.Enabled = true;
			this.dtFinTraslado.Enabled = true;
			this.txtDestino.Enabled = true;
			this.cmbResponsable.Enabled = true;
			this.txtPlaca.Enabled = true;
			this.txtNotas.Enabled = true;
			this.txtSerial.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			ModificaGrilla(true);
			
			bNuevo = false;
			bEdicion = true;

			this.txtNumero.Focus();
		}

		private void Consultar(int idMaestroGuiaDeRemision)
		{
			try
			{
				#region Botones
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;				
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;
				#endregion Botones

				string sSQL = string.Format("Exec consultaIndividualGRemisionCarpa {0}", idMaestroGuiaDeRemision);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					/*m.idMaestroGuiaRemision, m.Bodega, m.Numero, 
					 * m.FechaInicio, m.FechaFin, m.idTipoFactura, m.idMotivoTraslado, b.Nombre, Destino, 
					(Select Ruc From Seteo), (Select Empresa From Seteo), m.idPersonal, m.Placa, m.Observaciones, m.usuario*/

					this.txtIdGuiaRemision.Value = dr.GetInt32(0);					
					this.txtNumero.Text = dr.GetString(2);
					this.dtIniTraslado.Value = dr.GetDateTime(3);
					this.dtFinTraslado.Value = dr.GetDateTime(4);
					this.cmbBodega.Value = dr.GetInt32(1);
					this.cmbTipoFactura.Value = dr.GetInt32(5);
					this.cmbTraslado.Value = dr.GetInt32(6);
					this.txtOrigen.Text = dr.GetString(7);
					this.txtDestino.Text = dr.GetString(8);
					this.txtRucDestinatario.Text = dr.GetString(9);
					this.txtRazonSocialDestinatario.Text = dr.GetString(10);
					this.cmbResponsable.Value = dr.GetInt32(11);
					this.txtPlaca.Text = dr.GetValue(12).ToString();
					this.txtNotas.Text = dr.GetValue(13).ToString();
					bImpreso = dr.GetBoolean(14);
					this.txtEstado.Value = dr.GetInt32(15);
					this.lblUsuario.Text = dr.GetString(16);
				}
				dr.Close();			

				this.uGridArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleGuiaDeRemision {0}", idMaestroGuiaDeRemision));

				if ((int)this.txtEstado.Value == 0) this.lblEstado.Text = "PROCESADO";
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "ANULADO";

				if ((int)this.txtEstado.Value == 0)
				{
					if (miAcceso.Editar) this.btnEditar.Enabled = true;	
			
					if (!bImpreso) 
					{
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					}
					else if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
				
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
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

			using (frmBuscaGuiasDeRemision Busqueda = new frmBuscaGuiasDeRemision())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consultar((int)Busqueda.uGridArticulos.ActiveRow.Cells["idMaestroGuiaRemision"].Value);
				}
			}
		}

		private void txtSerial_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsLetterOrDigit(e.KeyChar))
			{
				e.Handled = false;
			}			
			else
			{
				e.Handled = true;
			}
		}

		private void txtSerial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ValidaSerial (int iConsulta)
		{
			string sSQL = string.Format("Exec ValidaSerialEnGuiaremision '{0}', {1}, {2}", this.txtSerial.Text.ToString(), (int)this.cmbBodega.Value, iConsulta);
			if (iConsulta == 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) == 0)
				{
					MessageBox.Show("Serial No existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
			}

			if (iConsulta == 1)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) == 0)
				{
					MessageBox.Show("Serial no esta disponible en esta Bodega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
			}

			if (iConsulta == 2)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, sSQL) == 0)
				{
					MessageBox.Show("Serial no esta disponible", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtSerial.Text = "";
					this.txtSerial.Focus();
					return;
				}
			}
			
			if (iConsulta == 3)
			{
				//s.idSerial, s.Serial, s.idArticulo, a.Codigo, a.Articulo 
				int idSerial = 0;
				string sSerial = "";
				int idArticulo = 0;
				string sCodigo = "";
				string sArticulo = "";

				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					idSerial = dr.GetInt32(0);
					sSerial = dr.GetString(1);
					idArticulo = dr.GetInt32(2);
					sCodigo = dr.GetString(3);
					sArticulo = dr.GetString(4);
				}
				dr.Close();

				this.uGridArticulos.Rows.Band.AddNew();
				this.uGridArticulos.ActiveRow.Cells["idDetalleGuiaDeRemision"].Value = 0;
				this.uGridArticulos.ActiveRow.Cells["idMaestroGuiaRemision"].Value = (int)this.txtIdGuiaRemision.Value;
				this.uGridArticulos.ActiveRow.Cells["Cantidad"].Value = 1;
				this.uGridArticulos.ActiveRow.Cells["idSerial"].Value = idSerial;
				this.uGridArticulos.ActiveRow.Cells["Serial"].Value = sSerial;
				this.uGridArticulos.ActiveRow.Cells["idArticulo"].Value = idArticulo;	
				this.uGridArticulos.ActiveRow.Cells["Codigo"].Value = sCodigo;	
				this.uGridArticulos.ActiveRow.Cells["Articulo"].Value = sArticulo;
			}
		}

		private void txtSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				#region Valida serial existente en el grid
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridArticulos.Rows.All)
				{
					string sSerial = dr.Cells["Serial"].Value.ToString();
					if (this.txtSerial.Text.ToString().Trim() == sSerial.Trim())
					{
						MessageBox.Show("Ya registro esta serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.txtSerial.Text = "";
						this.txtSerial.Focus();
						return;
					}
				}
				#endregion Valida serial existente en el grid
				
				ValidaSerial(0);
				ValidaSerial(1);
				ValidaSerial(2);
				ValidaSerial(3);
				
				this.txtSerial.Text = "";

				this.txtSerial.Focus();
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtOrigen.Text = this.cmbBodega.Text.ToString();
				bAutomatico = Funcion.bEscalarSQL(cdsSeteoF, string.Format("SELECT AutGuiaRemision FROM Bodega Where Bodega = {0}", (int)this.cmbBodega.Value), true);
			}

			if (this.cmbBodega.ActiveRow != null)
				this.cmbResponsable.DataSource = FuncionesProcedimientos.Lista_Choferes(Convert.ToDateTime(this.dtIniTraslado.Value), (int)this.cmbBodega.Value, cdsSeteoF);
		}

		private void txtNumero_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				char Pad = '0';
				this.txtNumero.Text = this.txtNumero.Text.ToString().PadLeft(9, Pad);				
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtIniTraslado.Focus();
		}

		private void dtIniTraslado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFinTraslado.Focus();
		}

		private void dtFinTraslado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDestino.Focus();
		}

		private void txtDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbResponsable.Focus();
		}

		private void cmbResponsable_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPlaca.Focus();
		}

		private void txtPlaca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtSerial.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

//			ParameterField Compra = new ParameterField ();
//			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
//			Compra.ParameterFieldName = "@idMaestroGuiaRemision";
//			IdCompra.Value = int.Parse(this.txtIdGuiaRemision.Value.ToString());
//			Compra.CurrentValues.Add (IdCompra);
//			paramFields.Add (Compra);
			
			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			Compra.ParameterFieldName = "@idCPA";
			IdCompra.Value = int.Parse(this.txtIdGuiaRemision.Value.ToString());
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);

			//Reporte miRep = new Reporte("GuiaDeRemisionCarpa.rpt", "");
			Reporte miRep = new Reporte("Doc_GR.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Guia de Remisión");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Impreso = 1 Where idMaestroGuiaRemision = {0}", (int)this.txtIdGuiaRemision.Value));

			Cursor = Cursors.Default;
		}

		private void frmGuiaDeRemisionCarpa_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Guia de Remision", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmGuiaDeRemisionCarpa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdGuiaRemision.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (bImpreso) 
																	 {
																		 if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																	 }
																	 else if (miAcceso.ReImprimir) if (this.btnImprimir.Enabled) this.btnImprimir.Enabled = true;
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0305");

						this.Consultar((int)this.txtIdGuiaRemision.Value);					
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Esta seguro de anular este documento.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update MaestroGuiaRemision Set Estado = 1 Where idMaestroGuiaRemision = {0}", (int)this.txtIdGuiaRemision.Value));

			MessageBox.Show("Documento anulado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
		}

		private void txtPlaca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericosPasaporte(sender, e);
		}
	}
}

