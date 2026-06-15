using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAnticiposClientes.
	/// </summary>
	public class frmAnticiposClientes : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredefinida;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		public System.Windows.Forms.Label label8;
		public System.Windows.Forms.Label label5;
		public System.Windows.Forms.Label lblComprobante;
		public System.Windows.Forms.Label label4;
		public System.Windows.Forms.Label lblIce;
		public System.Windows.Forms.Label label2;
		public System.Windows.Forms.Label label7;
		public System.Windows.Forms.Label label3;
		public System.Windows.Forms.Label lblCliente;
		public System.Windows.Forms.Label lblIdentificacion;
		public System.Windows.Forms.Label label1;
		public System.Windows.Forms.Label lblBodega;
		public System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		public System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAnticipo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAbono;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelefono;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccion;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnSalir;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDctoSolidario;
		private System.Windows.Forms.Button btnConsultar;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.ComponentModel.IContainer components;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdOrigen;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		public System.Windows.Forms.Label label9;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		public System.Windows.Forms.Button btnCliente;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		public System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaReserva;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTurno;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFechaCupon;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaCupon;
		public System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;

		int IdAnticipo = 0;

		public frmAnticiposClientes(int idAnticipo)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			
			IdAnticipo = idAnticipo;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		
		bool bCobroManual = false;
		string sRuc = "";
		int iBodega = 0;
		decimal dCobro = 0.00m;
		bool bBloqueado = false;
		int iTipo = 0;
		int idOrigen = 0;
		bool bCuotaCDP = true;
		int idTipoRuc = 0;
		int idFormaPago = 0;
		
		public frmAnticiposClientes(bool BCobroManual, string SRuc, int IBodega, decimal DCobro, int ITipo, int IdOrigen)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();			
			
			bCobroManual = BCobroManual;
			sRuc = SRuc;
			iBodega = IBodega;
			dCobro = DCobro;			
			iTipo = ITipo;
			idOrigen = IdOrigen;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmAnticiposClientes(bool BCobroManual, string SRuc, int IBodega, decimal DCobro, int ITipo, int IdOrigen, int IdTipoRuc, int IdFormaPago)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();			
			
			bCobroManual = BCobroManual;
			sRuc = SRuc;
			iBodega = IBodega;
			dCobro = DCobro;			
			iTipo = ITipo;
			idOrigen = IdOrigen;
			idTipoRuc = IdTipoRuc;
			idFormaPago = IdFormaPago;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmAnticiposClientes(bool BCobroManual, string SRuc, int IBodega, decimal DCobro, int ITipo, int IdOrigen, bool BCuotaCDP)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();			
			
			bCobroManual = BCobroManual;
			sRuc = SRuc;
			iBodega = IBodega;
			dCobro = DCobro;			
			iTipo = ITipo;
			idOrigen = IdOrigen;
			bCuotaCDP = BCuotaCDP;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmAnticiposClientes()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAnticiposClientes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRuc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoRuc");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredefinida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.label8 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblComprobante = new System.Windows.Forms.Label();
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
			this.txtDctoSolidario = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAbono = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCliente = new System.Windows.Forms.Button();
			this.label10 = new System.Windows.Forms.Label();
			this.dtFechaReserva = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTurno = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkFechaCupon = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaCupon = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDctoSolidario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAbono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTurno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCupon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(624, 352);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 342;
			this.txtIdAsiento.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(640, 352);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 341;
			this.txtEstado.Visible = false;
			// 
			// txtBodegaPredefinida
			// 
			this.txtBodegaPredefinida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredefinida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredefinida.Enabled = false;
			this.txtBodegaPredefinida.Location = new System.Drawing.Point(656, 352);
			this.txtBodegaPredefinida.Name = "txtBodegaPredefinida";
			this.txtBodegaPredefinida.PromptChar = ' ';
			this.txtBodegaPredefinida.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredefinida.TabIndex = 340;
			this.txtBodegaPredefinida.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 384);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(720, 8);
			this.groupBox1.TabIndex = 339;
			this.groupBox1.TabStop = false;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.Location = new System.Drawing.Point(403, 400);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 337;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			this.btnAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnAsiento_MouseDown);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(9, 400);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 332;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(165, 400);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 336;
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
			this.btnImprimir.Location = new System.Drawing.Point(244, 400);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 335;
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
			this.btnGuardar.Location = new System.Drawing.Point(324, 400);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 334;
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
			this.btnCancelar.Location = new System.Drawing.Point(565, 400);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 338;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(486, 400);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 331;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(472, 75);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(81, 16);
			this.label8.TabIndex = 330;
			this.label8.Text = "Dcto. UniBanco";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(16, 75);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(26, 16);
			this.label5.TabIndex = 327;
			this.label5.Text = "Tipo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(376, 42);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(83, 16);
			this.lblComprobante.TabIndex = 324;
			this.lblComprobante.Text = "Centro de costo";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 264);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 317;
			this.label4.Text = "Detalle";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(16, 340);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(49, 16);
			this.lblIce.TabIndex = 316;
			this.lblIce.Text = "Abono $.";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(16, 234);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 16);
			this.label2.TabIndex = 315;
			this.label2.Text = "Forma de pago";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 203);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 314;
			this.label7.Text = "Teléfono";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(16, 171);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(50, 16);
			this.label3.TabIndex = 313;
			this.label3.Text = "Dirección";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(16, 139);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(38, 16);
			this.lblCliente.TabIndex = 312;
			this.lblCliente.Text = "Cliente";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 107);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(48, 16);
			this.lblIdentificacion.TabIndex = 311;
			this.lblIdentificacion.Text = "RUC/C.I.";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(504, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 310;
			this.label1.Text = "Numero";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 42);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 309;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 308;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDctoSolidario
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDctoSolidario.Appearance = appearance1;
			this.txtDctoSolidario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDctoSolidario.Enabled = false;
			this.txtDctoSolidario.Location = new System.Drawing.Point(584, 72);
			this.txtDctoSolidario.Name = "txtDctoSolidario";
			this.txtDctoSolidario.Size = new System.Drawing.Size(120, 22);
			this.txtDctoSolidario.TabIndex = 329;
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 312;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Descripcion";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(104, 73);
			this.cmbTipo.MaxDropDownItems = 30;
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(312, 21);
			this.cmbTipo.TabIndex = 328;
			this.cmbTipo.ValueMember = "Tipo";
			this.cmbTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipo_KeyDown);
			this.cmbTipo.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTipo_Validating);
			this.cmbTipo.ValueChanged += new System.EventHandler(this.cmbTipo_ValueChanged);
			this.cmbTipo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipo_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbFormaPago
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance2;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource4;
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
			this.cmbFormaPago.Location = new System.Drawing.Point(104, 232);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(224, 21);
			this.cmbFormaPago.TabIndex = 326;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			// 
			// ultraDataSource4
			// 
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbProyecto
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance3;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 408;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(480, 40);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(225, 21);
			this.cmbProyecto.TabIndex = 325;
			this.cmbProyecto.ValueMember = "idProyecto";
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(568, 335);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 322;
			// 
			// cmbTipoRuc
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance4;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DataSource = this.ultraDataSource3;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 233;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(104, 105);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 321;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// ultraDataSource3
			// 
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(672, 352);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdCliente.TabIndex = 319;
			this.txtIdCliente.Visible = false;
			// 
			// txtIdAnticipo
			// 
			this.txtIdAnticipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAnticipo.Enabled = false;
			this.txtIdAnticipo.Location = new System.Drawing.Point(688, 352);
			this.txtIdAnticipo.Name = "txtIdAnticipo";
			this.txtIdAnticipo.PromptChar = ' ';
			this.txtIdAnticipo.Size = new System.Drawing.Size(16, 22);
			this.txtIdAnticipo.TabIndex = 318;
			this.txtIdAnticipo.Visible = false;
			// 
			// txtDetalle
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance5;
			this.txtDetalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Enabled = false;
			this.txtDetalle.Location = new System.Drawing.Point(104, 264);
			this.txtDetalle.Multiline = true;
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(600, 64);
			this.txtDetalle.TabIndex = 306;
			this.txtDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDetalle_KeyDown);
			this.txtDetalle.Validating += new System.ComponentModel.CancelEventHandler(this.txtDetalle_Validating);
			this.txtDetalle.ValueChanged += new System.EventHandler(this.txtDetalle_ValueChanged);
			// 
			// txtAbono
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAbono.Appearance = appearance6;
			this.txtAbono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAbono.Enabled = false;
			this.txtAbono.FormatString = "#,##0.00";
			this.txtAbono.Location = new System.Drawing.Point(104, 337);
			this.txtAbono.Name = "txtAbono";
			this.txtAbono.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAbono.PromptChar = ' ';
			this.txtAbono.Size = new System.Drawing.Size(112, 22);
			this.txtAbono.TabIndex = 307;
			this.txtAbono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAbono_KeyDown);
			this.txtAbono.Click += new System.EventHandler(this.txtAbono_Click);
			this.txtAbono.Validating += new System.ComponentModel.CancelEventHandler(this.txtAbono_Validating);
			this.txtAbono.ValueChanged += new System.EventHandler(this.txtAbono_ValueChanged);
			this.txtAbono.Enter += new System.EventHandler(this.txtAbono_Enter);
			// 
			// txtTelefono
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTelefono.Appearance = appearance7;
			this.txtTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelefono.Enabled = false;
			this.txtTelefono.Location = new System.Drawing.Point(104, 200);
			this.txtTelefono.Name = "txtTelefono";
			this.txtTelefono.Size = new System.Drawing.Size(112, 22);
			this.txtTelefono.TabIndex = 305;
			// 
			// txtDireccion
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccion.Appearance = appearance8;
			this.txtDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccion.Enabled = false;
			this.txtDireccion.Location = new System.Drawing.Point(104, 168);
			this.txtDireccion.Name = "txtDireccion";
			this.txtDireccion.Size = new System.Drawing.Size(600, 22);
			this.txtDireccion.TabIndex = 304;
			// 
			// txtNombre
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance9;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(104, 136);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(512, 22);
			this.txtNombre.TabIndex = 303;
			// 
			// txtRuc
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance10;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(256, 104);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 302;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNumero
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance11;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(584, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 22);
			this.txtNumero.TabIndex = 301;
			// 
			// dtFecha
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance12;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(104, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 300;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(640, 400);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 343;
			this.btnSalir.Text = "Salir";
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
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(84, 400);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 333;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource1;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(104, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(225, 21);
			this.cmbBodega.TabIndex = 344;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// txtIdOrigen
			// 
			this.txtIdOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdOrigen.Enabled = false;
			this.txtIdOrigen.Location = new System.Drawing.Point(608, 352);
			this.txtIdOrigen.Name = "txtIdOrigen";
			this.txtIdOrigen.PromptChar = ' ';
			this.txtIdOrigen.Size = new System.Drawing.Size(16, 22);
			this.txtIdOrigen.TabIndex = 345;
			this.txtIdOrigen.Visible = false;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidArticulo.Enabled = false;
			this.txtidArticulo.Location = new System.Drawing.Point(592, 352);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.Size = new System.Drawing.Size(16, 22);
			this.txtidArticulo.TabIndex = 346;
			this.txtidArticulo.Visible = false;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(360, 234);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(42, 16);
			this.label9.TabIndex = 348;
			this.label9.Text = "Articulo";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtArticulo
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance13;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Enabled = false;
			this.txtArticulo.Location = new System.Drawing.Point(408, 231);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(296, 22);
			this.txtArticulo.TabIndex = 347;
			this.txtArticulo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtArticulo_KeyDown);
			this.txtArticulo.Validated += new System.EventHandler(this.txtArticulo_Validated);
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(392, 104);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 349;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(232, 11);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(97, 16);
			this.label10.TabIndex = 351;
			this.label10.Text = "Fecha Reservación";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label10.Visible = false;
			// 
			// dtFechaReserva
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaReserva.Appearance = appearance14;
			this.dtFechaReserva.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaReserva.DateButtons.Add(dateButton2);
			this.dtFechaReserva.Enabled = false;
			this.dtFechaReserva.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaReserva.Location = new System.Drawing.Point(336, 9);
			this.dtFechaReserva.Name = "dtFechaReserva";
			this.dtFechaReserva.NonAutoSizeHeight = 23;
			this.dtFechaReserva.Size = new System.Drawing.Size(112, 21);
			this.dtFechaReserva.SpinButtonsVisible = true;
			this.dtFechaReserva.TabIndex = 423;
			this.dtFechaReserva.Value = ((object)(resources.GetObject("dtFechaReserva.Value")));
			this.dtFechaReserva.Visible = false;
			// 
			// txtTurno
			// 
			this.txtTurno.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTurno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTurno.Enabled = false;
			this.txtTurno.Location = new System.Drawing.Point(576, 352);
			this.txtTurno.Name = "txtTurno";
			this.txtTurno.PromptChar = ' ';
			this.txtTurno.Size = new System.Drawing.Size(16, 22);
			this.txtTurno.TabIndex = 424;
			this.txtTurno.Visible = false;
			// 
			// chkFechaCupon
			// 
			this.chkFechaCupon.BackColor = System.Drawing.Color.Transparent;
			this.chkFechaCupon.Enabled = false;
			this.chkFechaCupon.Location = new System.Drawing.Point(440, 104);
			this.chkFechaCupon.Name = "chkFechaCupon";
			this.chkFechaCupon.Size = new System.Drawing.Size(152, 21);
			this.chkFechaCupon.TabIndex = 425;
			this.chkFechaCupon.Text = "Vencimiento Fecha Cupon";
			this.chkFechaCupon.Visible = false;
			this.chkFechaCupon.CheckedChanged += new System.EventHandler(this.chkFechaCupon_CheckedChanged);
			// 
			// dtFechaCupon
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaCupon.Appearance = appearance15;
			this.dtFechaCupon.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaCupon.DateButtons.Add(dateButton3);
			this.dtFechaCupon.Enabled = false;
			this.dtFechaCupon.Format = "dd/MM/yyyy HH:mm";
			this.dtFechaCupon.Location = new System.Drawing.Point(592, 104);
			this.dtFechaCupon.Name = "dtFechaCupon";
			this.dtFechaCupon.NonAutoSizeHeight = 23;
			this.dtFechaCupon.Size = new System.Drawing.Size(112, 21);
			this.dtFechaCupon.SpinButtonsVisible = true;
			this.dtFechaCupon.TabIndex = 426;
			this.dtFechaCupon.Value = ((object)(resources.GetObject("dtFechaCupon.Value")));
			this.dtFechaCupon.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(240, 203);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(52, 16);
			this.label11.TabIndex = 427;
			this.label11.Text = "Vendedor";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource5
			// 
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11});
			// 
			// cmbVendedor
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance16;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource5;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 240;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbVendedor.DisplayMember = "Vendedor";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(296, 200);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(240, 21);
			this.cmbVendedor.TabIndex = 429;
			this.cmbVendedor.ValueMember = "idNomina";
			// 
			// frmAnticiposClientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(720, 430);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.dtFechaCupon);
			this.Controls.Add(this.chkFechaCupon);
			this.Controls.Add(this.txtTurno);
			this.Controls.Add(this.dtFechaReserva);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblComprobante);
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
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.txtIdOrigen);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtBodegaPredefinida);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.txtDctoSolidario);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.txtIdAnticipo);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.txtAbono);
			this.Controls.Add(this.txtTelefono);
			this.Controls.Add(this.txtDireccion);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmAnticiposClientes";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Anticipos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAnticiposClientes_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmAnticiposClientes_Closing);
			this.Load += new System.EventHandler(this.frmAnticiposClientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDctoSolidario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAbono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTurno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaCupon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		bool bImpreso = false;
		private Acceso miAcceso;
		decimal dValor = 0.00m;
		decimal dSaldo = 0.00m;
		int iValCupon = 0;
    #endregion Variables

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmAnticiposClientes_Load(object sender, System.EventArgs e)
		{
			DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	
			miAcceso = new Acceso(cdsSeteoF, "0705");
			
			if (!miAcceso.Nuevo) this.btnNuevo.Enabled = false;
			if (!miAcceso.Buscar) this.btnConsultar.Enabled = false;
						
			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ProyectoCarga 24");
			
			if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 3, 45)"));
				string sSQLT = string.Format("Exec RetornaTurno '{0}'", dtHoy.ToString("yyyyMMdd HH:mm"));		
				this.txtTurno.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLT);
				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format(" select idNomina,ApellidoPaterno + ' '+ ApellidoMaterno + ' ' +PrimerNombre + ' ' +SegundoNombre as Vendedor From Nomina where Vendedor = 1 and Estado = 1"));
				this.lblIce.Text = "Valor $.";

				#region Fecha Cupon
				this.chkFechaCupon.Visible = true;
				this.dtFechaCupon.Visible = true;
				#endregion Fecha Cupon
			}
			else if (MenuLatinium.stIdDB == 38)
			{
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 3, 11)"));
				this.txtTurno.Value = 0;
			}
			else 
			{
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago = 1"));
				this.txtTurno.Value = 0;
			}

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredefinida.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			#region Fecha Reserva
			if (MenuLatinium.stIdDB != 1 || MenuLatinium.stIdDB != 33) 
			{
				this.label10.Visible = true;
				this.dtFechaReserva.Visible = true;
				this.dtFechaReserva.CalendarInfo.MinDate = DateTime.Today;
				this.dtFechaReserva.Value = System.DBNull.Value;
			}
			#endregion Fecha Reserva

			#region Cobro Manual
			if (bCobroManual)
			{
				ListaTiposDeAnticipos(true);

				this.txtEstado.Value = 1;
				this.cmbBodega.Value = iBodega;
				
				#region Fecha
				if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
					this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");
				else
					FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				#endregion Fecha
				
				this.cmbTipo.Value = iTipo;

				if (iTipo == 16)
				{
					this.label11.Visible = true;
					this.cmbVendedor.Visible = true;
					this.cmbVendedor.Enabled = true;
				}
				else
				{
					this.label11.Visible = false;
					this.cmbVendedor.Visible = false;
				}

				this.cmbTipoRuc.Value = 5;
				if (idTipoRuc > 0) this.cmbTipoRuc.Value = idTipoRuc;				
								
				this.lblEstado.Text = "PENDIENTE";
				this.txtIdOrigen.Value = idOrigen;

				if (sRuc.Length > 0)
				{
					this.txtRuc.Text = sRuc;
				
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select idCliente, Nombre, Direccion, Telefono From Cliente Where Proveedor = 0 And Ruc = '{0}'", this.txtRuc.Text.ToString()), true);
					dr.Read();
					if (dr.HasRows)
					{
						this.txtIdCliente.Value = (int)dr.GetValue(0);
						this.txtNombre.Text = dr.GetValue(1).ToString().Trim();
						this.txtDireccion.Text = dr.GetValue(2).ToString().Trim();
						this.txtTelefono.Text = dr.GetValue(3).ToString().Trim();
						dr.Close();	
					}
					dr.Close();	
				}
				else 
				{
					this.cmbTipoRuc.Enabled = true;
					this.txtRuc.Enabled = true;
					this.txtRuc.Focus();
				}
				
				this.cmbFormaPago.Value = idFormaPago;
				this.cmbFormaPago.Enabled = false;

				if (bCuotaCDP)
					this.txtDetalle.Text = "COBRO MANUAL DE CREDIPOINT";
				else
				{
					this.txtDetalle.Text = "REGISTRO A C";
					this.txtAbono.MinValue = dCobro;
					this.txtAbono.Enabled = true;
				}
		
				if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36) 
				{
					if (iTipo == 5) this.txtDetalle.Text = "VALIDO POR 30 DIAS A PARTIR DE SU EMISION\nVALIDO PARA SUITS PREMIUM DE $35\nSUJETO A DISPONIBILIDAD\nTIEMPO DE USO 4 HORAS";
					if (iTipo == 15)
					{
						this.txtDetalle.Text = "COBRO ANTICIPADO HOSPEDAJE";
						this.cmbFormaPago.Enabled = true;
					}
				}

				this.txtAbono.Value = dCobro;

				if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
				
				if (iTipo == 5)
					this.cmbFormaPago.Enabled = true;	
								
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnAsiento.Enabled = false;				
				this.btnCancelar.Enabled = true;
				this.btnCliente.Enabled = true;
			
				//this.cmbFormaPago.Focus();

				bNuevo = true;
				bEdicion = true;
			}
			#endregion Cobro Manual

			if (IdAnticipo > 0)	Consulta(IdAnticipo);			

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\ImpresionReciboCrediPOINT_2024_V1.rpt";
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdAnticipo.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtidArticulo.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbProyecto.Value = System.DBNull.Value;
			this.cmbTipo.Value = System.DBNull.Value;
			this.txtDctoSolidario.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaReserva.Value = System.DBNull.Value;
			this.dtFechaCupon.Value = System.DBNull.Value;			
			this.txtNumero.Text = "";
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtDireccion.Text = "";
			this.txtTelefono.Text = "";
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.txtDetalle.Text = "";
			this.txtAbono.Value = 0.00m;
			this.txtIdOrigen.Value = 0;
			this.lblEstado.Text = "";
			this.txtArticulo.Text = "";

			this.dtFecha.Enabled = false;
			this.dtFechaReserva.Enabled = false;
			this.dtFechaCupon.Enabled = false;
			this.cmbTipo.Enabled = false;
			this.txtDctoSolidario.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbProyecto.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.cmbFormaPago.Enabled = false;
			this.txtDetalle.Enabled = false;
			this.txtAbono.Enabled = false;			
			this.txtArticulo.Enabled = false;
			this.chkFechaCupon.Checked = false;
			this.chkFechaCupon.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnCliente.Enabled = false;

			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbVendedor.Visible = false;
			this.label11.Visible = false;

			this.btnAnular.Text = "Anular";
			
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Tipo, Descripcion From TiposDeAnticipos Order By Tipo Asc");

			dValor = 0.00m;
			dSaldo = 0.00m;
			iValCupon = 0;

			bImpreso = false;
			bBloqueado = false;

			if (bCobroManual) this.DialogResult = DialogResult.No;
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{			
			/* anular solo del día*/
			if (!miAcceso.AnularCrediPoint)
			{
				if((DateTime) this.dtFecha.Value < (DateTime) DateTime.Today)
				{
					MessageBox.Show ("Solo puede anular documentos del día", "POINT TECHNOLOGY", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					return;
				}
			}

			#region Anula Anticipo Existente
			if (!bNuevo)
			{				
				this.Consulta((int)this.txtIdAnticipo.Value);			

				if ((int)this.txtEstado.Value == 2)
				{
					MessageBox.Show("El Documento ya esta Anulado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
					return;
				}
				if (bBloqueado)
				{
					MessageBox.Show("El Documento esta Bloqueado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (!miAcceso.AnulaCancelado)
				{
					if ((int)this.txtEstado.Value == 3 
						&& ((int)this.cmbTipo.Value != 1 && (int)this.cmbTipo.Value != 3 && (int)this.cmbTipo.Value != 5 && (int)this.cmbTipo.Value != 7 
						&& (int)this.cmbTipo.Value != 12 && (int)this.cmbTipo.Value != 15 && (int)this.cmbTipo.Value != 16 && (int)this.cmbTipo.Value != 17
						&& (int)this.cmbTipo.Value != 18 && (int)this.cmbTipo.Value != 19 && (int)this.cmbTipo.Value != 20 && (int)this.cmbTipo.Value != 21
						&& (int)this.cmbTipo.Value != 25 && (int)this.cmbTipo.Value != 26 && (int)this.cmbTipo.Value != 27 && (int)this.cmbTipo.Value != 2 
						&& (int)this.cmbTipo.Value != 29 && (int)this.cmbTipo.Value != 30 && (int)this.cmbTipo.Value != 31 && (int)this.cmbTipo.Value != 32))
					{
						MessageBox.Show("No puede Anular el Documento Esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
						return;
					}
					if (((int)this.cmbTipo.Value == 1 || (int)this.cmbTipo.Value == 5) && Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ROUND(IsNull(Saldo, 0), 2) From Anticipos Where idAnticipo = {0}", (int)this.txtIdAnticipo.Value)) < Convert.ToDecimal(this.txtAbono.Value))
					{
						MessageBox.Show("No puede Anular el Documento ya ha sido utilizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}
				if ((int)this.cmbTipo.Value == 5 && Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select IsNull(Saldo, 0) From Anticipos Where idAnticipo = {0}", (int)this.txtIdAnticipo.Value)) < Convert.ToDecimal(this.txtAbono.Value))
				{
					MessageBox.Show("No puede Anular ya ha sido pagado al Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if ((int)this.cmbTipo.Value == 2 && !miAcceso.BAnulaCobranza)
				{
					MessageBox.Show("No puede Anular Abono de Credi Point", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					return;
				}
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CobroValidaEstadoCartera {0}", (int)this.txtIdAnticipo.Value)) > 0)
				{
					MessageBox.Show("No puede Anular Cartera En transito o Pagada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				string sSQL = string.Format("Exec ValidaAnulacionCobrosCuotasPosteriores {0}, 0", (int)this.txtIdAnticipo.Value);
				if (!Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
				{
					MessageBox.Show("No puede Anular existen cuotas posteriores pagadas o abonadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtIdAnticipo.Value, 24, cdsSeteoF)) return;

				string sMensaje = "Anular";				
				if ((int)this.txtEstado.Value == 2) sMensaje = "Eliminar";

				if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de '{0}'?", sMensaje), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					if ((int)this.txtEstado.Value == 2) 
					{	
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulaAnticipos {0}, {1}, {2}, '', 1", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.cmbTipo.Value), true);

						this.btnCancelar_Click(sender, e);

						return;
					}
					
					using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
					{
						if (DialogResult.OK == miNota.ShowDialog())
						{							
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulaAnticipos {0}, {1}, {2}, '{3}', 0", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.cmbTipo.Value, miNota.txtNotas.Text.ToString()), true);
					
							this.Consulta((int)this.txtIdAnticipo.Value);							
						}
					}					
				}

				return;
			}
			#endregion Anula Anticipo Existente

			#region Anula Anticipo Nuevo
			if (!Validacion.vbComboVacio(this.cmbTipo, "Seleccione un Tipo de Documento")) return;				
				
			string sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select NumAnticipo From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));

			if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de Anular el Recibo N. '{0}'?", Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select NumAnticipo from Bodega Where Bodega = {0}", (int)this.cmbBodega.Value))), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{			
				char Pad = '0';
				sNumero = sNumero.ToString().PadLeft(7, Pad);

				using (frmAnulaRec miMsje = new frmAnulaRec (sNumero))
				{
					if (DialogResult.OK == miMsje.ShowDialog())
					{			
						using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
						{
							if (DialogResult.OK == miNota.ShowDialog())
							{					
								this.txtEstado.Value = 2;								
								
								this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 24, {0}, 0", (int)this.cmbBodega.Value));
									
								#region Guarda Anticipo Anulado
								DateTime dtFecha = (DateTime)this.dtFecha.Value;

								string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', 0, 0", 
									(int)this.txtIdAnticipo.Value, 0, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), 0, 0, 0, this.txtDetalle.Text.ToString(), (int)this.txtEstado.Value, (decimal )this.txtAbono.Value, (int)this.cmbTipo.Value, this.txtDctoSolidario.Text.ToString());
								this.txtIdAnticipo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGraba);
								#endregion Guarda Anticipo Anulado

								#region Guarda Cobro
								DateTime dtFecCaducidad = DateTime.Today;
			
								string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
									0, (int)this.txtIdAnticipo.Value, 2, 1, 0, 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0, 0);

								Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
								#endregion Guarda Cobro
  						
								#region Controles
								this.Consulta((int)this.txtIdAnticipo.Value);

								this.cmbBodega.Enabled = false;
								this.dtFecha.Enabled = false;
								this.cmbTipo.Enabled = false;
								this.txtDctoSolidario.Enabled = false;
								this.cmbTipoRuc.Enabled = false;
								this.txtRuc.Enabled = false;								
								this.cmbFormaPago.Enabled = false;
								this.txtDetalle.Enabled = false;
								this.txtAbono.Enabled = false;
								
								bNuevo = false;
								bEdicion = false;
								#endregion Controles
							}
						}
					}
				}
			}			
			#endregion Anula Anticipo Nuevo
		}

		private void ListaTiposDeAnticipos(bool bNuevo)
		{
			this.cmbTipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaTiposDeAnticipos {0}", bNuevo));
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Value = 1;
			else this.cmbBodega.Value = (int)this.txtBodegaPredefinida.Value;

			ListaTiposDeAnticipos(true);

			#region Fecha			
			if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
				this.dtFecha.Value = Funcion.dtEscalarSQL(cdsSeteoF, "Select dbo.FacturacionRetornaFecha()");
			else
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			#endregion Fecha

			this.txtIdAsiento.Value = 0;
			this.cmbTipoRuc.Value = 5;
			this.cmbFormaPago.Value = 1;			
			this.txtEstado.Value = 1;
			this.lblEstado.Text = "PENDIENTE";

			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Enabled = true;
			else this.cmbBodega.Enabled = false;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			if (miAcceso.CambiarFechaPosterior) this.dtFechaReserva.Enabled = true;
			this.cmbTipo.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtDetalle.Enabled = true;
			this.txtAbono.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			//this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnCliente.Enabled = true;
			
			if (this.cmbBodega.Enabled) this.cmbBodega.Focus();
			else if (this.dtFecha.Enabled) this.dtFecha.Focus();
			else this.cmbTipo.Focus();

			bNuevo = true;
			bEdicion = true;
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdAnticipo.Value);
	
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El Documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (bBloqueado)
			{
				MessageBox.Show("El Documento esta Bloqueado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El Documento esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 4)
			{
				MessageBox.Show("El Documento esta Cruzado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CobroValidaEstadoCartera {0}", (int)this.txtIdAnticipo.Value)) > 0)
			{
				MessageBox.Show("No puede Editar Cartera En transito o Pagada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (!FuncionesContabilidad.bEstadoCaja((int)this.cmbBodega.Value, (DateTime)this.dtFecha.Value, (int)this.txtIdAnticipo.Value, 24, cdsSeteoF)) return;
			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Enabled = true; else this.cmbBodega.Enabled = false;
			
			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtDetalle.Enabled = true;

			if ((int)this.cmbTipo.Value == 1) 
			{
				if (MenuLatinium.stIdDB != 36 || MenuLatinium.stIdDB != 29 || MenuLatinium.stIdDB != 34) //29, 34
				{
					this.txtArticulo.Enabled = false;
				}
				else
				{
					this.txtArticulo.Enabled = true;
				}
			}

			if ((int)this.cmbTipo.Value != 2 && (int)this.cmbTipo.Value != 3) 
			{
				if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
				this.cmbTipo.Enabled = true;
				this.txtAbono.Enabled = true;				
			}

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

			if (this.cmbBodega.Enabled) this.cmbBodega.Focus();
			else if (this.dtFecha.Enabled) this.dtFecha.Focus();
			else this.cmbTipo.Focus();

			#region Fecha Cupones
			if (iValCupon == 1)
			{
				string sSQLBusqueda = string.Format("Select FechaSistema From CuponesAnticipoCompra Where idAnticipo = {0}", this.txtIdAnticipo.Value);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.dtFechaCupon.Value = dr.GetDateTime(0);
				}
				dr.Close();

				this.chkFechaCupon.Enabled = true;
			}
			#endregion Fecha Cupones
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipo.Focus();
		}

		private void cmbTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTipoRuc.Focus();
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
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

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
				if (this.txtArticulo.Enabled) this.txtArticulo.Focus(); else this.txtDetalle.Focus();
		}

		private void txtDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtAbono.Focus();
		}

		private void txtAbono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
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
				else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Cobros"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No puede ingresar Facturas en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;					
				}
			}
		}

		private void cmbTipo_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipo.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Tipo de Anticipo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					e.Cancel = true;
				}
				else if ((int)this.cmbTipo.Value == 2)
				{
					MessageBox.Show("Para Registrar Cobros de CrediPoint debe Ingresar a Cobranza", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					e.Cancel = true;
				}
				else if ((int)this.cmbTipo.Value == 3)
				{
					MessageBox.Show("Para Registrar Cobros de Credito debe Ingresar a Cuentas Por Cobrar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtDetalle_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				if (this.txtDetalle.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Detalle", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}
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

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtIdCliente.Value = 0;
				this.txtNombre.Text = "";
				
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

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cmbTipo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtDetalle.Enabled = true;
				this.txtAbono.Enabled = true;
				this.txtidArticulo.Value = 0;
				this.txtArticulo.Text = "";
				this.txtDetalle.Text = "";
				this.txtAbono.Value = 0.00m;
				
				if ((int)this.cmbTipo.Value == 5) this.txtDctoSolidario.Enabled = true;	else this.txtDctoSolidario.Enabled = false;
				
				if ((int)this.cmbTipo.Value == 1)
				{
					if (MenuLatinium.stIdDB == 36 || MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34) //29, 34
					{
						this.txtArticulo.Enabled = false;
					}
					else
					{
						this.txtArticulo.Enabled = true; 
					}
				}
				else if ((int)this.cmbTipo.Value == 20)
				{
					this.txtidArticulo.Value = 9834;
					this.txtArticulo.Text = "PORT00074";
					this.txtDetalle.Enabled = false;
					this.txtDetalle.Text = "PRE VENTA PORTATIL HP 15GW0024LA/AMD RYZEN 5-3450U/4GB/SSD 256GB/15.6/W10H ";
					this.txtAbono.Enabled = true;					
					this.txtAbono.Value = 599.00m;					
				}
				else
				{
					this.txtidArticulo.Value = 0;
					this.txtArticulo.Text = "";
				}
				if ((int)this.cmbTipo.Value == 16)
				{
					this.label11.Visible = true;
					this.cmbVendedor.Visible = true;
					this.cmbVendedor.Enabled = true;
				}
				else
				{
					this.label11.Visible = false;
					this.cmbVendedor.Visible = false;
				}
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if ((int)this.cmbTipoRuc.Value == 4) this.txtRuc.MaxLength = 13;
				if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

				#region Consumidor Final
				if ((int)this.cmbTipoRuc.Value == 11)
				{
					this.txtRuc.Enabled = false;
					this.txtRuc.MaxLength = 13;
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select cr.idCliente, cl.Ruc From Cliente cl Inner Join ClienteTipoRuc cr On cr.idCliente = cl.idCliente Where cr.idTipoRuc = 11", true);
					dr.Read();
					if (dr.HasRows)
					{
						this.txtIdCliente.Value = dr.GetInt32(0);
						this.txtRuc.Value = dr.GetString(1);
					}
					dr.Close();
					CargaCliente();
					this.cmbTipoRuc.Enabled = false;
				}
				#endregion Consumidor Final
			}
		}

		private void Consulta(int idAnticipo)
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
				this.dtFechaReserva.Enabled = false;

				int IBodega = 0;
				ListaTiposDeAnticipos(false);

				string sSQLBusqueda = string.Format("Exec BuscaAnticiposIndividual {0}", idAnticipo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdAnticipo.Value = dr.GetInt32(0);						
					this.dtFecha.Value = dr.GetDateTime(1);
					this.txtNumero.Text = dr.GetValue(2).ToString();
					this.cmbTipo.Value = dr.GetInt32(3);
					this.txtDctoSolidario.Text = dr.GetValue(4).ToString();
					IBodega = dr.GetInt32(5);
					if (dr.GetValue(6) != System.DBNull.Value) this.cmbProyecto.Value = dr.GetInt32(6);					
					this.txtIdCliente.Value = dr.GetInt32(8);
					this.cmbTipoRuc.Value = dr.GetInt32(9);
					this.txtRuc.Text = dr.GetValue(10).ToString();
					this.txtNombre.Text = dr.GetValue(11).ToString();
					this.txtDireccion.Text = dr.GetValue(12).ToString();
					this.txtTelefono.Text = dr.GetValue(13).ToString();
					this.cmbFormaPago.Value = dr.GetInt32(14);
					this.txtDetalle.Text = dr.GetValue(16).ToString();
					this.txtEstado.Value = dr.GetInt32(17);
					this.lblEstado.Text = dr.GetValue(18).ToString();
					this.txtAbono.Value = dr.GetDecimal(19);
					if (dr.GetValue(20) != System.DBNull.Value)this.txtIdAsiento.Value = dr.GetInt32(20);
					if (dr.GetValue(21) != System.DBNull.Value)this.txtIdOrigen.Value = dr.GetInt32(21);
					bBloqueado = dr.GetBoolean(22);
					this.txtidArticulo.Value = dr.GetInt32(23);
					this.txtArticulo.Text = dr.GetString(24);
					this.dtFechaReserva.Value = dr.GetDateTime(25);
					this.txtTurno.Value = dr.GetInt32(26);
					this.cmbVendedor.Value = dr.GetInt32(27);
				}
				dr.Close();
			
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
				if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "ANULADO";
				if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "CANCELADO";
						
				if ((int)this.txtEstado.Value != 2)
				{
					if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Cobros"))
					{					
						if ((int)this.cmbTipo.Value != 3 && (int)this.cmbTipo.Value != 2)	
							if (miAcceso.Editar && !bBloqueado) this.btnEditar.Enabled = true;

						if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
						if (miAcceso.Anular && !bBloqueado) this.btnAnular.Enabled = true;
					}

					if (miAcceso.Imprimir && !bImpreso) this.btnImprimir.Enabled = true;
					if (miAcceso.ReImprimir && bImpreso) this.btnImprimir.Enabled = true;

					dValor = Convert.ToDecimal(this.txtAbono.Value);			
				}
				else if ((int)this.txtEstado.Value == 2 && !bBloqueado)
				{
					this.btnAnular.Enabled = true;
					this.btnAnular.Text = "Eliminar";
				}
				if((int)this.cmbTipo.Value == 16)
				{
					this.label11.Visible = true;
					this.cmbVendedor.Visible = true;				
				}
								
				this.cmbBodega.Value = IBodega;
				this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 24, cdsSeteoF);
				
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

				this.btnCancelar.Enabled = true;

				#region Fecha Cupones
				if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
				{
					iValCupon = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From CuponesAnticipoCompra Where idAnticipo = {0}", this.txtIdAnticipo.Value));
				}
				#endregion Fecha Cupones
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaAnticipos Busqueda = new frmBuscaAnticipos((int)this.txtBodegaPredefinida.Value, 0, miAcceso.BDia))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idAnticipo"].Value);
				}
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iOrigen = 2;
			int iVendedor = 0;
			if ((int)this.cmbTipo.Value == 2) iOrigen = 3;
			if ((int)this.cmbTipo.Value == 3) iOrigen = 4;

			if (!bNuevo) dSaldo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Saldo From Anticipos Where idAnticipo = {0}", (int)this.txtIdAnticipo.Value));

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Cobros", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (!Validacion.vbComboVacio(this.cmbTipo, "Seleccione un Tipo de Documento")) return;

			if ((int)this.cmbTipo.Value == 2 && bNuevo)
			{
				MessageBox.Show("Para Registrar Cobros de CrediPoint debe Ingresar a Cobranza", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();			
				return;
			}

			if ((int)this.cmbTipo.Value == 3 && bNuevo)
			{
				MessageBox.Show("Para Registrar Cobros de Credito debe Ingresar a Cuentas Por Cobrar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();
				return;
			}
			
			if ((int)this.cmbTipo.Value == 12 && bNuevo && !bCobroManual)
			{
				MessageBox.Show("Para Registrar Abonos de CrediPOINT debe Ingresar a Cobranza", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTipo.Focus();
				return;
			}

			if ((int)this.cmbTipo.Value == 5)
			{
				int iReg = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaClienteSolidario {0}", (int)this.txtIdCliente.Value));

				if (iReg == 0)
				{
					MessageBox.Show("Cliente no tiene creditos pendientes de cartera solidario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTipo.Focus();
					return;
				}
			}

			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
			
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, 1, cdsSeteoF)) return;	
			
			if (MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33)
			{
				if ((int)this.cmbTipo.Value == 1 && (int)this.txtidArticulo.Value == 0)
				{
					MessageBox.Show("Para registrar un anticipo debe ingresar la descripcion del articulo por el cual el cliente deja el abono", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtArticulo.Focus();
					return;
				}

				if ((int)this.cmbTipo.Value == 1 && (int)this.txtidArticulo.Value == 0)
				{
					MessageBox.Show("Para registrar un anticipo debe ingresar la descripcion del articulo por el cual el cliente deja el abono", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtArticulo.Focus();
					return;
				}

				if ((int)this.cmbTipo.Value == 1 && (int)this.txtidArticulo.Value == 9834)
				{
					MessageBox.Show("La reserva de una PORT00074 debe realizarse con la Opción PRE VENTA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtArticulo.Focus();
					return;
				}

				if ((int)this.cmbTipo.Value == 1 && (int)this.txtidArticulo.Value == 9834)
				{
					MessageBox.Show("La reserva de una PORT00074 debe realizarse con la Opción PRE VENTA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtArticulo.Focus();
					return;
				}
			}

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
			
			if ((int)this.cmbTipo.Value == 20)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, "Select COUNT(*) From Anticipos Where Tipo = 20 And Estado <> 2") > 600)
				{
					MessageBox.Show("Ya se ha completado el limite de Pre Venta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return;
				}
			}
			
			if ((int)this.cmbTipo.Value == 20)
			{
				if (!Validacion.vbCampoDecimalVacio(this.txtAbono, "Ingrese el Valor", 599.00m, 599.00m)) return; 
			}

			string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec AnticiposValidaCredito '{0}', {1}", this.txtRuc.Text.ToString(), (int)this.cmbTipo.Value));
      	
			if (sMensaje.Length > 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
				return;
			}
			int idVendedor = 0;
			if (this.cmbVendedor.Value != null && this.cmbVendedor.Value != DBNull.Value)
				idVendedor = (int)this.cmbVendedor.Value;

			if ((int)this.cmbTipo.Value == 16 && bNuevo && idVendedor == 0)
			{
				MessageBox.Show("Seleccione un Vendedor","Point Technology",MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.cmbVendedor.Focus();
				return;
			}
			#endregion Validacion

			if (cmbVendedor.Value == null || cmbVendedor.Value == DBNull.Value)
				iVendedor = 0;
			else
				iVendedor = Convert.ToInt32(cmbVendedor.Value);

			if (MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
			{
				using (VentaFPagoLot miCobro = new VentaFPagoLot ((int)this.txtIdAnticipo.Value, (int)this.cmbFormaPago.Value, Convert.ToDecimal(this.txtAbono.Value), iOrigen, 4))
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
									string sSQLNumero = string.Format("Exec NumeracionLocales 24, {0}, 0", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNumero;
									this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
								}
								#endregion Numeracion

								#region Guarda Anticipo
								DateTime dtFecha = (DateTime) this.dtFecha.Value;
								DateTime dtFechaReserva = DateTime.Today;
								if (this.dtFechaReserva.Value != System.DBNull.Value) dtFechaReserva = (DateTime) this.dtFechaReserva.Value;

								string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', {13}, '{14}', '{15}', {16},{17}", 
									(int)this.txtIdAnticipo.Value, 0, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), (int)this.txtIdCliente.Value, (int)this.cmbFormaPago.Value, 0, 
									this.txtDetalle.Text.ToString(), (int)this.txtEstado.Value, Convert.ToDecimal(this.txtAbono.Value), (int)this.cmbTipo.Value, this.txtDctoSolidario.Text.ToString(),
									(int)this.txtIdOrigen.Value, (int)this.txtidArticulo.Value, dtFechaReserva.ToString("yyyyMMdd"), (int)this.txtTurno.Value,iVendedor);
								oCmdActualiza.CommandText = sSQLGraba;
								this.txtIdAnticipo.Value = (int)oCmdActualiza.ExecuteScalar();
								#endregion Guarda Anticipo

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
											(int)this.txtIdAnticipo.Value, iOrigen,
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

								#region Saldo
								string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 1", (int)this.txtIdAnticipo.Value, (int)this.cmbTipo.Value);
								oCmdActualiza.CommandText = sSQLSaldo;
								oCmdActualiza.ExecuteNonQuery();
								#endregion Saldo

								#region Fecha Cupon
								if(this.chkFechaCupon.Checked)
								{
									if(iValCupon == 1)
									{
										DateTime dtFechaCupon = DateTime.Today;
										if (this.dtFechaCupon.Value != System.DBNull.Value) dtFechaCupon = (DateTime) this.dtFechaCupon.Value;

										string sSQLFecCupon = string.Format("Exec CambioFechaCupones {0}, '{1}'", (int)this.txtIdAnticipo.Value, dtFechaCupon.ToString("yyyyMMdd"));
										oCmdActualiza.CommandText = sSQLFecCupon;										
										oCmdActualiza.ExecuteNonQuery();
									}
								}
								#endregion Fecha Cupon

								#region Asiento
								if ((int)this.cmbTipo.Value == 2)
								{										
									string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.txtIdOrigen.Value);
									oCmdActualiza.CommandText = sSQLAsiento;										
									oCmdActualiza.ExecuteNonQuery();
								}
								else if ((int)this.cmbTipo.Value == 3)
								{
									string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, {1}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value);
									oCmdActualiza.CommandText = sSQLAsiento;
									this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();
								}
								else
								{
									string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", (int)this.txtIdAsiento.Value, (int)this.txtIdAnticipo.Value, (int) this.cmbTipo.Value, iOrigen);
									oCmdActualiza.CommandText = sSQLAsiento;
									this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();

									if ((int)this.cmbTipo.Value == 11)
									{
										string sSQLADep = string.Format("Exec CreaAsientoDepreciacionCP {0}", (int)this.txtIdAnticipo.Value);
										oCmdActualiza.CommandText = sSQLADep;
										oCmdActualiza.ExecuteNonQuery();
									}
								}								
								#endregion Asiento

								oTransaction.Commit();
							
								oConexion.Close();
		
								#region Controles
								this.cmbBodega.Enabled = false;
								this.dtFecha.Enabled = false;
								this.dtFechaReserva.Enabled = false;
								this.cmbTipo.Enabled = false;
								this.txtDctoSolidario.Enabled = false;
								this.cmbTipoRuc.Enabled = false;
								this.txtRuc.Enabled = false;
								this.cmbFormaPago.Enabled = false;
								this.txtDetalle.Enabled = false;
								this.txtAbono.Enabled = false;
								this.txtArticulo.Enabled = false;
								this.chkFechaCupon.Enabled = false;
								this.chkFechaCupon.Checked = false;
								this.dtFechaCupon.Enabled = false;

								if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
								if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
								if (miAcceso.Editar) this.btnEditar.Enabled = true;
								if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
								if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
								if (miAcceso.Anular) this.btnAnular.Enabled = true;
								this.btnGuardar.Enabled = false;
								this.btnCliente.Enabled = false;

								bNuevo = false;
								bEdicion = false;

								iValCupon = 0;
								#endregion Controles
								
								#region Impresion
								if (bCobroManual) this.DialogResult = DialogResult.OK;
								else
								{
									using (frmNumeracionRecibo Numeracion = new frmNumeracionRecibo (this.txtNumero.Text.ToString()))
									{
										if (DialogResult.OK == Numeracion.ShowDialog())
											this.btnImprimir_Click(sender, e);
									}
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
					}
				}
			}
			else
			{
				using (VentaFormaDePago miCobro = new VentaFormaDePago ((int)this.txtIdAnticipo.Value, (int)this.cmbFormaPago.Value, Convert.ToDecimal(this.txtAbono.Value), iOrigen, (DateTime)this.dtFecha.Value, (int)this.cmbBodega.Value, false))
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
									string sSQLNumero = string.Format("Exec NumeracionLocales 24, {0}, 0", (int)this.cmbBodega.Value);
									oCmdActualiza.CommandText = sSQLNumero;
									this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
								}
								#endregion Numeracion

								#region Guarda Anticipo
								DateTime dtFecha = (DateTime) this.dtFecha.Value;
								DateTime dtFechaReserva = DateTime.Today;
								if (this.dtFechaReserva.Value != System.DBNull.Value) dtFechaReserva = (DateTime) this.dtFechaReserva.Value;

								string sSQLGraba = string.Format("Exec GuardaAnticipos {0}, {1}, '{2}', {3}, '{4}', {5}, {6}, {7}, '{8}', {9}, {10}, {11}, '{12}', {13}, '{14}', '{15}'", 
									(int)this.txtIdAnticipo.Value, 0, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), (int)this.txtIdCliente.Value, (int)this.cmbFormaPago.Value, 0, 
									this.txtDetalle.Text.ToString(), (int)this.txtEstado.Value, Convert.ToDecimal(this.txtAbono.Value), (int)this.cmbTipo.Value, this.txtDctoSolidario.Text.ToString(),
									(int)this.txtIdOrigen.Value, (int)this.txtidArticulo.Value, dtFechaReserva.ToString("yyyyMMdd"));
								oCmdActualiza.CommandText = sSQLGraba;
								this.txtIdAnticipo.Value = (int)oCmdActualiza.ExecuteScalar();
								#endregion Guarda Anticipo

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
											(int)this.txtIdAnticipo.Value, iOrigen,
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

								#region Saldo
								string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 1", (int)this.txtIdAnticipo.Value, (int)this.cmbTipo.Value);
								oCmdActualiza.CommandText = sSQLSaldo;
								oCmdActualiza.ExecuteNonQuery();
								#endregion Saldo

								#region Asiento
								if ((int)this.cmbTipo.Value == 2)
								{										
									string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.txtIdOrigen.Value);
									oCmdActualiza.CommandText = sSQLAsiento;										
									oCmdActualiza.ExecuteNonQuery();										
								}
								else if ((int)this.cmbTipo.Value == 3)
								{
									string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, {1}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value);
									oCmdActualiza.CommandText = sSQLAsiento;
									this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();
								}
								else
								{
									//								string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", (int)this.txtIdAsiento.Value, (int)this.txtIdAnticipo.Value, (int) this.cmbTipo.Value, iOrigen);
									//								oCmdActualiza.CommandText = sSQLAsiento;
									//								this.txtIdAsiento.Value = (int)oCmdActualiza.ExecuteScalar();

									if ((int)this.cmbTipo.Value == 11)
									{
										string sSQLADep = string.Format("Exec CreaAsientoDepreciacionCP {0}", (int)this.txtIdAnticipo.Value);
										oCmdActualiza.CommandText = sSQLADep;
										oCmdActualiza.ExecuteNonQuery();
									}
								}								
								#endregion Asiento

								oTransaction.Commit();
							
								oConexion.Close();
		
								#region Controles
								this.cmbBodega.Enabled = false;
								this.dtFecha.Enabled = false;
								this.dtFechaReserva.Enabled = false;
								this.cmbTipo.Enabled = false;
								this.txtDctoSolidario.Enabled = false;
								this.cmbTipoRuc.Enabled = false;
								this.txtRuc.Enabled = false;
								this.cmbFormaPago.Enabled = false;
								this.txtDetalle.Enabled = false;
								this.txtAbono.Enabled = false;
								this.txtArticulo.Enabled = false;

								if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
								if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
								if (miAcceso.Editar) this.btnEditar.Enabled = true;
								if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
								if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
								if (miAcceso.Anular) this.btnAnular.Enabled = true;
								this.btnGuardar.Enabled = false;
								this.btnCliente.Enabled = false;

								bNuevo = false;
								bEdicion = false;
								#endregion Controles
								
								#region Impresion
								if (bCobroManual) this.DialogResult = DialogResult.OK;
								else
								{
									using (frmNumeracionRecibo Numeracion = new frmNumeracionRecibo (this.txtNumero.Text.ToString()))
									{
										if (DialogResult.OK == Numeracion.ShowDialog())
											this.btnImprimir_Click(sender, e);
									}
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
					}
				}	
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			Facturacion.ImpresionRecibos((int)this.txtIdAnticipo.Value, (int)this.cmbTipo.Value, cdsSeteoF);

			if (!miAcceso.ReImprimir) this.btnImprimir.Enabled = false;
			bImpreso = true;
			Cursor = Cursors.Default;
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{	
			int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Estado From Anticipos Where idAnticipo = {0}", (int)this.txtIdAnticipo.Value));
			if (iEstado == 2) 
			{
				MessageBox.Show("No puede crear un asiento de una transacción Anulada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.Consulta((int)this.txtIdAnticipo.Value);
			}			
			else
			{				
				string sSQLExisteAsiento = string.Format("Exec ValidaAsientoAnticipo {0}", (int)this.txtIdAnticipo.Value);
				string stMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLExisteAsiento, true);

				if (!stMensaje.StartsWith("No Existe"))
				{			
					if (DialogResult.Yes == MessageBox.Show(stMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
					{					
						this.Cursor = Cursors.WaitCursor;	
						#region Crea Asiento
						int iOrigen = 2;
						if ((int)this.cmbTipo.Value == 2) iOrigen = 3;
						if ((int)this.cmbTipo.Value == 3) iOrigen = 4;

						if ((int)this.cmbTipo.Value != 2 && (int)this.cmbTipo.Value != 3)
						{
							string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", (int)this.txtIdAsiento.Value, (int)this.txtIdAnticipo.Value, (int) this.cmbTipo.Value, iOrigen);
							this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

							if ((int)this.cmbTipo.Value == 11)
							{
								string sSQLADep = string.Format("Exec CreaAsientoDepreciacionCP {0}", (int)this.txtIdAnticipo.Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLADep);
							}
						}
						else if ((int)this.cmbTipo.Value == 2)
						{								
							string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.txtIdOrigen.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);									
						}
						else if ((int)this.cmbTipo.Value == 3)
						{
							string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, {1}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
						}
						else if ((int)this.cmbTipo.Value == 13)
						{								
							string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.txtIdOrigen.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);									
						}
						#endregion Crea Asiento
						Cursor = Cursors.Default;
					}
				}
				else
				{
					this.Cursor = Cursors.WaitCursor;	
					#region Crea Asiento
					int iOrigen = 2;
					if ((int)this.cmbTipo.Value == 2) iOrigen = 3;
					if ((int)this.cmbTipo.Value == 3) iOrigen = 4;

					if ((int)this.cmbTipo.Value != 3)
					{
						if ((int)this.cmbTipo.Value == 2)
						{										
							string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.txtIdOrigen.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);									
						}
						else if ((int)this.cmbTipo.Value == 3)
						{
							string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, {1}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
						}
						else if ((int)this.cmbTipo.Value == 13)
						{								
							string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", (int)this.txtIdAnticipo.Value, (int)this.txtIdAsiento.Value, (int)this.txtIdOrigen.Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);									
						}
						else
						{
							string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", (int)this.txtIdAsiento.Value, (int)this.txtIdAnticipo.Value, (int) this.cmbTipo.Value, iOrigen);
							this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

							if ((int)this.cmbTipo.Value == 11)
							{
								string sSQLADep = string.Format("Exec CreaAsientoDepreciacionCP {0}", (int)this.txtIdAnticipo.Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLADep);
							}
						}
					}
					#endregion Crea Asiento
					Cursor = Cursors.Default;
				}
			}
		}

		private void txtAbono_Enter(object sender, System.EventArgs e)
		{
			this.txtAbono.SelectAll();
		}

		private void txtAbono_Click(object sender, System.EventArgs e)
		{
			this.txtAbono.SelectAll();
		}

		private void btnAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right) 
			{
				if ((int)this.txtIdAsiento.Value == 0)
				{
					int idAsiento = 0;
					idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idAsiento, 0) From Anticipos Where idAsiento = {0}", (int)this.txtIdAnticipo.Value));
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

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega.ActiveRow != null) 
				{				
					string sSQL = string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value);
					this.cmbProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL, true);
				}
			}

			if (this.cmbBodega.ActiveRow != null) this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 24, cdsSeteoF);
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.dtFecha.Enabled) this.dtFecha.Focus(); else this.cmbTipo.Focus();
		}

		private void frmAnticiposClientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdAnticipo.Value > 0)
				{	
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) 
					{
						if (miAcceso.Imprimir && !bImpreso) this.btnImprimir_Click(sender, e);
						if (miAcceso.ReImprimir && bImpreso) this.btnImprimir_Click(sender, e);
					}
					if (e.KeyCode == Keys.F5)
					{
						this.Consulta((int)this.txtIdAnticipo.Value);				
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

		private void txtAbono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmAnticiposClientes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un recibo de cobro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void cmbTipo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtDetalle_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturacionListaBodegas {0}, 24, '{1}'", MenuLatinium.IdUsuario, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtArticulo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bEdicion)
			{
				if (e.KeyCode == Keys.F3)
				{
					int iBodega = 0;
					if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

					using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtArticulo.Text.ToString(), iBodega, 1, 0, Convert.ToDateTime(this.dtFecha.Value), false))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							this.txtArticulo.Text = miBusqueda.sCodigo;
							this.txtidArticulo.Value = (int)miBusqueda.idArticulo;
						}
					}
				}
			}
		}

		private void txtArticulo_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtArticulo.Text.ToString().Length > 0)
				{
					string sSQL = string.Format("Select idArticulo, Codigo From Articulo Where Discontinuo = 0 And (Codigo = '{0}' Or Articulo = '{0}')", 
						this.txtArticulo.Text.ToString().Trim()); 
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						this.txtidArticulo.Value = drArticulo.GetInt32(0);
						this.txtArticulo.Text = drArticulo.GetString(1);				
						drArticulo.Close();
					}										
					else
					{
						drArticulo.Close();
						MessageBox.Show("El articulo no existe o esta descontinuado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.txtidArticulo.Value = 0;
					}
				}
			}
		}
		
		private void CargaCliente()
		{
			string[] sParametros = new string[11];
			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtRuc.Text.ToString(), cdsSeteoF);

			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.cmbTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];
			if (sParametros[7] != null) this.txtDireccion.Text = sParametros[7];
			if (sParametros[8] != null) this.txtTelefono.Text = sParametros[8];		
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

		private void dtFechaReserva_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFechaReserva.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Fecha Reservacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}
		}

		private void dtFechaReserva_ValueChanged(object sender, System.EventArgs e)
		{

		}

		private void chkFechaCupon_CheckedChanged(object sender, System.EventArgs e)
		{
			if(this.chkFechaCupon.Checked)
			{
				this.dtFechaCupon.Enabled = true;
			}
			else
			{
				this.dtFechaCupon.Enabled = false;
			}
		}
	}
}

