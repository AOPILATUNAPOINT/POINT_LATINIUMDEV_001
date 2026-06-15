using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.IO;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDepositosPendientes.
	/// </summary>
	public class frmDepositosPendientes : System.Windows.Forms.Form
	{
		public System.Windows.Forms.Label label4;
		public System.Windows.Forms.Label lblIce;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		public System.Windows.Forms.Label lblCliente;
		public System.Windows.Forms.Label lblIdentificacion;
		public System.Windows.Forms.Label lblBodega;
		public System.Windows.Forms.Label label6;
		public System.Windows.Forms.Button btnCliente;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDetalle;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAbono;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBancos;
		public System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdOrigen;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredefinida;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAnticipo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		bool bCobroManual = false;
		string sRuc = "";
		int iBodega = 0;
		decimal dCobro = 0.00m;
		bool bBloqueado = false;
		int iTipo = 0;
		int idOrigen = 0;


		//int iBodega = 0;
	//	string sUrlNube = "";
		bool bActualiza = false;
		//bool bEdicion = false;
		int iDiasVencidos = 0;
		int iIdCompra = 0;	
		decimal dTotal= 0.00m;	
		int NumCuotas =0;	
		int iIdEntidadFinanciera;	
		decimal ValorCuotas;
		int NumeroLote;
		bool Confirmado;	
		string NumeroDeOperacion;
		bool Reliquido;
		int ER;
		bool Asignado;
		string Estado;
		int EstadoSorteo = 0;
		int EstadoCelular = 0;
		DateTime Vence = DateTime.Today;
		decimal dInteres = 0.00m;
		bool bFiniquitos = false;
		decimal dSaldoSD = 0.00m;
		DateTime dtUltimaFechaPago = DateTime.Today;
		DateTime dFechaPago;
		DateTime FechaBefore;
		int iFormaPago=0;
		int idDepositosPorliquida=0;
		private DateTime dtFechaDeposito = DateTime.MinValue;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdDepositosPendientes;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumDeposito;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private System.Windows.Forms.Button btnSubir;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnImprimir;
		public System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAutorizacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProntoPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoVencido;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasVencidos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoAFavor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProximaVencer;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSiguienteCuota;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaVencida;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridPagos;
		bool bCuotaCDP = true;

		public frmDepositosPendientes()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDepositosPendientes));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vence");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuota");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InteresTotal");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalTotal");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABCapital");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gastos");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DsctoGastos");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DsctoIntMora");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retraso");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCancelacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Secuencial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalDiferimiento");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferimiento");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			this.label4 = new System.Windows.Forms.Label();
			this.lblIce = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btnCliente = new System.Windows.Forms.Button();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtDetalle = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAbono = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBancos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label26 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdOrigen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdDepositosPendientes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredefinida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAnticipo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumDeposito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSubir = new System.Windows.Forms.Button();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtIdAutorizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProntoPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasVencidos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoAFavor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProximaVencer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSiguienteCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaVencida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.gridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAbono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDepositosPendientes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAFavor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProximaVencer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSiguienteCuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaVencida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridPagos)).BeginInit();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 208);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(40, 16);
			this.label4.TabIndex = 360;
			this.label4.Text = "Detalle";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIce
			// 
			this.lblIce.AutoSize = true;
			this.lblIce.Location = new System.Drawing.Point(16, 280);
			this.lblIce.Name = "lblIce";
			this.lblIce.Size = new System.Drawing.Size(49, 16);
			this.lblIce.TabIndex = 359;
			this.lblIce.Text = "Abono $.";
			this.lblIce.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
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
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(16, 112);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(40, 16);
			this.lblCliente.TabIndex = 358;
			this.lblCliente.Text = "Cliente";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(16, 80);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(40, 16);
			this.lblIdentificacion.TabIndex = 357;
			this.lblIdentificacion.Text = "Cédula";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(16, 48);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 356;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 18);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 355;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnCliente
			// 
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(240, 80);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 22);
			this.btnCliente.TabIndex = 8;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource1;
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
			this.cmbBodega.Location = new System.Drawing.Point(104, 48);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(225, 21);
			this.cmbBodega.TabIndex = 6;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 337);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(88, 23);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(104, 337);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 1;
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
			this.btnEditar.Location = new System.Drawing.Point(192, 337);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 23);
			this.btnEditar.TabIndex = 2;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(280, 337);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 14;
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
			this.btnCancelar.Location = new System.Drawing.Point(544, 337);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 3;
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
			this.btnAnular.Location = new System.Drawing.Point(632, 304);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(72, 23);
			this.btnAnular.TabIndex = 361;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// txtDetalle
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDetalle.Appearance = appearance1;
			this.txtDetalle.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDetalle.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDetalle.Enabled = false;
			this.txtDetalle.Location = new System.Drawing.Point(104, 208);
			this.txtDetalle.Multiline = true;
			this.txtDetalle.Name = "txtDetalle";
			this.txtDetalle.Size = new System.Drawing.Size(608, 64);
			this.txtDetalle.TabIndex = 12;
			this.txtDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDetalle_KeyDown);
			// 
			// txtAbono
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAbono.Appearance = appearance2;
			this.txtAbono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAbono.Enabled = false;
			this.txtAbono.FormatString = "#,##0.00";
			this.txtAbono.Location = new System.Drawing.Point(104, 280);
			this.txtAbono.Name = "txtAbono";
			this.txtAbono.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAbono.PromptChar = ' ';
			this.txtAbono.Size = new System.Drawing.Size(112, 21);
			this.txtAbono.TabIndex = 13;
			this.txtAbono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAbono_KeyDown);
			this.txtAbono.ValueChanged += new System.EventHandler(this.txtAbono_ValueChanged);
			this.txtAbono.Enter += new System.EventHandler(this.txtAbono_Enter);
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(104, 112);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(512, 21);
			this.txtNombre.TabIndex = 352;
			// 
			// txtRuc
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance4;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(104, 80);
			this.txtRuc.MaxLength = 13;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 21);
			this.txtRuc.TabIndex = 7;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// dtFecha
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance5;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(104, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 5;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(520, 18);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 372;
			this.label1.Text = "Numero";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// txtNumero
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance6;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(592, 16);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 21);
			this.txtNumero.TabIndex = 371;
			this.txtNumero.ValueChanged += new System.EventHandler(this.txtNumero_ValueChanged);
			// 
			// cmbBancos
			// 
			this.cmbBancos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBancos.DataSource = this.ultraDataSource10;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBancos.DisplayMember = "Descripcion";
			this.cmbBancos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBancos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBancos.Enabled = false;
			this.cmbBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBancos.Location = new System.Drawing.Point(104, 144);
			this.cmbBancos.Name = "cmbBancos";
			this.cmbBancos.Size = new System.Drawing.Size(216, 22);
			this.cmbBancos.TabIndex = 10;
			this.cmbBancos.ValueMember = "idCuenta";
			this.cmbBancos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBancos_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn4,
																																	ultraDataColumn5});
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(16, 147);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(36, 16);
			this.label26.TabIndex = 379;
			this.label26.Text = "Banco";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(361, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 381;
			// 
			// txtidCompra
			// 
			this.txtidCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.Location = new System.Drawing.Point(585, 272);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(16, 21);
			this.txtidCompra.TabIndex = 388;
			this.txtidCompra.Visible = false;
			// 
			// txtIdOrigen
			// 
			this.txtIdOrigen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdOrigen.Enabled = false;
			this.txtIdOrigen.Location = new System.Drawing.Point(601, 272);
			this.txtIdOrigen.Name = "txtIdOrigen";
			this.txtIdOrigen.PromptChar = ' ';
			this.txtIdOrigen.Size = new System.Drawing.Size(16, 21);
			this.txtIdOrigen.TabIndex = 387;
			this.txtIdOrigen.Visible = false;
			// 
			// txtIdDepositosPendientes
			// 
			this.txtIdDepositosPendientes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdDepositosPendientes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdDepositosPendientes.Enabled = false;
			this.txtIdDepositosPendientes.Location = new System.Drawing.Point(617, 272);
			this.txtIdDepositosPendientes.Name = "txtIdDepositosPendientes";
			this.txtIdDepositosPendientes.PromptChar = ' ';
			this.txtIdDepositosPendientes.Size = new System.Drawing.Size(16, 21);
			this.txtIdDepositosPendientes.TabIndex = 386;
			this.txtIdDepositosPendientes.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(633, 272);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 385;
			this.txtEstado.Visible = false;
			// 
			// txtBodegaPredefinida
			// 
			this.txtBodegaPredefinida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredefinida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredefinida.Enabled = false;
			this.txtBodegaPredefinida.Location = new System.Drawing.Point(649, 272);
			this.txtBodegaPredefinida.Name = "txtBodegaPredefinida";
			this.txtBodegaPredefinida.PromptChar = ' ';
			this.txtBodegaPredefinida.Size = new System.Drawing.Size(16, 21);
			this.txtBodegaPredefinida.TabIndex = 384;
			this.txtBodegaPredefinida.Visible = false;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(665, 272);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdCliente.TabIndex = 383;
			this.txtIdCliente.Visible = false;
			// 
			// txtIdAnticipo
			// 
			this.txtIdAnticipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAnticipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAnticipo.Enabled = false;
			this.txtIdAnticipo.Location = new System.Drawing.Point(681, 272);
			this.txtIdAnticipo.Name = "txtIdAnticipo";
			this.txtIdAnticipo.PromptChar = ' ';
			this.txtIdAnticipo.Size = new System.Drawing.Size(16, 21);
			this.txtIdAnticipo.TabIndex = 382;
			this.txtIdAnticipo.Visible = false;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(16, 176);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(76, 16);
			this.lblNumero.TabIndex = 390;
			this.lblNumero.Text = "Num Deposito";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumDeposito
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumDeposito.Appearance = appearance7;
			this.txtNumDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumDeposito.Enabled = false;
			this.txtNumDeposito.Location = new System.Drawing.Point(104, 176);
			this.txtNumDeposito.MaxLength = 25;
			this.txtNumDeposito.Name = "txtNumDeposito";
			this.txtNumDeposito.Size = new System.Drawing.Size(128, 21);
			this.txtNumDeposito.TabIndex = 11;
			this.txtNumDeposito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumDeposito_KeyDown);
			this.txtNumDeposito.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumDeposito_KeyPress);
			this.txtNumDeposito.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumDeposito_Validating);
			this.txtNumDeposito.ValueChanged += new System.EventHandler(this.txtNumDeposito_ValueChanged);
			// 
			// btnSubir
			// 
			this.btnSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSubir.CausesValidation = false;
			this.btnSubir.Enabled = false;
			this.btnSubir.Image = ((System.Drawing.Image)(resources.GetObject("btnSubir.Image")));
			this.btnSubir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubir.Location = new System.Drawing.Point(368, 337);
			this.btnSubir.Name = "btnSubir";
			this.btnSubir.Size = new System.Drawing.Size(80, 23);
			this.btnSubir.TabIndex = 391;
			this.btnSubir.Text = "&Subir";
			this.btnSubir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSubir.Click += new System.EventHandler(this.btnSubir_Click);
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(232, 288);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(72, 16);
			this.linkLabel1.TabIndex = 392;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "Documento";
			this.linkLabel1.Visible = false;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(632, 337);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 4;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click_2);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(456, 337);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 23);
			this.btnImprimir.TabIndex = 394;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(360, 147);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 395;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 320);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(720, 8);
			this.groupBox1.TabIndex = 396;
			this.groupBox1.TabStop = false;
			// 
			// txtIdAutorizacion
			// 
			this.txtIdAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAutorizacion.Enabled = false;
			this.txtIdAutorizacion.Location = new System.Drawing.Point(456, 272);
			this.txtIdAutorizacion.Name = "txtIdAutorizacion";
			this.txtIdAutorizacion.PromptChar = ' ';
			this.txtIdAutorizacion.Size = new System.Drawing.Size(15, 21);
			this.txtIdAutorizacion.TabIndex = 1064;
			this.txtIdAutorizacion.Visible = false;
			// 
			// txtProntoPago
			// 
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProntoPago.Appearance = appearance8;
			this.txtProntoPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProntoPago.Enabled = false;
			this.txtProntoPago.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtProntoPago.FormatString = "#,##0.00";
			this.txtProntoPago.Location = new System.Drawing.Point(472, 272);
			this.txtProntoPago.MinValue = -2147483647;
			this.txtProntoPago.Name = "txtProntoPago";
			this.txtProntoPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProntoPago.PromptChar = ' ';
			this.txtProntoPago.Size = new System.Drawing.Size(16, 22);
			this.txtProntoPago.TabIndex = 1063;
			this.txtProntoPago.Visible = false;
			// 
			// txtSaldoVencido
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoVencido.Appearance = appearance9;
			this.txtSaldoVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoVencido.Enabled = false;
			this.txtSaldoVencido.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoVencido.FormatString = "#,##0.00";
			this.txtSaldoVencido.Location = new System.Drawing.Point(488, 272);
			this.txtSaldoVencido.MinValue = -2147483647;
			this.txtSaldoVencido.Name = "txtSaldoVencido";
			this.txtSaldoVencido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoVencido.PromptChar = ' ';
			this.txtSaldoVencido.Size = new System.Drawing.Size(14, 22);
			this.txtSaldoVencido.TabIndex = 1062;
			this.txtSaldoVencido.Visible = false;
			// 
			// txtSaldoTotal
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoTotal.Appearance = appearance10;
			this.txtSaldoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoTotal.Enabled = false;
			this.txtSaldoTotal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoTotal.FormatString = "#,##0.00";
			this.txtSaldoTotal.Location = new System.Drawing.Point(504, 272);
			this.txtSaldoTotal.MinValue = 0;
			this.txtSaldoTotal.Name = "txtSaldoTotal";
			this.txtSaldoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoTotal.PromptChar = ' ';
			this.txtSaldoTotal.Size = new System.Drawing.Size(14, 22);
			this.txtSaldoTotal.TabIndex = 1061;
			this.txtSaldoTotal.Visible = false;
			// 
			// txtDiasVencidos
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasVencidos.Appearance = appearance11;
			this.txtDiasVencidos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasVencidos.Enabled = false;
			this.txtDiasVencidos.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDiasVencidos.FormatString = "";
			this.txtDiasVencidos.Location = new System.Drawing.Point(520, 272);
			this.txtDiasVencidos.MinValue = -2147483647;
			this.txtDiasVencidos.Name = "txtDiasVencidos";
			this.txtDiasVencidos.PromptChar = ' ';
			this.txtDiasVencidos.Size = new System.Drawing.Size(14, 22);
			this.txtDiasVencidos.TabIndex = 1060;
			this.txtDiasVencidos.Visible = false;
			// 
			// txtSaldoAFavor
			// 
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoAFavor.Appearance = appearance12;
			this.txtSaldoAFavor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoAFavor.Enabled = false;
			this.txtSaldoAFavor.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoAFavor.FormatString = "#,##0.00";
			this.txtSaldoAFavor.Location = new System.Drawing.Point(536, 272);
			this.txtSaldoAFavor.MinValue = -2147483647;
			this.txtSaldoAFavor.Name = "txtSaldoAFavor";
			this.txtSaldoAFavor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoAFavor.PromptChar = ' ';
			this.txtSaldoAFavor.Size = new System.Drawing.Size(14, 22);
			this.txtSaldoAFavor.TabIndex = 1059;
			this.txtSaldoAFavor.Visible = false;
			// 
			// txtProximaVencer
			// 
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProximaVencer.Appearance = appearance13;
			this.txtProximaVencer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProximaVencer.Enabled = false;
			this.txtProximaVencer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtProximaVencer.FormatString = "";
			this.txtProximaVencer.Location = new System.Drawing.Point(544, 272);
			this.txtProximaVencer.MinValue = 0;
			this.txtProximaVencer.Name = "txtProximaVencer";
			this.txtProximaVencer.PromptChar = ' ';
			this.txtProximaVencer.Size = new System.Drawing.Size(14, 22);
			this.txtProximaVencer.TabIndex = 1058;
			this.txtProximaVencer.Visible = false;
			// 
			// txtSiguienteCuota
			// 
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSiguienteCuota.Appearance = appearance14;
			this.txtSiguienteCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSiguienteCuota.Enabled = false;
			this.txtSiguienteCuota.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSiguienteCuota.FormatString = "#,##0.00";
			this.txtSiguienteCuota.Location = new System.Drawing.Point(560, 272);
			this.txtSiguienteCuota.MinValue = 0;
			this.txtSiguienteCuota.Name = "txtSiguienteCuota";
			this.txtSiguienteCuota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSiguienteCuota.PromptChar = ' ';
			this.txtSiguienteCuota.Size = new System.Drawing.Size(14, 22);
			this.txtSiguienteCuota.TabIndex = 1057;
			this.txtSiguienteCuota.Visible = false;
			// 
			// txtCuotaVencida
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaVencida.Appearance = appearance15;
			this.txtCuotaVencida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaVencida.Enabled = false;
			this.txtCuotaVencida.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCuotaVencida.FormatString = "";
			this.txtCuotaVencida.Location = new System.Drawing.Point(576, 272);
			this.txtCuotaVencida.MinValue = 0;
			this.txtCuotaVencida.Name = "txtCuotaVencida";
			this.txtCuotaVencida.PromptChar = ' ';
			this.txtCuotaVencida.Size = new System.Drawing.Size(14, 22);
			this.txtCuotaVencida.TabIndex = 1056;
			this.txtCuotaVencida.Visible = false;
			// 
			// gridPagos
			// 
			this.gridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridPagos.CausesValidation = false;
			this.gridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			appearance16.BackColor = System.Drawing.Color.White;
			this.gridPagos.DisplayLayout.Appearance = appearance16;
			this.gridPagos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 105;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 76;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 53;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance17;
			ultraGridColumn9.Header.Caption = "Cuota";
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 43;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 107;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 51;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance18;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Valor";
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Width = 47;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance19;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Total Interes";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Width = 42;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance20;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Interés";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 25;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance21;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Total Capital";
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn16.Width = 34;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance22;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Width = 25;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance23;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Gestión";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.MaskInput = "";
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 28;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance24;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "% Dscto";
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.PromptChar = ' ';
			ultraGridColumn20.Width = 66;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance25;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 15;
			ultraGridColumn21.MaskInput = "";
			ultraGridColumn21.PromptChar = ' ';
			ultraGridColumn21.Width = 29;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance26;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "% Dscto";
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn22.Width = 72;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance27;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn23.Width = 33;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance28;
			ultraGridColumn24.Header.VisiblePosition = 22;
			ultraGridColumn24.Width = 42;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 57;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance29;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 20;
			ultraGridColumn26.Width = 35;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Pago";
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 66;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "Documento N.";
			ultraGridColumn28.Header.VisiblePosition = 24;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 87;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.Caption = "Forma de pago";
			ultraGridColumn29.Header.VisiblePosition = 26;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 107;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 27;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 108;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "Local";
			ultraGridColumn31.Header.VisiblePosition = 25;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 183;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.Header.VisiblePosition = 28;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 207;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance30;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "T Diferimiento";
			ultraGridColumn33.Header.VisiblePosition = 9;
			ultraGridColumn33.Width = 27;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance31;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 16;
			ultraGridColumn34.Width = 32;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 29;
			ultraGridColumn35.Width = 69;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			ultraGridBand3.Header.Caption = "";
			ultraGridBand3.HeaderVisible = true;
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand3.Override.ActiveRowAppearance = appearance32;
			ultraGridBand3.SummaryFooterCaption = "";
			this.gridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.gridPagos.DisplayLayout.Override.CardAreaAppearance = appearance33;
			appearance34.ForeColor = System.Drawing.Color.Black;
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.gridPagos.DisplayLayout.Override.CellAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance35.FontData.BoldAsString = "True";
			appearance35.FontData.Name = "Arial";
			appearance35.FontData.SizeInPoints = 8F;
			appearance35.ForeColor = System.Drawing.Color.White;
			appearance35.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.gridPagos.DisplayLayout.Override.HeaderAppearance = appearance35;
			appearance36.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance36.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.RowAlternateAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance38;
			this.gridPagos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.gridPagos.Location = new System.Drawing.Point(104, 208);
			this.gridPagos.Name = "gridPagos";
			this.gridPagos.Size = new System.Drawing.Size(600, 64);
			this.gridPagos.TabIndex = 1065;
			this.gridPagos.Visible = false;
			// 
			// frmDepositosPendientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(720, 366);
			this.Controls.Add(this.txtIdAutorizacion);
			this.Controls.Add(this.txtProntoPago);
			this.Controls.Add(this.txtSaldoVencido);
			this.Controls.Add(this.txtSaldoTotal);
			this.Controls.Add(this.txtDiasVencidos);
			this.Controls.Add(this.txtSaldoAFavor);
			this.Controls.Add(this.txtProximaVencer);
			this.Controls.Add(this.txtSiguienteCuota);
			this.Controls.Add(this.txtCuotaVencida);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.btnSubir);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblIce);
			this.Controls.Add(this.txtNumDeposito);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.txtIdOrigen);
			this.Controls.Add(this.txtIdDepositosPendientes);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtBodegaPredefinida);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.txtIdAnticipo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbBancos);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.txtDetalle);
			this.Controls.Add(this.txtAbono);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.gridPagos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmDepositosPendientes";
			this.Text = "Depositos Pendientes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDepositosPendientes_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDepositosPendientes_Closing);
			this.Load += new System.EventHandler(this.frmDepositosPendientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAbono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdDepositosPendientes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredefinida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAFavor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProximaVencer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSiguienteCuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaVencida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridPagos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bNuevo = false;
		bool bEdicion = false;
		bool bImpreso = false;
		bool bAbierto = false;
		private Acceso miAcceso;
		decimal dValor = 0.00m;
		decimal dSaldo = 0.00m;		
		#endregion Variables

		string sCarpeta = "";
		string sUrlNube = "";
		string txtDirectorioarchivo ="";

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmDepositosPendientes_Load(object sender, System.EventArgs e)
		{
			#region seguridad
			miAcceso = new Acceso(cdsSeteoF, "0738");

			if (!Funcion.Permiso("1861", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Depósitos Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			#endregion seguridad

			#region variables
			this.dtFecha.Enabled=false;
			this.cmbBodega.Enabled=false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnCliente.Enabled = false;
			this.btnSubir.Enabled=false;
			#endregion variables

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			this.btnCancelar_Click(sender, e);
						
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			this.txtBodegaPredefinida.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
							
			this.cmbBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCuentasDepositos "));

			bNuevo = true;
			bEdicion = true;
		}

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region inicializaVariables
			this.txtIdAnticipo.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdDepositosPendientes.Value = 0;
			this.txtidCompra.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbBancos.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtDetalle.Text = "";
			this.txtAbono.Value = 0.00m;
			this.txtIdOrigen.Value = 0;
			this.txtNumDeposito.Value = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			#endregion inicializaVariables

			#region variablesVisibles
			this.dtFecha.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbBancos.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtDetalle.Enabled = false;
			this.txtAbono.Enabled = false;			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnCliente.Enabled = false;
			this.txtNumDeposito.Enabled=false;
			this.linkLabel1.Visible = false;

			#endregion variablesVisibles

			this.btnAnular.Text = "Anular";
						
			bBloqueado = false;

			if (bCobroManual) this.DialogResult = DialogResult.No;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Value = 57;
			else this.cmbBodega.Value = (int)this.txtBodegaPredefinida.Value;
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
				
			this.txtEstado.Value = 1;
			this.lblEstado.Text = "PENDIENTE";

			if ((int)this.txtBodegaPredefinida.Value == 0) this.cmbBodega.Enabled = true;						
			this.txtRuc.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			//this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnCliente.Enabled = true;
						
			if (this.cmbBodega.Enabled) this.cmbBodega.Focus();
			else if (this.dtFecha.Enabled) this.dtFecha.Focus();

			bNuevo = true;
			bEdicion = true;		
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 45, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;
			
			#region validaciones
			bool bCredito = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec validaCreditosPendientes '{0}'", this.txtRuc.Text.ToString()), true);
			if (!bCredito)
			{
				MessageBox.Show("El cliente no cuenta con créditos abiertos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion validaciones

			using (frmLiquidaDepositosPendientes LDP = new frmLiquidaDepositosPendientes(this.txtRuc.Text.ToString(),1))
			{
				if (DialogResult.OK == LDP.ShowDialog())
				{
					if (LDP.GridFacturas.Rows.Count > 0)
					{
						this.txtidCompra.Value =(int)LDP.GridFacturas.ActiveRow.Cells["idCompra"].Value;
						this.txtIdCliente.Value = (int)LDP.GridFacturas.ActiveRow.Cells["idCliente"].Value;
						this.txtNombre.Value = LDP.txtNombre.Value;
						this.cmbBodega.Enabled=false;
						this.txtRuc.Enabled=false;

						this.cmbBancos.Enabled = true;			
						this.txtNumDeposito.Enabled = true;
						this.txtDetalle.Enabled = true;
						this.txtAbono.Enabled = true;
						if(miAcceso.BSubirArchivos) this.btnSubir.Enabled = true;
						this.btnGuardar.Enabled = true;

					}	
				}
				else this.btnCancelar_Click(sender, e);
			}

			//FuncionesProcedimientos.AbreFacturacionClientes(5, this.txtRuc.Text.ToString(), false, miAcceso.EditarCuentaNombre);
		}

		public void  NuevoGoogles (string sDirectorio)
		{
			#region variables
			string carpeta_1="";
			string carpeta_2="";
			string carpeta_3="";
			string carpeta ="";
			string txtURL ="";
			#endregion variables

			#region Editar google 
			#region carpetas
			carpeta_1 = @"\\192.168.1.10\Proyectos\DepositosPendientes";
			carpeta_2 = @"\\192.168.1.10\Proyectos\DepositosPendientes\"+MenuLatinium.stDirFacturacion;

			carpeta = @"\\192.168.1.10\Proyectos\DepositosPendientes\"+MenuLatinium.stDirFacturacion+@"\" + this.txtRuc.Text.Replace(" ", "_") +"_"+this.txtNumDeposito.Text+"_"+ Convert.ToString (this.txtidCompra.Value);
			carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + this.txtRuc.Text +"_"+this.txtNumDeposito.Text+"_"+ Convert.ToString (this.txtidCompra.Value);
			#endregion carpetas
			try
			{
				#region  sube Archivo
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
				#region directorios
				if (!Directory.Exists(carpeta_1))
				{
					Directory.CreateDirectory(carpeta);
					
				}
				if (!Directory.Exists(carpeta_2))
				{
					Directory.CreateDirectory(carpeta);
					
				}

				if (!Directory.Exists(carpeta))
				{
					Directory.CreateDirectory(carpeta);
					
				}
		#endregion directorios
				txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",sDirectorio,@"\")); 
				carpeta  = (carpeta + @"\" + txtURL.Replace(" ", "_"));
          
				System.IO.File.Copy (  txtDirectorioarchivo ,carpeta, false);
				sCarpeta = carpeta;
				sUrlNube = carpeta_3.Replace(" ", "_")+ @"\" + txtURL.Replace(" ", "_");
				MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
				#endregion  sube Archivo
								
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);

			}
			#endregion Editar google 
		}

		int idCompraPago = 0;
		bool bNormalAtrasado = false;	


		private string ConsumirAPI(string urlBase)
		{
			string respuesta = "";

			try
			{
				// Construir la URL con el parámetro
				string urlCompleta = urlBase ;

				// Crear la solicitud
				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlCompleta);
				request.Method = "GET";
				request.Timeout = 10000; // 10 segundos (ajustable)

				// Obtener la respuesta
				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					if (response.StatusCode == HttpStatusCode.OK)
					{
						using (StreamReader reader = new StreamReader(response.GetResponseStream()))
						{
							respuesta = reader.ReadToEnd();
						}
					}
					else
					{
						MessageBox.Show("Error en el servicio: " + response.StatusCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return "Error";
					}
				}
			}
			catch (WebException ex)
			{
				MessageBox.Show("Error de conexión al servicio:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "Error";
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error inesperado:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "Error";
			}

			return "OK";
		}



		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int sel = 0;

			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Cobros", cdsSeteoF)) return;
			if (!Validacion.vbTexto(this.txtRuc, 10, 10, "Cedula")) return;
			if (!Validacion.vbIdentificacion(this.txtRuc, 5, cdsSeteoF)) return;
			if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, 1, cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
			if (!Validacion.vbComboVacio(this.cmbBancos, "Seleccione un Banco")) return;
			if (!Validacion.vbTexto(this.txtNumDeposito, 5, 20, "Deposito")) return;

			if (this.txtNumDeposito.Text.ToString().Length > 0)
			{
				if (this.txtNumDeposito.Text.ToString().Substring(0, 1) == "0")
				{
					MessageBox.Show("El numero de depositos no puede tener ceros a la izquierda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumDeposito.Focus();
					return;
				}
			}

			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeroDepositoPendiente {0}, '{1}', 2", (int)this.cmbBancos.Value, this.txtNumDeposito.Text.ToString()), true))
			{
				MessageBox.Show("El Número de Documento ya se encuentra registrado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtNumDeposito.Focus();
				return;
			}

			if (!Validacion.vbTexto(this.txtDetalle, 10, 300, "Detalle")) return;
			if (!Validacion.vbCampoDecimalVacio(this.txtAbono, "Ingrese el Valor", 0.01m, 2500.00m)) return;

			if (sUrlNube.Length < 10)
			{
				MessageBox.Show("Seleccione el documento a subir.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaTiempoDepositosIngresados {0}", (int)this.txtIdCliente.Value), true) > 0)
				if (DialogResult.No == MessageBox.Show("Del cliente se ha ingresado un deposito en los ultimos 5 dias.\n\n ¿Esta seguro de ingresar otro deposito?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) return;
			#endregion Validacion

			if (bNuevo)
			{
				if (DialogResult.Yes == MessageBox.Show("¿Está seguro de Guardar? ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					#region Api 

					string url = "http://192.168.1.95:3055/api/v1/point/DepositosPendientes/"+ MenuLatinium.stDirFacturacion + "/"  + (string)this.txtRuc.Text+@"/"+sCarpeta;

					//string url = "http://192.168.1.95:3055/api/v1/point/DepartamentoTecnico/archivo_20251001_160636_508.png/archivo_20251001_160636_508.png/archivo_20251001_160636_508.png";  // La URL base
					string jsonRespuesta = ConsumirAPI(url);

					if (jsonRespuesta == "Error")
					{
						MessageBox.Show("Error en la subida de Imagenes \n\n  Comunicate con el área de Desarrollo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					// Luego puedes parsear el JSON como te mostré antes.
					//MessageBox.Show("Respuesta:\n" + jsonRespuesta, "API", MessageBoxButtons.OK, MessageBoxIcon.Information);
        
					#endregion Api
			
					string baseUrl = "https://storage.googleapis.com/sparta_bucket/Despositos/"+ MenuLatinium.stDirFacturacion + "/"  + (string)this.txtRuc.Text+@"/"+sCarpeta;
					sUrlNube = baseUrl;
				}
				else
				{
					return;
				
				}
			}

			#region Guarda Anticipo
			try
			{
				#region Numeracion
				if (bNuevo)
				{
					string sSQLNumero = string.Format("Exec NumeracionLocales 83, {0}, 0", (int)this.cmbBodega.Value);
					this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQLNumero, true);
				}
				#endregion Numeracion

				#region Guarda Depositos Pendientes
				DateTime dtFecha = (DateTime)this.dtFecha.Value;

				string sSQLGraba = string.Format(
					"Exec GrabaDepositosPendientes {0}, '{1}', {2}, '{3}', {4}, {5}, '{6}', {7}, {8}, {9}, {10}, '{11}', {12}, '{13}'",
					(int)this.txtIdDepositosPendientes.Value,
					dtFecha.ToString("yyyyMMdd"),
					(int)this.cmbBodega.Value,
					this.txtNumero.Text.ToString(),
					(int)this.txtIdCliente.Value,
					(int)this.cmbBancos.Value,
					this.txtDetalle.Text.ToString(),
					Convert.ToDecimal(this.txtAbono.Value),
					(int)this.txtEstado.Value,
					(int)this.txtidCompra.Value,
					(int)this.txtIdAnticipo.Value,
					this.txtDetalle.Text.ToString(),
					this.txtNumDeposito.Text.ToString(),
          sUrlNube
					);
				this.txtIdDepositosPendientes.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGraba, true);
				int idCompraPago = 0;     
				int idDepositosPorliquida = 0;				
				#endregion Guarda Depositos Pendientes

				#region Graba Google
				/*if (!bNuevo)
				{
					if (txtDirectorioarchivo.Length > 10)
					{
						NuevoGoogle(txtDirectorioarchivo);
					}
				}
				if (sUrlNube.Length > 10)
				{
					string sSQLGoogle = string.Format("Exec grabaUrlNubeDepositos {0}, '{1}', '{2}'", (int)this.txtIdDepositosPendientes.Value, sCarpeta, sUrlNube);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLGoogle, true);
				}*/
				#endregion Graba Google

				#region Validación Cruce
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDepositosCruzar '{0}','{1}'", this.txtNumDeposito.Value, this.txtAbono.Value)) == 1 && (int)this.cmbBancos.Value == 12)
				{
					#region Cruze
					this.Cursor = Cursors.WaitCursor;

					#region proceso
					#region Variables
					decimal Cobro =  Convert.ToDecimal(this.txtAbono.Value);
					decimal Diferimiento = 0.00m;
					decimal ABDiferimiento = 0.00m;
					decimal DescDiferimiento = 0.00m;
					decimal Gasto = 0.00m;
					decimal IntMora = 0.00m;
					decimal IntFin = 0.00m;
					decimal Capital = 0.00m;
					decimal Abono = 0.00m;
					decimal Saldo = Convert.ToDecimal(this.txtAbono.Value);
					decimal NSaldo = 0.00m;
					decimal AbGastos = 0.00m;
					decimal AbIntMora = 0.00m;
					decimal AbInteres = 0.00m;
					decimal AbCapital = 0.00m;
					int Estado = 0;
					string sConcepto = "";
					string sVacio = "";
					int idCompraOrigen = Convert.ToInt32(this.txtidCompra.Value);
					#endregion Variables

					#region Genera Abono Factura 24
					string sNumero = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 24, 57, 0");

					if (Cobro < Convert.ToDecimal(this.txtSaldoVencido.Value)) bNormalAtrasado = true;
					string sSQLPago = string.Format(
						"Exec GuardaPagoCobranzaDepositosPendientes {0}, {1}, {2}, 0, '{3}', '{4}', {5}, {6}",
						idCompraOrigen,57,Math.Round(Convert.ToDecimal(Cobro), 2), Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"),sNumero,bNormalAtrasado,iFormaPago);			
					idCompraPago = Funcion.iEscalarSQL(cdsSeteoF, sSQLPago, true);
					#endregion Genera Abono Factura 24

					#region Guarda Cobro de la factura
					int idBanco = 7;
					int idTarjeta = 0;
					int idCuenta = (int)this.cmbBancos.Value;
					int idPlan = 0;
					int idPlazo = 0;
					decimal dComTarjeta = 0.00m;
					DateTime dtFecCaducidad = DateTime.Today;

					string sSQLCobro = string.Format(
						"Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0",
						0, idCompraPago, 3, iFormaPago, Cobro,	idBanco, idTarjeta,sVacio, (string)this.txtNumDeposito.Value.ToString(),
						dtFecCaducidad.ToString("yyyyMMdd"),(string)this.txtNumDeposito.Value.ToString(),	sVacio, idCuenta, sVacio, sVacio, idPlan, idPlazo, dComTarjeta);					
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
					#endregion Guarda Cobro de la factura

					#region Calculo de Cobranza

					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
					{
						if (Saldo > 0)
						{
							Gasto = 0.00m;
							IntMora = 0.00m;
							IntFin = 0.00m;
							Capital = 0.00m;
							AbGastos = 0.00m;
							AbIntMora = 0.00m;
							AbInteres = 0.00m;
							AbCapital = 0.00m;
							Diferimiento = 0.00m;
							ABDiferimiento = 0.00m;
							DescDiferimiento = 0.00m;
							Estado = 0;

							Diferimiento = Convert.ToDecimal(dr.Cells["Diferimiento"].Value);
							IntMora = Convert.ToDecimal(dr.Cells["Mora"].Value);
							IntFin = Convert.ToDecimal(dr.Cells["Interes"].Value);
							Capital = Convert.ToDecimal(dr.Cells["Capital"].Value);
							Gasto = Convert.ToDecimal(dr.Cells["Gastos"].Value);
							Abono = 0.00m;

							#region Diferimiento
							if (Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2))
							{
								if (Convert.ToDateTime(dr.Cells["Vence"].Value) >= Convert.ToDateTime(this.dtFecha.Value) && (Capital + IntFin + IntMora + Gasto) <= Saldo)
								{
									ABDiferimiento += Diferimiento;
									DescDiferimiento += Diferimiento;
								}
								else
								{
									Abono += Diferimiento;
									ABDiferimiento += Diferimiento;
									Saldo -= Diferimiento;
								}
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2))
							{
								Diferimiento -= Saldo;
								Abono += Saldo;
								ABDiferimiento += Saldo;
								Saldo = 0.00m;
							}
							#endregion Diferimiento

							#region Gestión de cobranza (Gastos)
							if (Math.Round(Gasto, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Gasto, 2))
							{
								Abono += Gasto;
								AbGastos += Gasto;
								Saldo -= Gasto;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Gasto, 2) > 0.00m && Math.Round(Gasto, 2) >= Math.Round(Saldo, 2))
							{
								Gasto -= Saldo;
								Abono += Saldo;
								AbGastos += Saldo;
								Saldo = 0.00m;
							}
							#endregion Gestión de cobranza

							#region Interés por mora
							if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
							{
								Abono += IntMora;
								AbIntMora += IntMora;
								Saldo -= IntMora;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
							{
								IntMora -= Saldo;
								Abono += Saldo;
								AbIntMora += Saldo;
								Saldo = 0.00m;
							}
							#endregion Interés por mora

							#region Interés por financiamiento
							if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
							{
								Abono += IntFin;
								AbInteres += IntFin;
								Saldo -= IntFin;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
							{
								IntFin -= Saldo;
								Abono += Saldo;
								AbInteres += Saldo;
								Saldo = 0.00m;
							}
							#endregion Interés por financiamiento

							#region Capital
							if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
							{
								Abono += Capital;
								AbCapital += Capital;
								Saldo -= Capital;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
							{
								Capital -= Saldo;
								Abono += Saldo;
								AbCapital += Saldo;
								Saldo = 0.00m;
							}
							#endregion Capital

							#region Actualiza tabla Cre_TablaDeAmortización
							if (Abono != 0.00m)
							{
								decimal dSaldoPrb = 0.00m;
								dSaldoPrb = Convert.ToDecimal(dr.Cells["Saldo"].Value);
								NSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value) - Abono - ABDiferimiento;

								string sSQLAC = string.Format(
									"Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}",
									(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value,
									Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"), idCompraPago, AbInteres, AbCapital, AbGastos, AbIntMora, ABDiferimiento, DescDiferimiento);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAC, true);

								#region Estado Sorteo
								sConcepto = "";

								if (Math.Round(NSaldo, 2) <= 0.00m)
								{
									Estado = 2;
									EstadoSorteo = 1;
									EstadoCelular = 1;
									sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();

									int num_cuota = Convert.ToInt32(dr.Cells["NumeroCuota"].Value.ToString());
								}
								else if (Math.Round(NSaldo, 2) > 0.00m)
								{
									Estado = 3;
									sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
								}
								#endregion Estado Sorteo

								string sSQLTipoPago = string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
								int iPago = Funcion.iEscalarSQL(cdsSeteoF, sSQLTipoPago, true);

								#region  Descuento COVID (conservado como 0 si no aplica)
								decimal dDescGestion = 0.00m;
								decimal dDesMora = 0.00m;
								#endregion  Descuento COVID

								decimal dDebitado = AbIntMora + AbInteres + AbCapital + ABDiferimiento;

								decimal dRescate = 0.00m;
								string sSQLRescate = string.Format("Exec Cre_ValidaSaldoCuotaRescatada {0}, {1}, '{2}'",
									(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, dDebitado, Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"));
								dRescate = Funcion.decEscalarSQL(cdsSeteoF, sSQLRescate);

								string sSQLPCC = string.Format(
									"Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}, {17}, {18}, {19}",
									idCompraPago, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, AbGastos, AbIntMora, AbInteres, AbCapital, Abono,
									Math.Round(NSaldo, 2), Estado, sConcepto, Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"), iPago, dDebitado, NumeroDeOperacion,
									0, dDescGestion, dDesMora, ABDiferimiento, DescDiferimiento, dRescate);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLPCC, true);

								#region Actualiza Pago en gestiones
								decimal dValorPagadoGestiones = AbCapital + AbInteres;
								string sSQLAPG = string.Format(
									"Exec Cre_ActualizaValorPagoGestiones {0}, '{1}', '{2}', {3}, {4}",
									idCompraOrigen, Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"), Convert.ToDateTime(dr.Cells["Vence"].Value).ToString("yyyyMMdd"), dValorPagadoGestiones, Abono);							
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAPG, true);
								#endregion Actualiza Pago en gestiones
							}
							#endregion Actualiza tabla Cre_TablaDeAmortización
						}
					}
					#endregion Calculo de Cobranza

					#region Historial de autorizaciones
					if ((int)this.txtIdAutorizacion.Value > 0)
					{
						string sSQLHA = string.Format("Exec CobranzaActualizaEstadoAutorizaciones {0}", this.txtIdAutorizacion.Value);						
						Funcion.EjecutaSQL(cdsSeteoF, sSQLHA, true);
					}
					#endregion Historial de autorizaciones

					#region Asiento - Saldos
					string sSQLCarteraExterna = string.Format("Exec GeneraMovimientosCarteraExterna {0}, {1}", idCompraPago,Convert.ToInt32((int)this.txtidCompra.Value));					
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCarteraExterna, true);

					string sSQLSaldoA = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 9", idCompraPago);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldoA, true);

					string sSQLSaldoB = string.Format("Exec ActualizaSaldoFacturas {0}, 1, 9", idCompraPago);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldoB, true);
					#endregion Asiento - Saldos

					#region Actualiza Compromisos de Pago - Estado de Cuenta
					string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}",
						idCompraOrigen,Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"),Saldo);				
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActCompromisoPago, true);

					string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", idCompraOrigen);				
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAEC, true);
					#endregion Actualiza Compromisos de Pago - Estado de Cuenta

					#region Notificacion Pago
					string sSQLNotificaPago = string.Format("Exec GuardaNotifiacionCuotaPagos {0}, {1}, {2}, {3}, {4}",
						Convert.ToInt32((int)this.txtidCompra.Value), idCompraPago, 0, 0, 1);				
					Funcion.EjecutaSQL(cdsSeteoF, sSQLNotificaPago, true);
					#endregion Notificacion Pago

					#region Libera celulares
					string sCountCel = string.Format("Select COUNT(1) From dbo.Cre_TablaDeAmortizacion Where Estado Not In (2, 4) and idCompra = {0}", idCompraOrigen);					
					int iValEstadoCel = Funcion.iEscalarSQL(cdsSeteoF, sCountCel, true);
					if (EstadoCelular == 1)
					{
						string sSQLBLO = string.Format("Exec LiberacionDeCelulares {0}", idCompraOrigen);						
						Funcion.EjecutaSQL(cdsSeteoF, sSQLBLO, true);
					}
					#endregion Libera celulares

					#region Actualiza estado Depositos Pendientes
					string sUpdateDP = string.Format("Update DepositosPendientes Set IdAnticipo = {0}, idDepositosPorLiquidar = {1}, Estado = 4 , Sel = {2} Where IdDepositosPendientes = {3}",
						idCompraPago, idDepositosPorliquida, sel,  Convert.ToInt32((int)this.txtIdDepositosPendientes.Value));				
					Funcion.EjecutaSQL(cdsSeteoF, sUpdateDP, true);
					#endregion Actualiza estado Depositos Pendientes

					#region Actualiza estado Depositos Por Liquidar
					string sUpdateDPL = string.Format("Update DepositosPorLiquidar Set idCompra = {0}, Estado = 2 Where idDepositosPorLiquidar = {1}",
						idCompraPago, idDepositosPorliquida);					
					Funcion.EjecutaSQL(cdsSeteoF, sUpdateDPL, true);
					#endregion Actualiza estado Depositos Por Liquidar

					#endregion proceso

					this.Cursor = Cursors.Default;
					#endregion Cruze
				}
				#endregion Validación Cruce

				#region Controles
				this.cmbBodega.Enabled = false;
				this.cmbBancos.Enabled = false;
				this.dtFecha.Enabled = false;
				this.txtRuc.Enabled = false;
				this.txtDetalle.Enabled = false;
				this.txtAbono.Enabled = false;
				this.btnSubir.Enabled = false;
				this.txtNumDeposito.Enabled = false;
				this.linkLabel1.Visible = false;
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				if (miAcceso.Editar) this.btnEditar.Enabled = true;
				if (miAcceso.Anular) this.btnAnular.Enabled = true;
				this.btnGuardar.Enabled = false;
				this.btnCliente.Enabled = false;

				bNuevo = false;
				bEdicion = false;
				#endregion Controles
			}
			catch (System.Exception ex)
			{				
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			#endregion Guarda Anticipo
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaDepositosPendientes Busqueda = new frmConsultaDepositosPendientes((int)this.txtBodegaPredefinida.Value, 0))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridInformacion.ActiveRow.Cells["IdDepositosPendientes"].Value);
				}
			}
		}
		private void Consulta(int IdDepositosPendientes)
		{
			try
			{
				#region controles 
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = true;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;	
				this.btnImprimir.Enabled = true;	
				#endregion controles

				int IBodega = 0;
				int IBanco = 0;
				string numeroDeposito ="";

				#region consulta individual
				string sSQLBusqueda = string.Format("Exec BuscaDepositosPendientesIndividual {0}", IdDepositosPendientes);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLBusqueda, true);
				dr.Read();
				if (dr.HasRows)
				{
					#region asigna variables 
					this.txtIdDepositosPendientes.Value = dr.GetInt32(0);						
					this.dtFecha.Value = dr.GetDateTime(1);
					this.txtNumero.Text = dr.GetValue(2).ToString();
					IBodega = dr.GetInt32(3);			
					this.txtIdCliente.Value = dr.GetInt32(4);
					this.txtRuc.Text = dr.GetValue(5).ToString();
					this.txtNombre.Text = dr.GetValue(6).ToString();
					this.txtDetalle.Text = dr.GetValue(7).ToString();
					this.txtEstado.Value = dr.GetInt32(8);
					this.lblEstado.Text = dr.GetValue(9).ToString();
					this.txtAbono.Value = dr.GetDecimal(10);
					this.txtIdAnticipo.Value = dr.GetInt32(11);
					this.txtidCompra.Value = dr.GetInt32(12);
					IBanco = dr.GetInt32(13);	
					if (dr.GetValue(14) != System.DBNull.Value) sCarpeta = dr.GetString(14).ToString();
					if (dr.GetValue(15) != System.DBNull.Value) sUrlNube = dr.GetString(15).ToString();
					if (sUrlNube.Length > 10)
					{
						this.linkLabel1.Visible = true;
						this.btnSubir.Enabled = false;
					}
					numeroDeposito = dr.GetValue(16).ToString();
					//this.txtNumDeposito.Text = dr.GetValue(16).ToString();
					this.lblUsuario.Text = dr.GetValue(17).ToString();
					#endregion asigna variables 
				}
				dr.Close();
				#endregion consulta individual

			  #region Estados
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
				if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "VERIFICADO";
				if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "ANULADO";
				if ((int)this.txtEstado.Value == 4) this.lblEstado.Text = "CRUZADO";
				#endregion Estados
	
				this.cmbBodega.Value = IBodega;
				this.cmbBancos.Value = IBanco;
				this.txtNumDeposito.Text = numeroDeposito;
				
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;

				this.btnCancelar.Enabled = true;							
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdDepositosPendientes.Value);

			#region valida estado
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El Documento esta Cruzado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El Documento esta  Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}						
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CobroValidaEstadoCartera {0}", (int)this.txtIdAnticipo.Value)) > 0)
			{
				MessageBox.Show("No puede Editar Cartera En transito o Pagada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			#endregion valida estado

			#region controles
			this.txtRuc.Enabled = false;
			this.txtDetalle.Enabled = true;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnCliente.Enabled = false;
			this.cmbBodega.Enabled=false;
			this.cmbBancos.Enabled=true;
			this.txtAbono.Enabled=true;
			this.txtNumDeposito.Enabled=true;
			#endregion controles

			bNuevo = false;
			bEdicion = true;
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnCliente.Focus();
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (!Validacion.vbTexto(this.txtRuc, 10, 10, "Cedula")) e.Cancel = true;	
								
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, 1, cdsSeteoF)) e.Cancel = true;	

				if (!Validacion.vbIdentificacion(this.txtRuc, 5, cdsSeteoF)) e.Cancel = true;
			}			
		}

		private void btnSalir_Click_1(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtNumDeposito_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			#region Validacion
			if (bEdicion)
			{				
				if (!Validacion.vbTexto(this.txtNumDeposito, 5, 20, "Deposito")) e.Cancel = true;	
			
				if (this.txtNumDeposito.Text.ToString().Length > 0)
				{
					if (this.txtNumDeposito.Text.ToString().Substring(0, 1) == "0")
					{
						MessageBox.Show("El numero de depositos no puede tener ceros a la izquierda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtNumDeposito.Focus();
						e.Cancel = true;
					}
				}

				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeroDepositoPendiente {0}, '{1}', 2", (int)this.cmbBancos.Value, this.txtNumDeposito.Text.ToString()), true))
				{
					MessageBox.Show("El Número de Documento ya se encuentra registrado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumDeposito.Focus();
					e.Cancel = true;
				}
			}
			#endregion Validacion
		}

		public string NuevoGoogle(string txtDirectorioarchivo )
		{
			if ( !File.Exists(txtDirectorioarchivo))
			{
				MessageBox.Show("El archivo seleccionado no existe.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}

			try
			{
				// Configuración de la ruta destino
				string carpetaDestino = @"\\192.168.1.10\Proyectos\DepositosPendientes\" 
					+ MenuLatinium.stDirFacturacion + @"\" 
					+ (string)this.txtRuc.Text;

     
				// Crear carpeta si no existe
				if (!Directory.Exists(carpetaDestino))
				{
					Directory.CreateDirectory(carpetaDestino);
				}

				// Obtener info del archivo original
				FileInfo fi = new FileInfo(txtDirectorioarchivo);
				string extension = fi.Extension;

				// Generar nombre único con timestamp
				string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
				string nuevoNombreArchivo = "archivo_" + timestamp + extension;
			  sCarpeta = nuevoNombreArchivo;



				string rutaDestinoFinal = Path.Combine(carpetaDestino, nuevoNombreArchivo);

				// Si ya existe un archivo con ese nombre, lo elimina (altamente improbable por timestamp)
				if (File.Exists(rutaDestinoFinal))
				{
					File.Delete(rutaDestinoFinal);
				}

				// Conexión a red (forzar acceso)
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				Process.Start("cmd.exe", "/C net use \\\\192.168.1.10 /user:latinium 10dgpv15.p");

				// Mover archivo
				File.Copy(txtDirectorioarchivo, rutaDestinoFinal);

				// Confirmación opcional
				return rutaDestinoFinal;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
		}



		private void btnSubir_Click(object sender, System.EventArgs e)
		{
			if (bNuevo)
			{
				#region Validacion
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Cobros", cdsSeteoF)) return;
			
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

				if (!Validacion.vbComboVacio(this.cmbBancos, "Seleccione un Banco")) return;
			
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 1, 1, cdsSeteoF)) return;	

				if (!Validacion.vbTexto(this.txtDetalle, 10, 300, "Detalle")) return;
			
				if (!Validacion.vbCampoDecimalVacio(this.txtAbono, "Ingrese el Valor", dCobro, 100000.00m)) return;
				#endregion Validacion


				/* Cargar imagen */
				OpenFileDialog ofd = new OpenFileDialog();
				ofd.InitialDirectory = @"C:\";
				ofd.Filter = "Archivos de imagen (*.jpg;*.jpeg;*.png;*.gif)|*.jpg;*.jpeg;*.png;*.gif";

				if (ofd.ShowDialog() == DialogResult.OK)
				{
					string archivoSeleccionado = ofd.FileName;
				

					// Llama al método y obtiene la nueva ruta
					string nuevaRuta = NuevoGoogle(archivoSeleccionado);

					if ( nuevaRuta == null || nuevaRuta == "")
					{
					
						MessageBox.Show("Hubo un problema al subir la imagen.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						sUrlNube = nuevaRuta;
					}
				}
				else
				{
					MessageBox.Show("No se seleccionó una imagen.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}

			}
			
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (sUrlNube.Length > 10 )
			{
				Process.Start(sUrlNube);

				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
			}
		}

		private void btnSalir_Click_2(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			Facturacion.ImpresionRecibos((int)this.txtIdAnticipo.Value, 2, cdsSeteoF);

			this.btnImprimir.Enabled = false;
			
			bImpreso = true;
			Cursor = Cursors.Default;
		}

		private void txtNumDeposito_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			this.btnCliente_Click(sender, e);
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void txtNumDeposito_ValueChanged(object sender, System.EventArgs e)
		{
			#region cruce cuenta
			/*
			try
			{	
				string numDep = "";
				if (this.txtNumDeposito.Value != null)
					numDep = this.txtNumDeposito.Value.ToString().Trim();

				int idCuenta = 0;
				if (this.cmbBancos.Value != null && this.cmbBancos.Value.ToString() != "")
				{
					try { idCuenta = Convert.ToInt32(this.cmbBancos.Value); } 
					catch { idCuenta = 0; }
				}

				if (numDep == "" || idCuenta <= 0)
				{
					dtFechaDeposito = DateTime.MinValue;
					try { this.txtAbono.Value = 0.00m; } 
					catch { this.txtAbono.Text = "0.00"; }
					this.txtAbono.Enabled = true; 
					return;
				}

				dtFechaDeposito = DateTime.MinValue;
				try { this.txtAbono.Value = 0.00m; } 
				catch { this.txtAbono.Text = "0.00"; }
				this.txtAbono.Enabled = true; 

				string sSQL = string.Format(
					"SELECT TOP 1 CAST(Fecha AS DATETIME) AS Fecha, Monto " +
					"FROM Cont_EstadoDeCuenta " +
					"WHERE Documento = '{0}' AND idCuenta = {1} " +
					"ORDER BY Fecha DESC",
					numDep.Replace("'", "''"), idCuenta);

				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);

				bool encontrado = false;

				if (dr.Read())
				{
					DateTime fecha = dr.IsDBNull(0) ? DateTime.MinValue : dr.GetDateTime(0);
					decimal monto = dr.IsDBNull(1) ? 0.00m : dr.GetDecimal(1);

					if (fecha != DateTime.MinValue)
					{
						dtFechaDeposito = fecha;
						encontrado = true;
						try { this.txtAbono.Value = monto; }
						catch { this.txtAbono.Text = monto.ToString("0.00"); }
					}
				}
				dr.Close();

				if (encontrado)
				{
					this.txtAbono.Enabled = false;

					if ((int)this.txtidCompra.Value > 0)
					{
						ConsultaFacturas((int)this.txtidCompra.Value);
						Cobranza(dtFechaDeposito);
						ConsultaTablaDeAmortizacion((int)this.txtidCompra.Value, dtFechaDeposito);
					}
				}
				else
				{
					dtFechaDeposito = DateTime.MinValue;
					try { this.txtAbono.Value = 0.00m; } 
					catch { this.txtAbono.Text = "0.00"; }
					this.txtAbono.Enabled = true;
				}
			}
			catch (System.Exception ex)
			{
				MessageBox.Show("Error consultando depósito: " + ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error);
			}
			*/
			#endregion cruce cuenta
		}
		private void txtAbono_Enter(object sender, System.EventArgs e)
		{
			this.txtAbono.SelectAll();
		}

		private void frmDepositosPendientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{				
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdDepositosPendientes.Value > 0)
				{	
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (miAcceso.Imprimir && !bImpreso) this.btnImprimir_Click(sender, e);
					if (e.KeyCode == Keys.F5) this.Consulta((int)this.txtIdDepositosPendientes.Value);
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void frmDepositosPendientes_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un registro de deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.txtRuc.Focus();
		}

		private void cmbBancos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumDeposito.Focus();
		}

		private void txtNumDeposito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDetalle.Focus();
		}

		private void txtDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtAbono.Focus();
		}

		private void txtAbono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}
		private void ConsultaFacturas(int iIdCompra)
		{
			#region Valores Estado de Cuenta
			string sSQL = string.Format("Exec ConsultaFacturasDepositosPendientes {0}", iIdCompra);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if(dr.HasRows)
			{
				iIdCompra = (int)dr.GetInt32(0);	
				iIdEntidadFinanciera =(int)dr.GetInt32(1);	
				ValorCuotas = (decimal)dr.GetDecimal(2);
				NumeroLote = (int)dr.GetInt32(3);
				Confirmado =(bool)dr.GetBoolean(4);	
				NumeroDeOperacion = (string)dr.GetString(5);
				Reliquido= (bool)dr.GetBoolean(6);
				ER = (int)dr.GetInt32(7);
				Asignado = (bool)dr.GetBoolean(8);
			}
			dr.Close();
			#endregion Valores Estado de Cuenta
					
		}
		private void Cobranza(DateTime dtFechaPago)
		{	
			try
			{
											
				#region Valores Estado de Cuenta
				string sSQL = string.Format("Exec ValoresCobranza {0}, '{1}'", (int)this.txtidCompra.Value,Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"));
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if(dr.HasRows)
				{
					this.txtCuotaVencida.Value = dr.GetInt32(1);
					this.txtSiguienteCuota.Value = dr.GetDecimal(3);
					this.txtProximaVencer.Value = dr.GetInt32(4);
					iDiasVencidos = (int)dr.GetInt32(5); 
					this.txtSaldoAFavor.Value = dr.GetDecimal(6);
				}
				dr.Close();
				#endregion Valores Estado de Cuenta
					
				if (iDiasVencidos >= 0) this.txtDiasVencidos.Value = iDiasVencidos; 
				else this.txtDiasVencidos.Value = 0;

				string sSQLUFP = string.Format("Exec CobranzasUltimaFechaPago {0}", (int)this.txtidCompra.Value);
				dtUltimaFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, sSQLUFP);

				ConsultaTablaDeAmortizacion((int)this.txtidCompra.Value,dtFechaDeposito);	
					
				#region Valor re liquidacion
				decimal dGC = 0.00m;
				decimal dIM = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPagos in this.gridPagos.Rows.All)
				{
					dGC = dGC + Convert.ToDecimal(drPagos.Cells["Gastos"].Value);
					dIM = dIM + Convert.ToDecimal(drPagos.Cells["Mora"].Value);
				}

				sSQL = string.Format("Exec ConsultaValorReliquidacion {0}, '{1}', {2}, {3}, 1", 
					this.txtidCompra.Value,Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"), dGC, dIM);
				SqlDataReader drR = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drR.Read();
				if(drR.HasRows)
				{
					this.txtProntoPago.Value = drR.GetDecimal(0);
					dInteres = drR.GetDecimal(3);
				}
				drR.Close();
				#endregion Valor re liquidacion
									
				#region Valida Saldo Total
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dRow in this.gridPagos.Rows.All)
				{
					this.txtSaldoTotal.Value = Convert.ToDecimal(this.txtSaldoTotal.Value) + Convert.ToDecimal(dRow.Cells["Saldo"].Value);	
					
					DateTime Inicia = (DateTime)dRow.Cells["Inicia"].Value;
					Vence = (DateTime)dRow.Cells["Vence"].Value;
					
					if (Convert.ToDateTime(dtFechaPago) > Vence)
						this.txtSaldoVencido.Value = 
							Math.Round(Convert.ToDecimal(this.txtSaldoVencido.Value), 2) + Convert.ToDecimal(dRow.Cells["Saldo"].Value);							
						
					if (Convert.ToDateTime(dtFechaPago)== Vence )
						this.txtSaldoVencido.Value = 
							Math.Round(Convert.ToDecimal(this.txtSaldoVencido.Value), 2) + Convert.ToDecimal(dRow.Cells["Saldo"].Value) 
							- Convert.ToDecimal(dRow.Cells["Diferimiento"].Value);						
				}
				#endregion Valida Saldo Total
			}
			catch (Exception exc)
			{
				MessageBox.Show(string.Format("Exception Type: {0} {1}", exc.GetType(), exc.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
			}			
		}
		private void ConsultaTablaDeAmortizacion(int idCompra,DateTime dtFechaPago)
		{
			string sSQL = string.Format("Exec Cobranzas {0}, '{1}'", 
				idCompra, Convert.ToDateTime(dtFechaDeposito).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.gridPagos);			
		}

		private void txtAbono_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

	}
}
