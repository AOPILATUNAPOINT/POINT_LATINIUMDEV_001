using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCierreDeCaja.
	/// </summary>
	public class frmCierreDeCaja : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNota;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCheque;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnEfectivo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCCheque;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCEfectivo;
		private C1.Data.C1DataSet cdsSeteo;
		private Acceso miAcceso;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRevisado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private System.ComponentModel.IContainer components;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private System.Windows.Forms.Label lblEstadoDepositos;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox2;

		private C1DataRow drSeteoF;
		private System.Windows.Forms.Button btnVer;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnDepositar;
		private System.Windows.Forms.Button btnActualizar;
		int iBodega = 0;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		int idCaja = 0;
		decimal dEfectivo = 0.00m;
		private System.Windows.Forms.Button btnExportar;
		decimal dCheque = 0.00m;
		decimal valorDepositos = 0.00m;
		int iEstadoCaja = 0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnC_Entrega;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private System.Windows.Forms.Button btnAceptar;
		int iEstado = 0;
    
		int iBodegaR = 0;
		DateTime FechaCajaBefore;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCTC;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorriente;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago_1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		DateTime dtFechaR = DateTime.Today;
		string sArchivoCierreDeCaja = "CierreDeCaja.rpt";


		public frmCierreDeCaja(int IBodegaR , DateTime DtFechaR  )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			iBodegaR = IBodegaR;
			dtFechaR = DtFechaR;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}
		public frmCierreDeCaja()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCierreDeCaja));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucCedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Voucher");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeUna");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonoDeUna");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Registro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago_1", -1, "cmbFormaPago_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonoEntrada");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCaja", 0);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTarjetas", 1);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVerTarjeta", 2);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnRetencionVenta", 3);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDeuna", 4);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 16, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 17, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 19, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Voucher", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Voucher", 20, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 21, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 24, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 22, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 23, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 33, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 18, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 25, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 26, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 27, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 30, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 29, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 28, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BonoEntrada", 39, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BonoEntrada", 39, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DeUna", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DeUna", 31, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "BonoDeUna", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "BonoDeUna", 32, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Revisado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RucCedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Voucher");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DeUna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoDeUna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Registro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoEntrada");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BtnDeposito", 0);
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 6, true);
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.spnTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnNota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnCheque = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.spnCCheque = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnCEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkRevisado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.lblEstadoDepositos = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnDepositar = new System.Windows.Forms.Button();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btnExportar = new System.Windows.Forms.Button();
			this.spnC_Entrega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.spnCTC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbFormaPago_1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnC_Entrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCTC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago_1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			ultraGridColumn1.CellAppearance = appearance2;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			ultraGridColumn1.CellButtonAppearance = appearance3;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 17;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.Caption = "I";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn2.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn2.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn2.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn2.Width = 17;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.Caption = "T";
			ultraGridColumn3.Header.VisiblePosition = 5;
			ultraGridColumn3.Width = 17;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 6;
			ultraGridColumn4.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn4.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn4.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn4.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.RowLayoutColumnInfo.OriginX = 4;
			ultraGridColumn6.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn6.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn6.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 9;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 39;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 32;
			ultraGridColumn11.Hidden = true;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Header.Caption = "N. Documento";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.RowLayoutColumnInfo.OriginX = 6;
			ultraGridColumn12.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn12.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn12.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn12.Width = 85;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "RUC - Cedula";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.RowLayoutColumnInfo.OriginX = 8;
			ultraGridColumn14.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn14.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn14.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.RowLayoutColumnInfo.OriginX = 10;
			ultraGridColumn15.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn15.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn15.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn15.Width = 200;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.RowLayoutColumnInfo.OriginX = 12;
			ultraGridColumn16.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn16.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn16.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn16.Width = 75;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance6;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.RowLayoutColumnInfo.OriginX = 14;
			ultraGridColumn17.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn17.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn17.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn17.Width = 75;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance7;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.RowLayoutColumnInfo.OriginX = 16;
			ultraGridColumn18.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn18.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn18.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn18.Width = 75;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance8;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Transf";
			ultraGridColumn19.Header.VisiblePosition = 23;
			ultraGridColumn19.RowLayoutColumnInfo.OriginX = 24;
			ultraGridColumn19.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn19.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn19.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn19.Width = 75;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance9;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "T/Cred";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.RowLayoutColumnInfo.OriginX = 18;
			ultraGridColumn20.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn20.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn20.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn20.Width = 75;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance10;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.RowLayoutColumnInfo.OriginX = 22;
			ultraGridColumn21.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn21.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn21.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn21.Width = 75;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance11;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "N De Cred";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.RowLayoutColumnInfo.OriginX = 22;
			ultraGridColumn22.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn22.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn22.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn22.Width = 75;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance12;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn23.RowLayoutColumnInfo.OriginX = 28;
			ultraGridColumn23.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn23.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn23.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn23.Width = 75;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance13;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "C/Point";
			ultraGridColumn24.Header.VisiblePosition = 26;
			ultraGridColumn24.RowLayoutColumnInfo.OriginX = 30;
			ultraGridColumn24.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn24.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn24.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn24.Width = 75;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance14;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Crédito";
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.RowLayoutColumnInfo.OriginX = 32;
			ultraGridColumn25.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn25.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn25.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn25.Width = 75;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance15;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "Retención";
			ultraGridColumn26.Header.VisiblePosition = 27;
			ultraGridColumn26.RowLayoutColumnInfo.OriginX = 34;
			ultraGridColumn26.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn26.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn26.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn26.Width = 75;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance16;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "T/Deb";
			ultraGridColumn27.Header.VisiblePosition = 28;
			ultraGridColumn27.RowLayoutColumnInfo.OriginX = 20;
			ultraGridColumn27.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn27.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn27.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn27.Width = 75;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance17;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 29;
			ultraGridColumn28.RowLayoutColumnInfo.OriginX = 26;
			ultraGridColumn28.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn28.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn28.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn28.Width = 75;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance18;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 30;
			ultraGridColumn29.Width = 75;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance19;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "C/Entrega";
			ultraGridColumn30.Header.VisiblePosition = 31;
			ultraGridColumn30.Width = 85;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance20;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 33;
			ultraGridColumn31.RowLayoutColumnInfo.OriginX = 36;
			ultraGridColumn31.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn31.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn31.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn31.Width = 75;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance21;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 36;
			ultraGridColumn32.Width = 75;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance22;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 37;
			ultraGridColumn33.Width = 75;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance23;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 38;
			ultraGridColumn34.RowLayoutColumnInfo.OriginX = 38;
			ultraGridColumn34.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn34.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn34.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn34.Width = 85;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Format = "dd/MM/yyy HH:mm";
			ultraGridColumn35.Header.VisiblePosition = 42;
			ultraGridColumn35.Width = 120;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn36.Header.VisiblePosition = 43;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.Caption = "Vendedor";
			ultraGridColumn37.Header.VisiblePosition = 44;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn37.Width = 200;
			ultraGridColumn38.Header.VisiblePosition = 40;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn39.Header.Caption = "Forma De Pago";
			ultraGridColumn39.Header.VisiblePosition = 41;
			ultraGridColumn39.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn39.Width = 250;
			ultraGridColumn40.AutoEdit = false;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance24;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellButtonAppearance = appearance25;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.Caption = "Bono E. Crédito";
			ultraGridColumn40.Header.VisiblePosition = 34;
			ultraGridColumn40.Width = 104;
			ultraGridColumn41.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance26.Image = ((object)(resources.GetObject("appearance26.Image")));
			appearance26.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance26.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn41.CellButtonAppearance = appearance26;
			ultraGridColumn41.Header.Caption = "...";
			ultraGridColumn41.Header.VisiblePosition = 3;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn41.Width = 25;
			ultraGridColumn42.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance27.Image = ((object)(resources.GetObject("appearance27.Image")));
			appearance27.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance27.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn42.CellButtonAppearance = appearance27;
			ultraGridColumn42.Header.Caption = "...";
			ultraGridColumn42.Header.VisiblePosition = 4;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn42.Width = 25;
			ultraGridColumn43.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance28.Image = ((object)(resources.GetObject("appearance28.Image")));
			appearance28.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellButtonAppearance = appearance28;
			ultraGridColumn43.Header.Caption = "...";
			ultraGridColumn43.Header.VisiblePosition = 20;
			ultraGridColumn43.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn43.Width = 30;
			ultraGridColumn44.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance29.Image = ((object)(resources.GetObject("appearance29.Image")));
			appearance29.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn44.CellButtonAppearance = appearance29;
			ultraGridColumn44.Header.Caption = "R";
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn44.Width = 30;
			ultraGridColumn45.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn45.Header.Caption = "Ver";
			ultraGridColumn45.Header.VisiblePosition = 35;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn45.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
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
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			appearance30.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance30;
			ultraGridBand1.Header.Caption = "MOVIMIENTOS DE CAJA";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance31;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance32;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance33;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance34;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.Hidden = true;
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance35;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance36;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance37;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance38;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance39;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance40;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance41;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance42;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance43;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance44;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance45;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance46;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance47;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance48;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance49;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14,
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19});
			ultraGridBand1.SummaryFooterCaption = "Total Movimientos de Caja";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance50.ForeColor = System.Drawing.Color.Black;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance50;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance51.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance51;
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance52;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance53.FontData.BoldAsString = "True";
			appearance53.FontData.Name = "Arial";
			appearance53.FontData.SizeInPoints = 8F;
			appearance53.ForeColor = System.Drawing.Color.White;
			appearance53.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance53;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance55;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1264, 296);
			this.ultraGrid1.TabIndex = 14;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.DateTime);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.DateTime);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40});
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
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
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 173;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance56;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 0;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 180;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(240, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(224, 21);
			this.cmbBodega.TabIndex = 1;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(192, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 174;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// spnTotal
			// 
			this.spnTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnTotal.Appearance = appearance57;
			this.spnTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnTotal.Enabled = false;
			this.spnTotal.FormatString = "#,##0.00";
			this.spnTotal.Location = new System.Drawing.Point(1186, 528);
			this.spnTotal.Name = "spnTotal";
			this.spnTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTotal.PromptChar = ' ';
			this.spnTotal.Size = new System.Drawing.Size(80, 22);
			this.spnTotal.TabIndex = 182;
			this.spnTotal.Visible = false;
			// 
			// spnTarjeta
			// 
			this.spnTarjeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnTarjeta.Appearance = appearance58;
			this.spnTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnTarjeta.Enabled = false;
			this.spnTarjeta.FormatString = "#,##0.00";
			this.spnTarjeta.Location = new System.Drawing.Point(866, 528);
			this.spnTarjeta.Name = "spnTarjeta";
			this.spnTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnTarjeta.PromptChar = ' ';
			this.spnTarjeta.Size = new System.Drawing.Size(80, 22);
			this.spnTarjeta.TabIndex = 181;
			this.spnTarjeta.Visible = false;
			// 
			// spnNota
			// 
			this.spnNota.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnNota.Appearance = appearance59;
			this.spnNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnNota.Enabled = false;
			this.spnNota.FormatString = "#,##0.00";
			this.spnNota.Location = new System.Drawing.Point(946, 528);
			this.spnNota.Name = "spnNota";
			this.spnNota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnNota.PromptChar = ' ';
			this.spnNota.Size = new System.Drawing.Size(80, 22);
			this.spnNota.TabIndex = 180;
			this.spnNota.Visible = false;
			// 
			// spnCheque
			// 
			this.spnCheque.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCheque.Appearance = appearance60;
			this.spnCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCheque.Enabled = false;
			this.spnCheque.FormatString = "#,##0.00";
			this.spnCheque.Location = new System.Drawing.Point(786, 528);
			this.spnCheque.Name = "spnCheque";
			this.spnCheque.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCheque.PromptChar = ' ';
			this.spnCheque.Size = new System.Drawing.Size(80, 22);
			this.spnCheque.TabIndex = 179;
			this.spnCheque.Visible = false;
			// 
			// spnEfectivo
			// 
			this.spnEfectivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnEfectivo.Appearance = appearance61;
			this.spnEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnEfectivo.Enabled = false;
			this.spnEfectivo.FormatString = "#,##0.00";
			this.spnEfectivo.Location = new System.Drawing.Point(706, 528);
			this.spnEfectivo.Name = "spnEfectivo";
			this.spnEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnEfectivo.PromptChar = ' ';
			this.spnEfectivo.Size = new System.Drawing.Size(80, 22);
			this.spnEfectivo.TabIndex = 178;
			this.spnEfectivo.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn41.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn41,
																																 ultraDataColumn42});
			// 
			// ultraDataSource7
			// 
			ultraDataColumn43.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn43,
																																 ultraDataColumn44});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(1048, 7);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 221;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// spnCCheque
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCCheque.Appearance = appearance62;
			this.spnCCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCCheque.Enabled = false;
			this.spnCCheque.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnCCheque.FormatString = "#,##0.00";
			this.spnCCheque.Location = new System.Drawing.Point(952, 8);
			this.spnCCheque.Name = "spnCCheque";
			this.spnCCheque.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCCheque.PromptChar = ' ';
			this.spnCCheque.Size = new System.Drawing.Size(60, 21);
			this.spnCCheque.TabIndex = 227;
			this.spnCCheque.Visible = false;
			// 
			// spnCEfectivo
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCEfectivo.Appearance = appearance63;
			this.spnCEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCEfectivo.Enabled = false;
			this.spnCEfectivo.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnCEfectivo.FormatString = "#,##0.00";
			this.spnCEfectivo.Location = new System.Drawing.Point(1216, 8);
			this.spnCEfectivo.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnCEfectivo.Name = "spnCEfectivo";
			this.spnCEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCEfectivo.PromptChar = ' ';
			this.spnCEfectivo.Size = new System.Drawing.Size(60, 21);
			this.spnCEfectivo.TabIndex = 226;
			this.spnCEfectivo.Visible = false;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(System.DateTime);
			ultraDataColumn49.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn51.DataType = typeof(System.Double);
			ultraDataColumn51.DefaultValue = 0;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn52.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn52,
																																 ultraDataColumn53});
			// 
			// ultraDataSource6
			// 
			ultraDataColumn54.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn54,
																																 ultraDataColumn55});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(640, 472);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(632, 8);
			this.groupBox1.TabIndex = 239;
			this.groupBox1.TabStop = false;
			// 
			// chkRevisado
			// 
			this.chkRevisado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkRevisado.Appearance = appearance64;
			this.chkRevisado.BackColor = System.Drawing.Color.Transparent;
			this.chkRevisado.Location = new System.Drawing.Point(1084, 8);
			this.chkRevisado.Name = "chkRevisado";
			this.chkRevisado.Size = new System.Drawing.Size(72, 22);
			this.chkRevisado.TabIndex = 240;
			this.chkRevisado.Text = "Revisado";
			this.chkRevisado.Visible = false;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance65;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(768, 360);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(504, 104);
			this.txtObservaciones.TabIndex = 241;
			this.txtObservaciones.ValueChanged += new System.EventHandler(this.txtObservaciones_ValueChanged);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(int);
			ultraDataColumn58.DataType = typeof(int);
			ultraDataColumn59.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(System.DateTime);
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn62.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn63.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63});
			// 
			// ultraDataSource9
			// 
			ultraDataColumn64.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn64,
																																 ultraDataColumn65});
			// 
			// ultraDataSource10
			// 
			ultraDataColumn66.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn66,
																																	ultraDataColumn67});
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource8;
			appearance66.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance66;
			this.uGridDepositos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 78;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 54;
			ultraGridColumn51.Header.Caption = "Forma de pago";
			ultraGridColumn51.Header.VisiblePosition = 2;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn51.Width = 132;
			ultraGridColumn52.Header.Caption = "Banco";
			ultraGridColumn52.Header.VisiblePosition = 3;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn52.Width = 185;
			ultraGridColumn53.Header.Caption = "Número";
			ultraGridColumn53.Header.VisiblePosition = 4;
			ultraGridColumn53.Width = 90;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn54.Header.Caption = "Fecha";
			ultraGridColumn54.Header.VisiblePosition = 5;
			ultraGridColumn54.Width = 70;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance67;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 6;
			ultraGridColumn55.PromptChar = ' ';
			ultraGridColumn55.Width = 72;
			ultraGridColumn56.Header.VisiblePosition = 7;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 56;
			ultraGridColumn57.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance68.Image = ((object)(resources.GetObject("appearance68.Image")));
			appearance68.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn57.CellButtonAppearance = appearance68;
			ultraGridColumn57.Header.Caption = "Depositar";
			ultraGridColumn57.Header.VisiblePosition = 8;
			ultraGridColumn57.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn57.Width = 40;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57});
			ultraGridBand3.Header.Caption = "Depositos";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance69;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings20});
			ultraGridBand3.SummaryFooterCaption = "Total Depositos";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance70.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance70;
			appearance71.ForeColor = System.Drawing.Color.Black;
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDepositos.DisplayLayout.Override.CellAppearance = appearance71;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance72.FontData.BoldAsString = "True";
			appearance72.FontData.Name = "Arial";
			appearance72.FontData.SizeInPoints = 8F;
			appearance72.ForeColor = System.Drawing.Color.White;
			appearance72.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance73;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance74.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance74;
			this.uGridDepositos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(5, 360);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(627, 152);
			this.uGridDepositos.TabIndex = 250;
			this.uGridDepositos.AfterRowFilterChanged += new Infragistics.Win.UltraWinGrid.AfterRowFilterChangedEventHandler(this.uGridDepositos_AfterRowFilterChanged);
			this.uGridDepositos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDepositos_AfterRowInsert);
			this.uGridDepositos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDepositos_KeyPress);
			this.uGridDepositos.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDepositos_BeforeCellUpdate);
			this.uGridDepositos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDepositos_BeforeRowsDeleted);
			this.uGridDepositos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDepositos_KeyDown);
			this.uGridDepositos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDepositos_ClickCellButton);
			this.uGridDepositos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDepositos_AfterCellUpdate);
			this.uGridDepositos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDepositos_InitializeLayout);
			// 
			// lblEstadoDepositos
			// 
			this.lblEstadoDepositos.AutoSize = true;
			this.lblEstadoDepositos.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.lblEstadoDepositos.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstadoDepositos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoDepositos.Location = new System.Drawing.Point(576, 7);
			this.lblEstadoDepositos.Name = "lblEstadoDepositos";
			this.lblEstadoDepositos.Size = new System.Drawing.Size(0, 23);
			this.lblEstadoDepositos.TabIndex = 242;
			this.lblEstadoDepositos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(672, 376);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(76, 16);
			this.label3.TabIndex = 252;
			this.label3.Text = "Observaciones";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1288, 8);
			this.groupBox2.TabIndex = 253;
			this.groupBox2.TabStop = false;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(480, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 708;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(996, 488);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 719;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1188, 488);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 720;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnDepositar
			// 
			this.btnDepositar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDepositar.CausesValidation = false;
			this.btnDepositar.Image = ((System.Drawing.Image)(resources.GetObject("btnDepositar.Image")));
			this.btnDepositar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDepositar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnDepositar.Location = new System.Drawing.Point(668, 488);
			this.btnDepositar.Name = "btnDepositar";
			this.btnDepositar.Size = new System.Drawing.Size(80, 23);
			this.btnDepositar.TabIndex = 721;
			this.btnDepositar.Tag = "";
			this.btnDepositar.Text = "&Depositar";
			this.btnDepositar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDepositar.Visible = false;
			this.btnDepositar.Click += new System.EventHandler(this.btnDepositar_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(756, 488);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(136, 23);
			this.btnActualizar.TabIndex = 725;
			this.btnActualizar.Text = "Actualizar Numeración";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(900, 488);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(88, 23);
			this.btnGuardar.TabIndex = 726;
			this.btnGuardar.Text = "Cerrar Caja";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// cmbCuentaCorriente
			// 
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataSource = this.ultraDataSource10;
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.Caption = "";
			ultraGridColumn59.Header.VisiblePosition = 1;
			ultraGridColumn59.Width = 250;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn58,
																										 ultraGridColumn59});
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCuentaCorriente.DisplayMember = "Descripcion";
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(872, 216);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(272, 88);
			this.cmbCuentaCorriente.TabIndex = 246;
			this.cmbCuentaCorriente.ValueMember = "idCuenta";
			this.cmbCuentaCorriente.Visible = false;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource9;
			ultraGridColumn60.Header.VisiblePosition = 0;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.Caption = "";
			ultraGridColumn61.Header.VisiblePosition = 1;
			ultraGridColumn61.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn60,
																										 ultraGridColumn61});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(16, 400);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(224, 86);
			this.cmbFormaPago.TabIndex = 247;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1092, 488);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 727;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// spnC_Entrega
			// 
			appearance75.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnC_Entrega.Appearance = appearance75;
			this.spnC_Entrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnC_Entrega.Enabled = false;
			this.spnC_Entrega.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnC_Entrega.FormatString = "#,##0.00";
			this.spnC_Entrega.Location = new System.Drawing.Point(1152, 8);
			this.spnC_Entrega.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnC_Entrega.Name = "spnC_Entrega";
			this.spnC_Entrega.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnC_Entrega.PromptChar = ' ';
			this.spnC_Entrega.Size = new System.Drawing.Size(60, 21);
			this.spnC_Entrega.TabIndex = 728;
			this.spnC_Entrega.Visible = false;
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource11;
			ultraGridColumn62.Header.VisiblePosition = 0;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.Width = 200;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn62,
																										 ultraGridColumn63});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbVendedores.DisplayMember = "Nombre";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(1024, 200);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(216, 104);
			this.cmbVendedores.TabIndex = 729;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn68.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn68,
																																	ultraDataColumn69});
			// 
			// btnAceptar
			// 
			this.btnAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAceptar.Enabled = false;
			this.btnAceptar.Location = new System.Drawing.Point(640, 488);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(24, 24);
			this.btnAceptar.TabIndex = 730;
			this.btnAceptar.Text = "...";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// spnCTC
			// 
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCTC.Appearance = appearance76;
			this.spnCTC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCTC.Enabled = false;
			this.spnCTC.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnCTC.FormatString = "#,##0.00";
			this.spnCTC.Location = new System.Drawing.Point(1016, 8);
			this.spnCTC.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnCTC.Name = "spnCTC";
			this.spnCTC.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCTC.PromptChar = ' ';
			this.spnCTC.Size = new System.Drawing.Size(60, 21);
			this.spnCTC.TabIndex = 731;
			this.spnCTC.Visible = false;
			// 
			// cmbFormaPago_1
			// 
			this.cmbFormaPago_1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago_1.DataSource = this.ultraDataSource12;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 1;
			ultraGridColumn65.Width = 250;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn64,
																										 ultraGridColumn65});
			this.cmbFormaPago_1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbFormaPago_1.DisplayMember = "Descripcion";
			this.cmbFormaPago_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago_1.Location = new System.Drawing.Point(505, 216);
			this.cmbFormaPago_1.Name = "cmbFormaPago_1";
			this.cmbFormaPago_1.Size = new System.Drawing.Size(272, 88);
			this.cmbFormaPago_1.TabIndex = 732;
			this.cmbFormaPago_1.ValueMember = "idFormaPago_1";
			this.cmbFormaPago_1.Visible = false;
			// 
			// ultraDataSource12
			// 
			ultraDataColumn70.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn70,
																																	ultraDataColumn71});
			// 
			// frmCierreDeCaja
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1282, 520);
			this.Controls.Add(this.cmbFormaPago_1);
			this.Controls.Add(this.spnCTC);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.spnC_Entrega);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.btnDepositar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Controls.Add(this.lblEstadoDepositos);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.chkRevisado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.spnCCheque);
			this.Controls.Add(this.spnCEfectivo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.spnTotal);
			this.Controls.Add(this.spnTarjeta);
			this.Controls.Add(this.spnNota);
			this.Controls.Add(this.spnCheque);
			this.Controls.Add(this.spnEfectivo);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.uGridDepositos);
			this.Controls.Add(this.cmbVendedores);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCierreDeCaja";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cuadre De Caja";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCierreDeCaja_KeyDown);
			this.Load += new System.EventHandler(this.frmCierreDeCaja_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnC_Entrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCTC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago_1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		string strHostName = string.Empty;

		private void frmCierreDeCaja_Load(object sender, System.EventArgs e)
		{			
			Funcion.Resolucion(this, this.Height, this.Width);
		
			strHostName = System.Environment.MachineName;

			Seguridades();						
		
			FuncionesProcedimientos.EstadoGrids(false, this.uGridDepositos);
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			FuncionesProcedimientos.FechasDocumentos(this.dtFecha);

			if (iBodegaR == 0) 
			{
				string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
				iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

				if (iBodega > 0)
				{
					this.cmbBodega.Value = iBodega; 
					this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("select C.idCuenta, C.Descripcion from BodegaCuentasBancarias B  INNER JOIN Cuenta C ON C.idCuenta = B.idCuenta WHERE B.idBodega =  {0}", iBodega));			
				}
				else 
				{
					this.cmbBodega.Enabled = true;
					this.cmbBodega.Value = 57;
					this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("select C.idCuenta, C.Descripcion from BodegaCuentasBancarias B  INNER JOIN Cuenta C ON C.idCuenta = B.idCuenta WHERE B.idBodega =  {0}", (int) this.cmbBodega.Value));			
				}
				this.dtFecha.Enabled = true;
				//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;			
				FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

				//Versionamiento Archivo Cierre de caja
				//Funcion.VersionamientoReporte(sArchivoCierreDeCaja);

			}
			else
			{
				this.cmbBodega.Value = iBodegaR;
				this.dtFecha.Value = dtFechaR;	
				this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("select C.idCuenta, C.Descripcion from BodegaCuentasBancarias B  INNER JOIN Cuenta C ON C.idCuenta = B.idCuenta WHERE B.idBodega =  {0}", iBodegaR));			
			}
		}
		
		private void UnloadMe()
		{
			this.Close();
		}

		private void Seguridades()
		{
			miAcceso = new Acceso(cdsSeteoF, "0722");

			if (!Funcion.Permiso("66",cdsSeteo))
			{					
				MessageBox.Show("No tiene acceso a cuadre de caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
		
			if (miAcceso.B1DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);	
			//if (MenuLatinium.stDirFacturacion == "DENTALIS")
			//{
				if (miAcceso.CambiarFecha)
				{
					this.uGridDepositos.DisplayLayout.Bands[0].Columns["FechaDeposito"].CellActivation = Activation.AllowEdit;
					//this.uGridDepositos.DisplayLayout.Bands[0].Columns["FechaDeposito"].CellActivation = Activation.AllowEdit;

				}	
			//}
			if (miAcceso.B15DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-15);
			else if (miAcceso.B30DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-30);
			else if (miAcceso.B60DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-60);
			if (miAcceso.BExportar) this.btnExportar.Enabled = true;	
			if (!miAcceso.Nuevo) this.ultraGrid1.DisplayLayout.Bands[0].Columns["idPersonal"].CellActivation = Activation.Disabled; // ACCESO A MODIFICAR VENDEDOR
			if (miAcceso.Anular) this.btnAceptar.Enabled = true; // ACCESO A ASIGNAR LAS CUENTAS BANCARIAS	
			if (!miAcceso.BIrA)  this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnVerTarjeta"].Hidden = true;
			if (!miAcceso.BAPVerPagados)  this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnDeuna"].Hidden = true;
			if (!miAcceso.Retencion)this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnRetencionVenta"].Hidden = true;
			if (!miAcceso.EditarDetalleFormaPago) this.ultraGrid1.DisplayLayout.Bands[0].Columns["idFormaPago_1"].CellActivation = Activation.Disabled;			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		
		private void EstadoControles(bool Estado)
		{
			this.btnGuardar.Enabled = Estado;
			this.btnImprimir.Enabled = Estado;
		}

		private void CalculaSubTotal()
		{
			this.spnEfectivo.Value = 0.00m;
			this.spnCheque.Value = 0.00m;
						
			this.spnCEfectivo.Value = 0.00m;
			this.spnCCheque.Value = 0.00m;
			this.spnCTC.Value = 0.00m;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				decimal Efectivo = 0.00m;
				decimal Cheque = 0.00m;
				decimal TC = 0.00m;
				
				if (dr.Cells["Efectivo"].Value != DBNull.Value)
				{
					Efectivo = Convert.ToDecimal(dr.Cells["Efectivo"].Value);
					Cheque = Convert.ToDecimal(dr.Cells["Cheque"].Value);
					TC = Convert.ToDecimal(dr.Cells["Tarjeta"].Value)+ Convert.ToDecimal(dr.Cells["Debito"].Value) ;
					
					this.spnEfectivo.Value = Convert.ToDecimal(this.spnEfectivo.Value) + Efectivo;
					this.spnCheque.Value = Convert.ToDecimal(this.spnCheque.Value) + Cheque;
					
					this.spnCEfectivo.Value = Convert.ToDecimal(this.spnCEfectivo.Value) + Efectivo;
					this.spnCCheque.Value = Convert.ToDecimal(this.spnCCheque.Value) + Cheque;	
					this.spnCTC.Value	= Convert.ToDecimal(this.spnCTC.Value) + TC;		
				}
			}
		}

		private void CalculaTotal()
		{
			this.spnCEfectivo.Value = 0.00m;
			this.spnCCheque.Value = 0.00m;
			this.spnC_Entrega.Value = 0.00m;
			this.spnCTC.Value	= 0.00m;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				decimal Efectivo = 0.00m;
				decimal Cheque = 0.00m;
				decimal dC_Entrega = 0.00m;
				decimal TC = 0.00m;
				
				if (dr.Cells["Efectivo"].Value != DBNull.Value)
				{
					Efectivo = Convert.ToDecimal(dr.Cells["Efectivo"].Value);
					Cheque = Convert.ToDecimal(dr.Cells["Cheque"].Value);
					dC_Entrega = Convert.ToDecimal(dr.Cells["C_Entrega"].Value);
					TC = Convert.ToDecimal(dr.Cells["Tarjeta"].Value)+ Convert.ToDecimal(dr.Cells["Debito"].Value);
					
					this.spnCEfectivo.Value = Convert.ToDecimal(this.spnCEfectivo.Value) + Efectivo;
					this.spnCCheque.Value = Convert.ToDecimal(this.spnCCheque.Value) + Cheque;
					this.spnC_Entrega.Value = Convert.ToDecimal(this.spnC_Entrega.Value) + dC_Entrega;
					this.spnCTC.Value = Convert.ToDecimal(this.spnCTC.Value) + TC;

				}
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["Estado"].Value != DBNull.Value)
			{
				if ((string)e.Row.Cells["Estado"].Value == "MANUAL")
				{					
					e.Row.Appearance.BackColor = Color.Yellow;					
					e.Row.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.Rectangular;
				}
				else if ((string)e.Row.Cells["Estado"].Value == "ANULADO")
				{					
					e.Row.Appearance.BackColor = Color.Tomato;
					e.Row.Appearance.ForeColor = Color.Black;					
				}
				else if ((int)e.Row.Cells["idEntidadFinanciera"].Value == 14)
				{					
					e.Row.Appearance.BackColor = Color.Green;
					e.Row.Appearance.ForeColor = Color.Black;					
				}
			}
		}

		private void MensajeEfectivo()
		{
			#region Mensaje para Efectivo
			//			if (Math.Round(Convert.ToDecimal(this.spnCEfectivo.Value), 2) > 0.00m)
			//			{
			//				if (Math.Round(Convert.ToDecimal(this.txtValorEfectivo.Value), 2) == Math.Round(Convert.ToDecimal(this.spnCEfectivo.Value), 2))
			//				{
			//					this.lblEstadoEfectivo.Text = "EFECTIVO CUADRADO";
			//					this.lblEstadoEfectivo.ForeColor = Color.Green;
			//				}
			//				else if (Math.Round(Convert.ToDecimal(this.txtValorEfectivo.Value), 2) < Math.Round(Convert.ToDecimal(this.spnCEfectivo.Value), 2))
			//				{
			//					decimal Faltante = Math.Round(Convert.ToDecimal(this.spnCEfectivo.Value), 2) - Math.Round(Convert.ToDecimal(this.txtValorEfectivo.Value), 2);
			//					this.lblEstadoEfectivo.Text = string.Format("FALTANTE EN EFECTIVO DE {0}", Math.Round(Faltante, 2));
			//					this.lblEstadoEfectivo.ForeColor = Color.Red;
			//				}
			//				else if (Math.Round(Convert.ToDecimal(this.txtValorEfectivo.Value), 2) > Math.Round(Convert.ToDecimal(this.spnCEfectivo.Value), 2))
			//				{
			//					decimal Sobrante = Math.Round(Convert.ToDecimal(this.txtValorEfectivo.Value), 2) - Math.Round(Convert.ToDecimal(this.spnCEfectivo.Value), 2);
			//					this.lblEstadoEfectivo.Text = string.Format("SOBRANTE EN EFECTIVO DE {0}", Math.Round(Sobrante, 2));
			//					this.lblEstadoEfectivo.ForeColor = Color.Blue;
			//				}
			//			}
			#endregion Mensaje para Efectivo
		}

		private void MensajeCheque()
		{
			#region Mensaje para Cheque
			//			if (Math.Round(Convert.ToDecimal(this.spnCCheque.Value), 2) > 0.00m)
			//			{
			//				if (Math.Round(Convert.ToDecimal(this.txtValorCheque.Value), 2) == Math.Round(Convert.ToDecimal(this.spnCCheque.Value), 2))
			//				{
			//					this.lblEstadoCheque.Text = "CHEQUES CUADRADO";
			//					this.lblEstadoCheque.ForeColor = Color.Green;
			//				}
			//				else if (Math.Round(Convert.ToDecimal(this.txtValorCheque.Value), 2) < Math.Round(Convert.ToDecimal(this.spnCCheque.Value), 2))
			//				{
			//					decimal Faltante = Math.Round(Convert.ToDecimal(this.spnCCheque.Value), 2) - Math.Round(Convert.ToDecimal(this.txtValorCheque.Value), 2);
			//					this.lblEstadoCheque.Text = string.Format("FALTANTE EN CHEQUES DE {0}", Math.Round(Faltante, 2));
			//					this.lblEstadoCheque.ForeColor = Color.Red;
			//				}
			//				else if (Math.Round(Convert.ToDecimal(this.txtValorCheque.Value), 2) > Math.Round(Convert.ToDecimal(this.spnCCheque.Value), 2))
			//				{
			//					decimal Sobrante = Math.Round(Convert.ToDecimal(this.txtValorCheque.Value), 2) - Math.Round(Convert.ToDecimal(this.spnCCheque.Value), 2);
			//					this.lblEstadoCheque.Text = string.Format("SOBRANTE EN CHEQUES DE {0}", Math.Round(Sobrante, 2));
			//					this.lblEstadoCheque.ForeColor = Color.Blue;
			//				}
			//			}
			#endregion Mensaje para Cheque
		}
		
		private void GeneraAsiento()
		{		
			#region Crea Asiento
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{
				string sSQL = string.Format("Exec CreaAsiCierreCaja {0}, {1}", (int)dr.Cells["idDetalleCaja"].Value, (int)dr.Cells["idCaja"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}			

			string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCaja {0}", idCaja);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLDiferencias, true);
			#endregion Crea Asiento
		}
				
		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
 	
			if (e.KeyChar == 13) if (this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.btnVer_Click(sender, e);
		}
		
		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) e.Cancel = true;
		}
		
		private void frmCierreDeCaja_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		
		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow != null)
			{
				if (this.ultraGrid1.ActiveRow.Selected) 
				{
					if (this.ultraGrid1.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value) 
					{
						if (miAcceso.BContabilizarFactMan) 
						{
							if (this.ultraGrid1.ActiveRow.Cells["Estado"].Value.ToString() == "MANUAL")
							{
								int IdCompra = (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value;
								string sNumero = this.ultraGrid1.ActiveRow.Cells["Factura"].Value.ToString();

								using (frmArticulosSeriesManuales SerialesManuales = new frmArticulosSeriesManuales(IdCompra, sNumero))
								{				
									if (DialogResult.OK == SerialesManuales.ShowDialog())
									{
										string sSQL = string.Format("Exec SerialesManualesPorFactura {0}", IdCompra);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

										this.btnVer_Click(sender, e);
									}
								}				
							}
						}
					}
				}
			}
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}

		private void uGridDepositos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{		
			for (int i=0; i<e.Rows.Length; i++)
			{
				string sNumCB = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDepositoCajaConciliado {0}, {1}", (int)e.Rows[i].Cells["idAsiento"].Value, (int)e.Rows[i].Cells["idCuenta"].Value));
				if (sNumCB.Length > 0)
				{
					MessageBox.Show("El deposito esta Bloqueado, Consulte con el Departamento de Contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}				
				else if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las líneas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
				{
					int idDetalleCaja = (int) e.Rows[i].Cells["idDetalleCaja"].Value;
					int idAsiento = (int) e.Rows[i].Cells["idAsiento"].Value;
					if (idDetalleCaja > 0) 
					{
						if  (idAsiento == 0)
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From AsientoDetalle Where idAsiento = {0} Delete From Asiento Where idAsiento = {0}", (int) e.Rows[i].Cells["idAsiento"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetalleCaja Where idDetalleCaja = {0}", idDetalleCaja));	
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From ValidaDeposito Where idDetalleCaja = {0}", idDetalleCaja));	
							e.DisplayPromptMsg = false;				
						}
						else
						{
							if (!miAcceso.Eliminar)
							{
								MessageBox.Show("No tiene acceso a Eliminar Depositos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								e.Cancel = true;
								return;
							}
							else
							{
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From AsientoDetalle Where idAsiento = {0} Delete From Asiento Where idAsiento = {0}", (int) e.Rows[i].Cells["idAsiento"].Value));
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetalleCaja Where idDetalleCaja = {0}", idDetalleCaja));	
								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From ValidaDeposito Where idDetalleCaja = {0}", idDetalleCaja));				
								e.DisplayPromptMsg = false;
							}
						}
					}
					else
					{
						e.DisplayPromptMsg = false;
					}
				}
			}
		}

		private void uGridDepositos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDetalleCaja"].Value = 0;
			e.Row.Cells["idCaja"].Value = 0;
			e.Row.Cells["idAsiento"].Value = 0;
			e.Row.Cells["Valor"].Value = 0.00m;
			e.Row.Cells["FechaDeposito"].Value = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF); 
		}

		private void uGridDepositos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0 &&
				(this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "IDFORMAPAGO" 
				|| this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "IDCUENTA" 
				|| this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO" 
				|| this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "VALOR")) e.Handled = true;

			if (this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO") Funcion.CamposNumericos(sender, e);
		}

		private void uGridDepositos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void uGridDepositos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDepositos);
		}

		private void Suma (Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 1)
			{
				dEfectivo = 0; 
				dCheque = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
					if ((int)dr.Cells["idFormaPago"].Value == 1)
						dEfectivo = dEfectivo + Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value);												
			}
		}

		private void uGridDepositos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			#region comentado  
			//			if (e.Cell.Column.ToString() == "Valor")
			//			{
			//				Suma(e);
			//
			//				decimal dDiferenciaEfectivo = 0;
			//
			//				if (Math.Round(dEfectivo, 2) == Math.Round((decimal)this.spnCEfectivo.Value, 2))
			//				{
			//					this.lblEstadoDepositos.Text = "EFECTIVO CUADRADO";
			//				}
			//
			//				if (Math.Round(dEfectivo, 2) > Math.Round((decimal)this.spnCEfectivo.Value, 2))
			//				{
			//					decimal dDiferenciaEfectivo = dEfectivo - Math.Round((decimal)this.spnCEfectivo.Value, 2);
			//					this.lblEstadoDepositos.Text = "EFECTIVO SOBRANTE " + dDiferenciaEfectivo.ToString();
			//				}
			//
			//				if (Math.Round(dEfectivo, 2) < Math.Round((decimal)this.spnCEfectivo.Value, 2))
			//				{
			//					decimal dDiferenciaEfectivo = Math.Round((decimal)this.spnCEfectivo.Value, 2) - dEfectivo;
			//					this.lblEstadoDepositos.Text = "EFECTIVO FALTANTE " + dDiferenciaEfectivo.ToString();
			//				}
			//
			//				if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 2)
			//				{
			//					if (Math.Round(dCheque, 2) == Math.Round((decimal)this.spnCCheque.Value, 2))
			//					{
			//						this.lblEstadoDepositos.Text = this.lblEstadoDepositos.Text + " CHEQUE CUADRADO";
			//					}
			//					if (Math.Round(dCheque, 2) > Math.Round((decimal)this.spnCCheque.Value, 2))
			//					{
			//						this.lblEstadoDepositos.Text = this.lblEstadoDepositos.Text + " CHEQUE SOBRANTE";
			//					}
			//					if (Math.Round(dCheque, 2) < Math.Round((decimal)this.spnCCheque.Value, 2))
			//					{
			//						this.lblEstadoDepositos.Text = this.lblEstadoDepositos.Text + " CHEQUE FALTANTE";
			//					}
			//				}
			//			}

			/*if (e.Cell.Column.ToString() == "idFormaPago")
				{
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 1)
					{
						if ((decimal)this.spnCEfectivo.Value == 0)
						{
							if (e.Cell.OriginalValue != System.DBNull.Value) e.Cell.Row.Cells["idFormaPago"].Value = e.Cell.OriginalValue;
							else e.Cell.Row.Cells["idFormaPago"].Value = System.DBNull.Value;
							MessageBox.Show("El ingreso de efectivo a caja es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
						}
					}
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 2)
					{
						if ((decimal)this.spnCCheque.Value == 0)
						{
							if (e.Cell.OriginalValue != System.DBNull.Value) e.Cell.Row.Cells["idFormaPago"].Value = e.Cell.OriginalValue;
							else e.Cell.Row.Cells["idFormaPago"].Value = System.DBNull.Value;
							MessageBox.Show("El ingreso de cheques a caja es 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}
					}
				}*/	
			#endregion comentado

			#region Cambio Fecha 
			
			if (miAcceso.CambiarFecha)
			{
				if (bActualiza) return;
				if (e.Cell.Column.ToString() == "FechaDeposito"  )
				{	
					bActualiza = true;

					int idDetalleCaja = (int) e.Cell.Row.Cells["idDetalleCaja"].Value ;
					if (idDetalleCaja > 0) 
					{
						string sNumCB = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDepositoCajaConciliado {0}, {1}", (int)e.Cell.Row.Cells["idAsiento"].Value, (int)e.Cell.Row.Cells["idCuenta"].Value));

						if (sNumCB.Length < 0)
						{
							MessageBox.Show("El deposito esta Bloqueado, Consulte con el Departamento de Contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						}				
						else if (uGridDepositos.ActiveRow.Cells["FechaDeposito"].Value !=  System.DBNull.Value)
						{	
							bActualiza = true;
							DateTime FechaCaja = (DateTime) e.Cell.Row.Cells["FechaDeposito"].Value ;
							DateTime  FechasDia = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
							DateTime  FechaResta = FechasDia.AddDays(-1);
							if( Convert.ToDateTime(FechaCaja.Date) == Convert.ToDateTime(FechaResta.Date))
							{
								e.Cell.Row.Cells["FechaDeposito"].Value=  e.Cell.Row.Cells["FechaDeposito"].Value ;
								bActualiza = false;
							}
							else
							{
								//e.Cell.Row.Cells["FechaDeposito"].Value = (DateTime)e.Cell.Row.Cells["FechaDeposito"].OriginalValue;
								e.Cell.Row.Cells["FechaDeposito"].Value = FechaCajaBefore;
								bActualiza = false;
							}
						}
					}
				}
			}	
			#endregion Cambio Fecha 

			#region Valida si idDetalleCaja es mayor a 0
			if ( uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value !=  System.DBNull.Value)
			{
				
				if ((int) uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value > 0 && (int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value == 0)
				{	
					if (!miAcceso.Editar)
					{
						if (bActualiza) return;

						if (e.Cell.Column.ToString() == "idFormaPago")
						{	
							if (	uGridDepositos.ActiveRow.Cells["idFormaPago"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value > 0)
								{										
									e.Cell.Row.Cells["idFormaPago"].Value = (int)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}
						}	

						if (e.Cell.Column.ToString() == "idCuenta")
						{	
							if (uGridDepositos.ActiveRow.Cells["idCuenta"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value > 0)
								{										
									e.Cell.Row.Cells["idCuenta"].Value = (int)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}	
						}

						if (e.Cell.Column.ToString() == "Numero")
						{	
							if (uGridDepositos.ActiveRow.Cells["Numero"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value > 0)
								{										
									e.Cell.Row.Cells["Numero"].Value = (string)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}
						}	

						if (e.Cell.Column.ToString() == "FechaDeposito")
						{	
							if (uGridDepositos.ActiveRow.Cells["FechaDeposito"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value > 0)
								{										
									e.Cell.Row.Cells["FechaDeposito"].Value = FechaCajaBefore;
									//return;
								}	
								bActualiza = false;
							}	
						}

						if (e.Cell.Column.ToString() == "Valor")
						{	
							if (uGridDepositos.ActiveRow.Cells["Valor"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idDetalleCaja"].Value > 0)
								{										
									e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}	
						}
					}
				}
			}
			#endregion Valida si idDetalleCaja es mayor a 0

			#region Valida si idAsiento es mayor a 0
			if ( uGridDepositos.ActiveRow.Cells["idAsiento"].Value !=  System.DBNull.Value)
			{
			
				if ((int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0)
				{	
					if (!miAcceso.Editar)
					{
						if (bActualiza) return;

						if (e.Cell.Column.ToString() == "idFormaPago")
						{	
							if (	uGridDepositos.ActiveRow.Cells["idFormaPago"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0)
								{										
									e.Cell.Row.Cells["idFormaPago"].Value = (int)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}
						}	

						if (e.Cell.Column.ToString() == "idCuenta")
						{	
							if (uGridDepositos.ActiveRow.Cells["idCuenta"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0)
								{										
									e.Cell.Row.Cells["idCuenta"].Value = (int)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}	
						}

						if (e.Cell.Column.ToString() == "Numero")
						{	
							if (uGridDepositos.ActiveRow.Cells["Numero"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0)
								{										
									e.Cell.Row.Cells["Numero"].Value = (string)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}
						}	

						if (e.Cell.Column.ToString() == "FechaDeposito")
						{	
							if (uGridDepositos.ActiveRow.Cells["FechaDeposito"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0)
								{										
									e.Cell.Row.Cells["FechaDeposito"].Value = FechaCajaBefore;
									//return;
								}	
								bActualiza = false;
							}	
						}

						if (e.Cell.Column.ToString() == "Valor")
						{	
							if (uGridDepositos.ActiveRow.Cells["Valor"].Value !=  System.DBNull.Value)
							{
								bActualiza = true;

								if ((int) uGridDepositos.ActiveRow.Cells["idAsiento"].Value > 0)
								{										
									e.Cell.Row.Cells["Valor"].Value = (decimal)e.Cell.OriginalValue;
									//return;
								}	
								bActualiza = false;
							}	
						}
					}
				}
			}
			#endregion Valida si idAsiento es mayor a 0
	
			#region Valida banco numero
			if (e.Cell.Column.ToString() == "idCuenta"  || e.Cell.Column.ToString() == "Numero" )
			{	
				bActualiza = true;
				if (uGridDepositos.ActiveRow.Cells["idFormaPago"].Value !=  System.DBNull.Value)
				{	

					if ((int) uGridDepositos.ActiveRow.Cells["idFormaPago"].Value == 4 || (int) uGridDepositos.ActiveRow.Cells["idFormaPago"].Value == 16)
					{					
					 
						this.uGridDepositos.ActiveRow.Cells["idFormaPago"].Value = 0;
						this.uGridDepositos.ActiveRow.Cells["idFormaPago"].Value ="";
									
					}	
				}
			
				bActualiza = false;
			}
			#endregion Valida banco numero

			//			if (e.Cell.Column.ToString() == "Valor"  )
			//			{	
			//				bActualiza = true;
			//
			//				if (uGridDepositos.ActiveRow.Cells["idFormaPago"].Value !=  System.DBNull.Value)
			//				{	
			//					if ((int) uGridDepositos.ActiveRow.Cells["idFormaPago"].Value == 16 )
			//					{					
			//					 
			//						this.uGridDepositos.ActiveRow.Cells["Valor"].Value =  (Convert.ToDecimal(	this.uGridDepositos.ActiveRow.Cells["Valor"].Value) * (-1));
			//									
			//					}	
			//				}
			//			
			//				bActualiza = false;
			//			}
		}
		
		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Impreso")
				{		
					this.ultraGrid1.UpdateData();
					
					string sSQL = string.Format("Exec ActualizaImpresoFacturasAnticipos {0}, {1}, '{2}'", 
						(bool)e.Cell.Row.Cells["Impreso"].Value, (int)e.Cell.Row.Cells["idCompra"].Value, e.Cell.Row.Cells["Origen"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}

				if (e.Cell.Column.ToString() == "R")
				{		
					this.ultraGrid1.UpdateData();

					string sSQL = string.Format("Exec CajaValidaEstadoCobrosExternos {0}", (int)e.Cell.Row.Cells["idCompra"].Value);
					string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQL);

					if (sMensaje.Length > 0)
					{
						e.Cell.Row.Cells["R"].Value = (bool)e.Cell.OriginalValue;
						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
						return;
					}
					
					if (e.Cell.Row.Cells["Estado"].Value.ToString() == "ANULADO")
					{
						e.Cell.Row.Cells["R"].Value = (bool)e.Cell.OriginalValue;
						MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
						return;
					}

					sSQL = string.Format("Exec CajaActualizaEstadoCobrosExternos {0}, {1}", 
						(bool)e.Cell.Row.Cells["R"].Value, (int)e.Cell.Row.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}					
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "R")
			{
				if (e.Cell.Row.Cells["idCompra"].Value == System.DBNull.Value)
					e.Cancel = true;

				if (e.Cell.Row.Cells["Origen"].Value.ToString() != "PAG")
					e.Cancel = true;
			}

			if (e.Cell.Column.ToString() == "Impreso")
			{
				if (e.Cell.Row.Cells["idCompra"].Value == System.DBNull.Value)
					e.Cancel = true;
				
				if (e.Cell.Row.Cells["Origen"].Value.ToString() == "RET" || e.Cell.Row.Cells["Origen"].Value.ToString() == "DA")
					e.Cancel = true;				
			}

			if (e.Cell.Column.ToString() == "Fecha")
			{
				if (e.Cell.Row.Cells["Origen"].Value.ToString() == "RET" || e.Cell.Row.Cells["Origen"].Value.ToString() == "DA")
					e.Cancel = true;
				
				if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 9)
					e.Cancel = true;				
			}
			if (e.Cell.Column.ToString() == "idPersonal")
			{
				if ((int)e.Cell.Row.Cells["idOrigen"].Value != 1)
				{
					MessageBox.Show("El Vendedor se puede editar solamente en Facturas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					e.Cancel = true;
					return;
				}

				if (e.Cell.Row.Cells["Estado"].Value.ToString() == "ANULADO")
				{
					MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}

				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta")) 
				{	
					MessageBox.Show(string.Format("No Puede Editar esta Factura, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", Convert.ToDateTime(this.dtFecha.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);												
					e.Cancel = true;
					return;
				}
			}			
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			if (e.Cell.Column.ToString() == "btnCaja")	
			{
				#region Validacion
				if (e.Cell.Row.Cells["Estado"].Value.ToString() == "ANULADO")
				{
					MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idOrigen"].Value == 5)
				{
					MessageBox.Show("El documento es de devolución de anticipos no puede modificar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idOrigen"].Value == 1)
				{
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 9)
					{
						if (Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value) == Convert.ToDecimal(e.Cell.Row.Cells["CrediPoint"].Value))
						{
							MessageBox.Show("El documento es de CrediPOINT y no tiene entrada, no puede modificar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}

					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 6)
					{
						MessageBox.Show("El documento es de Credito, no puede modificar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;						
					}
				}
				#endregion Validacion

				#region Anticipos
				bool bUsoAnticipo = false;
				decimal dSaldoAnticipo = 0;
							
				if ((int)e.Cell.Row.Cells["idOrigen"].Value == 1)
				{
					dSaldoAnticipo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.SaldoAnticipos({0})", (int)e.Cell.Row.Cells["idCliente"].Value));
				
					if (dSaldoAnticipo > 0)
					{
						if	(DialogResult.Yes == MessageBox.Show(string.Format("El cliente tiene un saldo a favor de {0} dólares, ¿Desea cruzarlo?", dSaldoAnticipo), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 				
							bUsoAnticipo = true;
						else
						{
							bUsoAnticipo = false;
							dSaldoAnticipo = 0.00m;
						}
					}
				}
				#endregion Anticipos

				#region Formas de pago
				using (VentaFormaDePago Cobro = new VentaFormaDePago ((int)e.Cell.Row.Cells["idCompra"].Value, Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value), 
								 (int)e.Cell.Row.Cells["idOrigen"].Value, Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value), bUsoAnticipo, dSaldoAnticipo, true, (int)e.Cell.Row.Cells["idCliente"].Value, 
								 (int)e.Cell.Row.Cells["Bodega"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value))
				{
					if (DialogResult.OK == Cobro.ShowDialog())
					{
						#region Guarda Cobro de la factura	
						if (Cobro.ultraGrid1.Rows.Count > 0)
						{
							#region Guarda
							decimal dAnticipo = 0.00m;								
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Cobro.ultraGrid1.Rows.All)
							{
								#region Variables
								int idDetFormaPago = 0;
								int idBanco = 0;
								int idTarjeta = 0;
								int idCuenta = 0;
								int idPlan = 0;
								int idPlazo = 0;
								decimal dComTarjeta = 0.00m;
								int idDocumento = 0;
								DateTime dtFecCaducidad = DateTime.Today;
							
								if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
								if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
								if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
								if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
								if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
								if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
								if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
								if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = (decimal)dr.Cells["ValorUsoTarjeta"].Value;
								if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;
								#endregion Variables

								string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}", 
									idDetFormaPago, 
									(int)e.Cell.Row.Cells["idCompra"].Value, (int)dr.Cells["Origen"].Value,  
									(int)dr.Cells["idFormaPago"].Value, (decimal)dr.Cells["Valor"].Value,
									idBanco, idTarjeta,
									dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
									dtFecCaducidad.ToString("yyyyMMdd"),								
									dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
									idCuenta, 
									dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(),
									idPlan, idPlazo, dComTarjeta, idDocumento);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro);

								/* VALOR DE ANTICIPOS */
								if ((int)dr.Cells["idFormaPago"].Value == 8) 
									dAnticipo = dAnticipo + Convert.ToDecimal(dr.Cells["Valor"].Value);

								/* ACTUALIZA SALDO DE NOTA DE CREDITO */
								if (idDocumento > 0)
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 5, 4", idDocumento));
							}
							#endregion Guarda

							#region Asientos
							if ((int)e.Cell.Row.Cells["Tipo"].Value == 0) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CreaAsientoDeVenta {0}", (int)e.Cell.Row.Cells["idCompra"].Value));
							else if ((int)e.Cell.Row.Cells["Tipo"].Value == 1 || (int)e.Cell.Row.Cells["Tipo"].Value == 5 || (int)e.Cell.Row.Cells["Tipo"].Value == 12)
							{
								string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", 
									(int)e.Cell.Row.Cells["idAsiento"].Value, (int)e.Cell.Row.Cells["idCompra"].Value, (int) (int)e.Cell.Row.Cells["Tipo"].Value, (int)e.Cell.Row.Cells["idOrigen"].Value);
								Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);
							}
							else if ((int)e.Cell.Row.Cells["Tipo"].Value == 11)
							{
								string sSQLAsiento = string.Format("Exec CreaAsientoAnticipo {0}, {1}, {2}, {3}", 
									(int)e.Cell.Row.Cells["idAsiento"].Value, (int)e.Cell.Row.Cells["idCompra"].Value, (int) (int)e.Cell.Row.Cells["Tipo"].Value, (int)e.Cell.Row.Cells["idOrigen"].Value);
								Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

								string sSQLADep = string.Format("Exec CreaAsientoDepreciacionCP {0}", (int)e.Cell.Row.Cells["idCompra"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLADep);
							}	
							else if ((int)e.Cell.Row.Cells["Tipo"].Value == 3)
							{
								string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, {1}", 
									(int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idAsiento"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
							}
							else if ((int)e.Cell.Row.Cells["Tipo"].Value == 2 || (int)e.Cell.Row.Cells["Tipo"].Value == 15 || (int)e.Cell.Row.Cells["Tipo"].Value == 16 || (int)e.Cell.Row.Cells["Tipo"].Value == 17 || (int)e.Cell.Row.Cells["Tipo"].Value == 18 || (int)e.Cell.Row.Cells["Tipo"].Value == 19)
							{								
								string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, {1}, {2}", 
									(int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idAsiento"].Value, (int)e.Cell.Row.Cells["idOrigen"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);									
							}
							#endregion Asientos

							#region Saldo Factura
							if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 9)
							{
								string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 1, 9", (int)e.Cell.Row.Cells["idCompra"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo);
							}
							#endregion Saldo Factura

							#region Anticipos
							if (dAnticipo > 0)
							{									
								string sSQLAnt = string.Format("Exec ActualizarAnticipos {0}, {1}, {2}", (int)e.Cell.Row.Cells["idCliente"].Value, dAnticipo, (int)e.Cell.Row.Cells["idCompra"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAnt);																
							}
							#endregion Anticipos
						}								
						#endregion Guarda Cobro de la factura

						this.btnVer_Click(sender, e);
					}					
				}
				#endregion Formas de pago
			}

			if (e.Cell.Column.ToString() == "btnTarjetas")	
			{
				#region Validacion
				if (e.Cell.Row.Cells["Estado"].Value.ToString() == "ANULADO")
				{
					MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idOrigen"].Value != 1)
				{
					MessageBox.Show("Solo para facturas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idOrigen"].Value == 1)
				{
					decimal dTotalTarjeta = Convert.ToDecimal(e.Cell.Row.Cells["Tarjeta"].Value) + Convert.ToDecimal(e.Cell.Row.Cells["Debito"].Value);
					if (dTotalTarjeta == 0.00m)
					{
						MessageBox.Show("El documento no tiene valores por tarjeta de credito o debito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				#endregion Validacion

				#region Lista
				using (frmCajaListaPagosConTarjetasPorFactura Tarjetas = new frmCajaListaPagosConTarjetasPorFactura ((int)e.Cell.Row.Cells["idCompra"].Value))
				{
					if (DialogResult.OK == Tarjetas.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AutPrincipal = 1 Where idCompra = {0}", (int)e.Cell.Row.Cells["idCompra"].Value));
					}					
				}
				#endregion Lista
			}

			if (e.Cell.Column.ToString() == "btnVerTarjeta")	
			{
				if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 3)
				{
					#region VerTarjeta

					using (frmVentaDetalleFormasDePago Tarjetas = new frmVentaDetalleFormasDePago ((int)e.Cell.Row.Cells["idCompra"].Value))
					{
						if (DialogResult.OK == Tarjetas.ShowDialog())
						{

						}					
					}
					#endregion VerTarjeta
				}
				else
				{
					return;
				}				
			}

			#region Retencion
			if (e.Cell.Column.ToString() == "btnRetencionVenta")	
			{
				#region validacion
				int iFP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFormaPago From Compra Where idCompra = {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value));
				int idCliente = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idCliente from compra where idCompra = {0}", (int)e.Cell.Row.Cells["idCompra"].Value));
				int Ruc = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select top 1 cl.idTipoRuc from compra c inner join Cliente cl on c.idCliente = cl.idcliente where cl.idcliente = {0}", idCliente));
				int sValidar = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From DetalleFormaPago Where Origen = 1 And idCompra = {0} And idFormaPago not in (1, 8, 10)", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value));
				
				if (iFP != 1) 
				{
					MessageBox.Show("Solo Para Forma de Pago Efectivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if(sValidar > 0)
				{
					MessageBox.Show("La Factura Tiene Pagos con Tarjetas de Credito o Debito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if(Ruc > 4)
				{
					MessageBox.Show("Cliente No Genera Retencion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				
				if (e.Cell.Row.Cells["Estado"].Value.ToString() == "ANULADO")
				{
					MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)e.Cell.Row.Cells["idOrigen"].Value != 1)
				{
					MessageBox.Show("Solo para facturas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#endregion validacion

				if ((int)e.Cell.Row.Cells["idOrigen"].Value == 1)
				{
					if((int)e.Cell.Row.Cells["idFormaPago"].Value == 1)// || (int)e.Cell.Row.Cells["idFormaPago"].Value == 6)
					{
						//int iCont = 0;
						//int idCompra = 0;
						DateTime dtFecha = DateTime.Today;
						string sNumero = this.ultraGrid1.ActiveRow.Cells["Factura"].Value.ToString();
						dtFecha = (DateTime)this.ultraGrid1.ActiveRow.Cells["Fecha"].Value;

						int idRetencion = 
							Funcion.iEscalarSQL(cdsSeteoF, 
							string.Format("if (select count(*) from Retenciones Where idCompra = {0}) > 0 select idretencion from Retenciones where idCompra = {0} else select idCompra = 0", 
							(int)e.Cell.Row.Cells["idCompra"].Value));

						using(frmRetencionesVenta Retencion = new frmRetencionesVenta(idRetencion, (int)e.Cell.Row.Cells["idCompra"].Value, sNumero, (decimal)e.Cell.Row.Cells["Total"].Value, dtFecha, 1))
						{
							if (DialogResult.OK == Retencion.ShowDialog())
							{
								DateTime dtFechaRet = (DateTime)Retencion.dtFecha.Value;				
								DateTime dtFechaCadR = (DateTime)Retencion.dtCaducidad.Value;

								#region Guarda Maestro
								if((int)Retencion.txtIdRetencion.Value == 0)
								{
									string sSQLGR = string.Format("Exec GuardaRetenciones {0}, {1}, 0, 0, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}", 
										(int)Retencion.txtIdRetencion.Value, (int)Retencion.txtIdCompra.Value, Retencion.txtSerieFact.Text.ToString(), Retencion.txtNumero.Text.ToString(), Retencion.txtAutFactura.Text.ToString(), 
										dtFechaCadR.ToString("yyyyMMdd"), dtFechaRet.ToString("yyyyMMdd"), "", Convert.ToDecimal(Retencion.txtRetenidoIVA.Value), Convert.ToDecimal(Retencion.txtRetenidoRenta.Value), Convert.ToDecimal(Retencion.txtRetenido.Value));
									Retencion.txtIdRetencion.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGR);
								}
								#endregion Guarda Maestro

								#region Valida si la retencion existe
								string sSQLValida = string.Format("Exec ValidaDetalleDeRetencion {0}", (int)Retencion.txtIdRetencion.Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLValida);					
								#endregion Valida si la retencion existe

								#region Guarda detalle
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Retencion.ultraGrid1.Rows.All)
								{
									string sSQLGDR = string.Format("Exec GuardaDetalleRetenciones {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
										(int)dr.Cells["idDetalleRetencion"].Value, (int)Retencion.txtIdRetencion.Value, (bool)dr.Cells["IVARenta"].Value, (int)dr.Cells["idRetencionIVARenta"].Value, Convert.ToDecimal(dr.Cells["Base"].Value), Convert.ToDecimal(dr.Cells["Valor"].Value),
										Convert.ToDecimal(dr.Cells["Porcentaje"].Value));
									Funcion.EjecutaSQL(cdsSeteoF, sSQLGDR, true);
								}
								#endregion Guarda detalle
								
								#region Inserta Detalle Forma Pago
								DateTime dtFecCaducidad = DateTime.Today;

								string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
									0, (int)Retencion.txtIdRetencion.Value, 6, 1, Convert.ToDecimal(Retencion.txtRetenido.Value), 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);							
								#endregion Inserta Detalle Forma Pago

								#region Crea Asiento
								string sSQLAsientoR = string.Format("Exec CreaAsientoRetencionVenta 0, {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}, '{9}'", 
									(int)Retencion.txtIdRetencion.Value, (int)e.Cell.Row.Cells["idCliente"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value, e.Cell.Row.Cells["Nombre"].Value.ToString(), 
									Retencion.txtNumero.Text.ToString(), e.Cell.Row.Cells["Factura"].Value.ToString(), dtFechaRet.ToString("yyyyMMdd"), 19, 
									Convert.ToDecimal(Retencion.txtRetenido.Value), e.Cell.Row.Cells["Nombre"].Value.ToString());
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAsientoR, true);								
								#endregion Crea Asiento

								this.btnVer_Click(sender, e);
							}
						}
					}
					else
					{
						MessageBox.Show("El documento no tiene valores por Efectivo o Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
			}
			#endregion Retencion

			if (e.Cell.Column.ToString() == "btnDeuna")	
			{
				if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 32)
				{
					string sURl ="";
					 sURl = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select UrlGoogle from compra where idcompra = {0}", (int)e.Cell.Row.Cells["idCompra"].Value), true);		

					if (sURl.Length > 10 )
					{
						System.Diagnostics.Process.Start("chrome.exe", sURl);
					}
					else
					{
						MessageBox.Show("No existe el documento, comuniquese con el Administrador.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					}
				}
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 45, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
				//				this.btnVer_Click(sender, e);
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			idCaja = 0;
			iEstado = 0;
			iEstadoCaja = 0;

			this.lblEstado.Text = "";
			this.lblEstadoDepositos.Text = "";
			this.txtObservaciones.Text = "";
			this.txtObservaciones.Enabled = false;

			this.btnDepositar.Enabled = false;
			this.btnActualizar.Enabled = false;
			this.btnGuardar.Enabled = false;
			
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Factura"].CellActivation = Activation.ActivateOnly;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.ActivateOnly;
			
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["R"].CellActivation = Activation.Disabled;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impreso"].CellActivation = Activation.Disabled;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnCaja"].CellActivation = Activation.Disabled;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnTarjetas"].CellActivation = Activation.Disabled;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnRetencionVenta"].CellActivation = Activation.Disabled;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridDepositos);
			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", false, "", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Caja")) return;

			this.Cursor = Cursors.WaitCursor;
						
			string sSQL = string.Format("Exec ReporteCierreDeCajaSistema '{0}', {1}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.ultraGrid1);
			
			CalculaTotal();

			#region Encabezado
			string sSQLMaestro = string.Format("Exec ConsultaCaja '{0}', {1}", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value);
			SqlDataReader drM = Funcion.rEscalarSQL(cdsSeteoF, sSQLMaestro, true);
			drM.Read();
			if (drM.HasRows)
			{
				idCaja = (int)drM.GetValue(0);
				iEstado = (int)drM.GetValue(4);								
				this.txtObservaciones.Text = drM.GetString(6);
				iEstadoCaja = drM.GetInt32(7);			
			}			
			drM.Close();			
			#endregion Encabezado

			#region Actualiza Caja
			DateTime Fecha = (DateTime)this.dtFecha.Value;
			
			int iProyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value), true);		
	
			decimal dValDeposito = Convert.ToDecimal(this.spnCEfectivo.Value) + Convert.ToDecimal(this.spnCCheque.Value) + Convert.ToDecimal(this.spnC_Entrega.Value)+ Convert.ToDecimal(this.spnCTC.Value); 

			string sSQLActCaja = string.Format("Exec GuardaCaja {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}", 
				idCaja, (int)this.cmbBodega.Value, iProyecto, Fecha.ToString("yyyyMMdd"), "", "", 
				(decimal)this.spnCEfectivo.Value, (decimal)this.spnCCheque.Value, dValDeposito, 1, (bool)this.chkRevisado.Checked, this.txtObservaciones.Text.ToString(), (decimal)this.spnC_Entrega.Value, (decimal)this.spnCTC.Value );
			idCaja = Funcion.iEscalarSQL(cdsSeteoF, sSQLActCaja);
						
			this.lblEstadoDepositos.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeCaja {0}", idCaja));			
			#endregion Actualiza Caja

			#region Detalle
			if (idCaja > 0)
			{	
				if (!miAcceso.BProcesar) 
				{
					if ((decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value > 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2, 3) Order By idFormaPago");
					if ((decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value == 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In ( 1, 3) Order By idFormaPago");
					if ((decimal)this.spnCEfectivo.Value == 0 && (decimal)this.spnCCheque.Value > 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In ( 2, 3) Order By idFormaPago");

					if (((decimal)this.spnCTC.Value) > 0 && (decimal)this.spnCEfectivo.Value == 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago = 3 Order By idFormaPago");

					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value > 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2, 19, 3) Order By idFormaPago");
					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value == 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 19, 3) Order By idFormaPago");
					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value == 0 && (decimal)this.spnCCheque.Value > 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 19, 3) Order By idFormaPago");
					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value == 0 && (decimal)this.spnCCheque.Value == 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 19, 3) Order By idFormaPago");
				}
				else
				{
					if ((decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value > 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2, 4,16, 3) Order By idFormaPago");
					if ((decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value == 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In ( 1, 4, 16, 3) Order By idFormaPago");
					if ((decimal)this.spnCEfectivo.Value == 0 && (decimal)this.spnCCheque.Value > 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In ( 2, 4, 16, 3) Order By idFormaPago");

					if (((decimal)this.spnCTC.Value) > 0 && (decimal)this.spnCEfectivo.Value == 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago = 3 Order By idFormaPago");

					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value > 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2, 19, 4, 16,3)) Order By idFormaPago");
					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value > 0 && (decimal)this.spnCCheque.Value == 0) this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 19, 4, 16, 3) Order By idFormaPago");
					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value == 0 && (decimal)this.spnCCheque.Value > 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 19, , 4, 16,3) Order By idFormaPago");
					if ((decimal)this.spnC_Entrega.Value > 0 && (decimal)this.spnCEfectivo.Value == 0 && (decimal)this.spnCCheque.Value == 0)	this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 19, , 4, 16, 3) Order By idFormaPago");
				

				}
				this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCaja {0}", idCaja));

				if (this.uGridDepositos.Rows.Count > 0) this.lblEstado.Text = "CONTABILIZADO";					
			}
			#endregion Detalle		

			#region Controles
			if (iEstadoCaja == 1) 
			{
				this.lblEstado.Text = "CAJA ABIERTA";
				this.btnGuardar.Text = "Cerrar Caja";
				this.btnDepositar.Enabled = true;
				this.txtObservaciones.Enabled = true;

				if (miAcceso.BActNum) 
				{
					this.btnActualizar.Enabled = true;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Factura"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.AllowEdit;					
				}

				if (miAcceso.BRevisadoCE) this.ultraGrid1.DisplayLayout.Bands[0].Columns["R"].CellActivation = Activation.AllowEdit;
				if (miAcceso.BModificarImpreso) this.ultraGrid1.DisplayLayout.Bands[0].Columns["Impreso"].CellActivation = Activation.AllowEdit;
				if (miAcceso.BModifcarFormaDePago) 
				{
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnCaja"].CellActivation = Activation.AllowEdit;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnTarjetas"].CellActivation = Activation.AllowEdit; 
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnRetencionVenta"].CellActivation = Activation.AllowEdit;
				}
				if (miAcceso.Nuevo)  this.ultraGrid1.DisplayLayout.Bands[0].Columns["idPersonal"].CellActivation = Activation.AllowEdit;
				if (miAcceso.Anular) this.btnAceptar.Enabled = true;	

				if (miAcceso.BCerrarCaja) this.btnGuardar.Enabled = true;
				if (!miAcceso.BIrA)		this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnVerTarjeta"].Hidden = true;
				if (miAcceso.BIrA)		this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnVerTarjeta"].Hidden = false;
				
				if (this.ultraGrid1.Rows.Count > 0) 
				{
					FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
					FuncionesProcedimientos.EstadoGrids(true, this.uGridDepositos);
				}
			}
			if (iEstadoCaja == 2) 
			{
				this.lblEstado.Text = "CAJA CERRADA";
				this.btnGuardar.Text = "Abrir Caja";
				this.btnAceptar.Enabled = false;	
				if (miAcceso.BAbrirCaja) this.btnGuardar.Enabled = true;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idPersonal"].CellActivation = Activation.NoEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnVerTarjeta"].Hidden = true;
			}

			if (this.ultraGrid1.Rows.Count > 0) this.btnImprimir.Enabled = true;
			#endregion Controles

			string sSQLValidaHost = string.Format("Exec [ValidaImpresionCaja] {0}",idCaja );
			int iVal = Funcion.iEscalarSQL(cdsSeteoF, sSQLValidaHost, true);

			if (miAcceso.ReImprimir == true)
			{
				if(iVal == 1)
				{
					this.btnImprimir.Enabled = true;
				}
			}
			else
			{
				if(iVal == 1)
				{
					this.btnImprimir.Enabled = false;
				}
			}

			this.Cursor = Cursors.Default;			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			string sSQLValidaHost = string.Format("Exec [ValidaImpresionCaja] {0}",idCaja);
			int iVal = Funcion.iEscalarSQL(cdsSeteoF, sSQLValidaHost, true);
			      
			if (iVal > 0 && !miAcceso.ReImprimir)
			{
				MessageBox.Show("La caja ya fue impresa, No tiene acceso a Re imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (iVal > 1)
				if (DialogResult.No == MessageBox.Show("¿Esta seguro de re imprimir el Cuadre de Caja?, Si presiona SI ya no podra volver a Reimprimir", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					return;

			#region Imprimir
			try
			{
				this.Cursor = Cursors.WaitCursor;
				DateTime dtFecha = (DateTime) this.dtFecha.Value;

				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue discreteValFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@Fecha";
				discreteValFecha.Value = dtFecha;
				paramFieldFecha.CurrentValues.Add (discreteValFecha);
				paramFields.Add (paramFieldFecha);

				ParameterField paramFieldBodega = new ParameterField ();
				ParameterDiscreteValue discreteValBodega = new ParameterDiscreteValue ();
				paramFieldBodega.ParameterFieldName = "@Bodega";
				discreteValBodega.Value = (int)this.cmbBodega.Value;
				paramFieldBodega.CurrentValues.Add (discreteValBodega);
				paramFields.Add (paramFieldBodega);

				ParameterField paramFieldLocal = new ParameterField ();
				ParameterDiscreteValue discreteValLocal = new ParameterDiscreteValue ();
				paramFieldLocal.ParameterFieldName = "@Local";
				discreteValLocal.Value = this.cmbBodega.Text.ToString();
				paramFieldLocal.CurrentValues.Add (discreteValLocal);
				paramFields.Add (paramFieldLocal);

				ParameterField paramFieldUsuario = new ParameterField ();
				ParameterDiscreteValue discreteValUsuario = new ParameterDiscreteValue ();
				paramFieldUsuario.ParameterFieldName = "@Usuario";
				discreteValUsuario.Value = MenuLatinium.stUsuario;
				paramFieldUsuario.CurrentValues.Add (discreteValUsuario);
				paramFields.Add (paramFieldUsuario);

				ParameterField paramFieldidCaja = new ParameterField ();
				ParameterDiscreteValue discreteValidCaja = new ParameterDiscreteValue ();
				paramFieldidCaja.ParameterFieldName = "idCaja";
				discreteValidCaja.Value = idCaja.ToString();
				paramFieldidCaja.CurrentValues.Add (discreteValidCaja);
				paramFields.Add (paramFieldidCaja);

				ParameterField paramFieldEs = new ParameterField ();
				ParameterDiscreteValue discreteValEs = new ParameterDiscreteValue ();
				paramFieldEs.ParameterFieldName = "@Estacion";
				discreteValEs.Value = strHostName;
				paramFieldEs.CurrentValues.Add (discreteValEs);
				paramFields.Add (paramFieldEs);

				Reporte miRep = new Reporte("CierreDeCaja.rpt", "");

				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Cierre de Caja");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
								
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Caja Set NumImp = ISNULL(NumImp, 0) + 1 Where idCaja = {0}", idCaja));

				Cursor = Cursors.Default;
			}
			catch(Exception exc)
			{
				MessageBox.Show(string.Format("Error de impresion : '{0}'", exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			#endregion Imprimir
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnDepositar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", true, "Bancos", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Caja")) return;
			
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No Existen documentos para realizar el deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if ((decimal)this.spnCEfectivo.Value == 0.00m && (decimal)this.spnCCheque.Value == 0.00m && (decimal)this.spnC_Entrega.Value == 0.00m && (decimal)this.spnCTC.Value == 0.00m)
			{
				MessageBox.Show("No existen Valores para depositar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}			
					
			#region Validacion Detalle
			if (this.uGridDepositos.Rows.Count == 0)
			{
				MessageBox.Show("No ha ingresado los depositos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{	
				bool bNdcNDb = false;
				if ((int) dr.Cells["idFormaPago"].Value == 4 || (int) dr.Cells["idFormaPago"].Value == 16 || (int) dr.Cells["idFormaPago"].Value == 3  )
				{
					bNdcNDb = true;  
				}
				if (!bNdcNDb)
				{
					if (dr.Cells["idCuenta"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione la cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDepositos.ActiveRow = dr;
						this.uGridDepositos.ActiveRow.Selected = true;
						return;
					}		
				}
				if (dr.Cells["FechaDeposito"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha de depósito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;
					return;
				}
       
				if (Convert.ToInt32(dr.Cells["idDetalleCaja"].Value) == 0)
				{		
					if (Convert.ToDateTime(dr.Cells["FechaDeposito"].Value) < Convert.ToDateTime(this.dtFecha.Value))
					{
						MessageBox.Show("La Fecha del Deposito No puede ser Anterior a la Fecha de la Caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDepositos.ActiveRow = dr;
						this.uGridDepositos.ActiveRow.Selected = true;
						return;
					}
				}	
//				DateTime dtFechaActual = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);
//
//				if (Convert.ToDateTime(dr.Cells["FechaDeposito"].Value) > dtFechaActual)
//				{
//					MessageBox.Show("La Fecha del Deposito No puede ser posterior a Hoy", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.uGridDepositos.ActiveRow = dr;
//					this.uGridDepositos.ActiveRow.Selected = true;
//					return;
//				}

				if (Funcion.ValidaPeriodo((DateTime)dr.Cells["FechaDeposito"].Value, cdsSeteoF, "Bancos"))
				{						
					DateTime dtFechaVP = (DateTime)dr.Cells["FechaDeposito"].Value;
					MessageBox.Show(string.Format("No puede Registrar depositos en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;
					return;
				}

				if (!bNdcNDb)
				{
					string sNumCB = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDepositoCajaConciliado {0}, {1}", (int)dr.Cells["idAsiento"].Value, (int)dr.Cells["idCuenta"].Value));
					if (sNumCB.Length > 0)
					{
						MessageBox.Show("El deposito esta Conciliado, Consulte con el Departamento de Contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDepositos.ActiveRow = dr;
						this.uGridDepositos.ActiveRow.Selected = true;
						return;
					}
				
					if (dr.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Escriba el número de depósito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDepositos.ActiveRow = dr;
						this.uGridDepositos.ActiveRow.Selected = true;
						return;
					}
				}
				if ((int) dr.Cells["idFormaPago"].Value == 3)
				{
					if (dr.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Escriba el número de LOTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDepositos.ActiveRow = dr;
						this.uGridDepositos.ActiveRow.Selected = true;
						return;
					}

//					if (Convert.ToDateTime(this.dtFecha.Value) != Convert.ToDateTime(dr.Cells["FechaDeposito"].Value))
//					{
//						MessageBox.Show("La fecha de cierre de tarjeta de credito no puede ser diferente al de la caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.uGridDepositos.ActiveRow = dr;
//						this.uGridDepositos.ActiveRow.Selected = true;
//						return;
//					}
				}
				if (dr.Cells["Valor"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el valor del deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;
					return;
				}
				if (Convert.ToDecimal(dr.Cells["Valor"].Value) == 0.00m)
				{
					MessageBox.Show("Ingrese el valor del deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;
					return;
				}
				if (!bNdcNDb && (int) dr.Cells["idFormaPago"].Value != 3)
				{
					if ((int)dr.Cells["idAsiento"].Value == 0)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeroDocumento {0}, '{1}', 2, {2}", (int)dr.Cells["idCuenta"].Value, dr.Cells["Numero"].Value.ToString(), (int)dr.Cells["idAsiento"].Value), true))
						{
							MessageBox.Show("El numero de deposito ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridDepositos.ActiveRow = dr;
							this.uGridDepositos.ActiveRow.Selected = true;
							return;
						}
					}
				}
				DateTime dFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF); 
				if ((int)dr.Cells["idDetalleCaja"].Value == 0)
				{
					if (MenuLatinium.stDirFacturacion != "DENTALIS")
					{
						//DateTime.Today;
						if (Convert.ToDateTime(dr.Cells["FechaDeposito"].Value).Date != dFecha.Date)
						{						
							MessageBox.Show("La Fecha del DEPOSITO no puede ser diferente al dia actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDepositos.ActiveRow = dr;
							this.uGridDepositos.ActiveRow.Selected = true;
							return;
						}
					}
					else
					{
						if (!miAcceso.CambiarFechaPosterior)
						{
							MessageBox.Show("La Fecha del DEPOSITO no puede ser diferente al dia actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDepositos.ActiveRow = dr;
							this.uGridDepositos.ActiveRow.Selected = true;
							return;
						}
					}
				}
			}
			#endregion Validacion Detalle
			
			#region Proyecto
			string sSQLProyecto = string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value);
			int iProyecto = Funcion.iEscalarSQL(cdsSeteoF, sSQLProyecto, true);
			#endregion Proyecto

			decimal dValDeposito = 0.00m;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
				dValDeposito = dValDeposito + Convert.ToDecimal(dr.Cells["Valor"].Value); 
			
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
					string sSQLMaestro = string.Format("Exec GuardaCaja {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}", 
						idCaja, (int)this.cmbBodega.Value, iProyecto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), "", "", 
						(decimal)this.spnCEfectivo.Value, (decimal)this.spnCCheque.Value, dValDeposito, 1, (bool)this.chkRevisado.Checked, this.txtObservaciones.Text.ToString(), (decimal)this.spnC_Entrega.Value, (decimal)this.spnCTC.Value);			
					oCmdActualiza.CommandText = sSQLMaestro;
					idCaja = (int)oCmdActualiza.ExecuteScalar();
			
					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
					{
						bool bNdcNDb = false;
						if ((int) dr.Cells["idFormaPago"].Value == 4 || (int) dr.Cells["idFormaPago"].Value == 16 || (int) dr.Cells["idFormaPago"].Value == 3  ) bNdcNDb = true;  
						string sSQLDetalleEf = "";
						if(!bNdcNDb)
						{
							sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
								(int)dr.Cells["idDetalleCaja"].Value, idCaja, (int)dr.Cells["idCuenta"].Value, (int)dr.Cells["idFormaPago"].Value, dr.Cells["Numero"].Value.ToString(),  
								Convert.ToDateTime(dr.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)dr.Cells["Valor"].Value);
							oCmdActualiza.CommandText = sSQLDetalleEf;
							oCmdActualiza.ExecuteNonQuery();
							
						}
						else
						{
							if ((int) dr.Cells["idFormaPago"].Value == 4  )  
							{
								sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
									(int)dr.Cells["idDetalleCaja"].Value, idCaja, 0, (int)dr.Cells["idFormaPago"].Value, "",  
									Convert.ToDateTime(dr.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)dr.Cells["Valor"].Value);
								oCmdActualiza.CommandText = sSQLDetalleEf;
								oCmdActualiza.ExecuteNonQuery();
							}
							if ( (int) dr.Cells["idFormaPago"].Value == 3 )  
							{
								sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
									(int)dr.Cells["idDetalleCaja"].Value, idCaja, 0, (int)dr.Cells["idFormaPago"].Value, dr.Cells["Numero"].Value.ToString(),  
									Convert.ToDateTime(dr.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)dr.Cells["Valor"].Value);
								oCmdActualiza.CommandText = sSQLDetalleEf;
								oCmdActualiza.ExecuteNonQuery();
							}
							if ((int) dr.Cells["idFormaPago"].Value == 16  )  
							{
								sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
									(int)dr.Cells["idDetalleCaja"].Value, idCaja, 0, (int)dr.Cells["idFormaPago"].Value, "",  
									Convert.ToDateTime(dr.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)dr.Cells["Valor"].Value * (-1));
								oCmdActualiza.CommandText = sSQLDetalleEf;
								oCmdActualiza.ExecuteNonQuery();
							}
						}					
					}
					
					oCmdActualiza.CommandText = string.Format("Exec ConsultaDetalleCaja {0}", idCaja);
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.uGridDepositos.DataSource = oDTDetalle;
					
					oCmdActualiza.CommandText = string.Format("Exec MensajeCaja {0}", idCaja);
					this.lblEstadoDepositos.Text = oCmdActualiza.ExecuteScalar().ToString();
					#endregion Detalle

					#region Crea Asiento
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
					{
						bool bNdcNDb = false;
						string sSQL = "";
						if ((int) dr.Cells["idFormaPago"].Value == 4 || (int) dr.Cells["idFormaPago"].Value == 16 || (int) dr.Cells["idFormaPago"].Value == 3   ) bNdcNDb = true;  

						if(!bNdcNDb)
						{
//							if ((int)dr.Cells["idAsiento"].Value == 0)
//							{
								sSQL =  string.Format("Exec CreaAsiCierreCaja {0}, {1}", (int)dr.Cells["idDetalleCaja"].Value, (int)dr.Cells["idCaja"].Value);						
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();
//							}
						}
						else
						{
							if ((int) dr.Cells["idFormaPago"].Value != 3)
							{
								sSQL = string.Format("Exec CreaAsiCierreCajaNDCDB {0}, {1}, {2}", (int)dr.Cells["idDetalleCaja"].Value, (int)dr.Cells["idCaja"].Value, (int)dr.Cells["idFormaPago"].Value);						
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();
							}
						}					
					}			

					string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCajas {0}", idCaja);
					oCmdActualiza.CommandText = sSQLDiferencias;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Crea Asiento
					
					oTransaction.Commit();						

					oConexion.Close();								

					MessageBox.Show(Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeCaja {0}", idCaja)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);		

					this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCaja {0}", idCaja));		
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					oConexion.Close();								

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
				}
				finally
				{
					oConexion.Close();								
				}
			}
			
			Cursor = Cursors.Default;		
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			string sMensaje = "";
			string sNumero = "";

			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros de caja para actualizar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
				return;
			}
			
			#region Valida Numeracion Existente
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["Origen"].Value.ToString() == "FAC")
				{
					char Pad = '0';
					sNumero = dr.Cells["Factura"].Value.ToString().PadLeft(9, Pad);
					
					string sSQL = string.Format("Exec ValidaNumeracionExistente {0}, '{1}', {2}", (int)dr.Cells["idCompra"].Value, sNumero, (int)this.cmbBodega.Value);
					
					if (Funcion.sEscalarSQL(cdsSeteoF, sSQL, true).ToString().Length > 0)
					{
						MessageBox.Show(Funcion.sEscalarSQL(cdsSeteoF, sSQL, true).ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;						
						return;
					}
				}
				else if (dr.Cells["Origen"].Value.ToString() != "PAG")
				{
					char Pad = '0';
					sNumero = dr.Cells["Factura"].Value.ToString().PadLeft(7, Pad);
					
					string sSQL = string.Format("Select COUNT(*) From Anticipos Where Secuencial = '{1}' And Bodega = {2} And idAnticipo <> {0}", (int)dr.Cells["idCompra"].Value, sNumero, (int)this.cmbBodega.Value);
					
					if (Funcion.iEscalarSQL(cdsSeteoF, sSQL, true) > 0)
					{
						MessageBox.Show(string.Format("El Numero de Recibo '{0}' ya existe", sNumero), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;						
						return;
					}
				}
			}
			#endregion Valida Numeracion Existente
				
			#region Actualizacion
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				int idTipoFactura = 0;

				if (dr.Cells["Origen"].Value.ToString() == "FAC") idTipoFactura = 1;
				else if (dr.Cells["Origen"].Value.ToString() == "CP" || dr.Cells["Origen"].Value.ToString() == "ANT" || dr.Cells["Origen"].Value.ToString() == "AB") idTipoFactura = 24;

				string sSQL = string.Format("Exec ActualizaNumeracionCuadrecaja {0}, {1}, '{2}', '{3}'", 
					(int)dr.Cells["idCompra"].Value, idTipoFactura, dr.Cells["Factura"].Value.ToString(), Convert.ToDateTime(dr.Cells["Fecha"].Value).ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
			#endregion Actualizacion

			this.btnVer_Click(sender, e);			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			this.btnVer_Click(sender, e);

			int iEst = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select EstadoCaja From Caja Where idCaja = {0}", idCaja));

			#region Validaciones
			if (iEst != iEstadoCaja)
			{
				MessageBox.Show("El registro de esta caja ha sido modificada por otro usuario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.btnVer_Click(sender, e);
				return;
			}
			
			if (iEstadoCaja == 1)
			{
				//decimal dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				//200dDescuento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaDescuentoPromociones {0}, {1}", idArticulo, idTPrecio));
				#region Caja
				#region Valida depositos segun forma de pago
				#region Validda Depostitos registrados
				decimal dValorDepositado = 0.00m;
				valorDepositos=0.00m;
				valorDepositos = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaValorDeposito {0}", idCaja));

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
					if ((int)dr.Cells["idFormaPago"].Value == 1)
						dValorDepositado = dValorDepositado + (decimal)dr.Cells["Valor"].Value;

				if (valorDepositos != dValorDepositado)
				{
					//if (valorDepositos < dValorDepositado)
					//{
						MessageBox.Show("No es posible cerrar la caja hasta que se registren todos los depósitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					//}
				}
				#endregion Validda Depostitos registrados

				#region efectivo
				if ((decimal)this.spnCEfectivo.Value > 0.00m)
				{
					int iCont = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						if ((int)dr.Cells["idFormaPago"].Value == 1 || (int)dr.Cells["idFormaPago"].Value == 4) iCont++;

					if (iCont == 0)
					{
						MessageBox.Show("No se ha registrado el deposito en efectivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;
					}
				}
				#endregion efectivo

				#region Cheque
				if ((decimal)this.spnCCheque.Value > 0.00m)
				{
					int iCont = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						if ((int)dr.Cells["idFormaPago"].Value == 2) iCont++;

					if (iCont == 0)
					{
						MessageBox.Show("No se ha registrado el deposito en cheques", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;
					}
				}
				#endregion Cheque

				#region Entrega
				if ((decimal)this.spnC_Entrega.Value > 0.00m)
				{
					int iCont = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						if ((int)dr.Cells["idFormaPago"].Value == 19) iCont++;

					if (iCont == 0)
					{						
						MessageBox.Show("No se ha registrado el deposito en efectivo por contraentrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;
					}
				}
				#endregion Entrega
				#region CTC

				if ((decimal)this.spnCTC.Value > 0.00m)
				{
					int iCont = 0;
					decimal dTC = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						if ((int)dr.Cells["idFormaPago"].Value == 3) iCont++;
					

					if (iCont == 0)
					{						
						MessageBox.Show("No se ha registrado el item del DATAFAST", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;
					}
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						if ((int)dr.Cells["idFormaPago"].Value == 3) dTC = dTC + (decimal)dr.Cells["Valor"].Value;

					if(dTC < (decimal)this.spnCTC.Value)
					{
						MessageBox.Show("Faltante TC/TD por ingresar  ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;

					}
				}
				#endregion CTC

				#endregion Valida depositos segun forma de pago

				#region Valida Pago Agil
				int iContPAG = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)				
					if (dr.Cells["Origen"].Value.ToString() == "PAG" && dr.Cells["Estado"].Value.ToString() != "ANULADO")
						if (!(bool)dr.Cells["R"].Value) iContPAG++;

				if (iContPAG > 0)
				{
					MessageBox.Show(string.Format("No se ha revisado {0} cobros de pago agil", iContPAG), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}
				#endregion Valida Pago Agil

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
				{
					string sSQL = string.Format("Exec CierreDeCajaValidaEstadoDepositos {0}", (int)dr.Cells["idAsiento"].Value);
					if (!Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
					{
						MessageBox.Show("El deposito no ha sido conciliado, No puede cerrar la caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;
					}
				}				

				//				if (iEstado == 1)
				//				{
				//					MessageBox.Show("No se ha registrado el deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				//					return;
				//				}

				if (iEstado == 4)
				{
					MessageBox.Show(this.lblEstadoDepositos.Text.ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}
				#endregion Caja
			}
			#endregion Validaciones

			string sMensaje = "Cerrar";
			if (iEstadoCaja == 2) sMensaje = "Abrir";

			if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Esta seguro de '{0}' esta caja?", sMensaje), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				int iNuevoEstado = 1;
				if (iEstadoCaja == 1) iNuevoEstado = 2;

				string sSQL = string.Format("Exec RevisadoCaja {0}, '{1}', {2}, {3}", (bool)this.chkRevisado.Checked, this.txtObservaciones.Text.ToString(), idCaja, iNuevoEstado);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				
				if (iEstadoCaja == 1)
				{
					string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCajas {0}", idCaja);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDiferencias);
				}

				MessageBox.Show("Proceso terminado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);		
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)			
				if (this.cmbBodega.Value != System.DBNull.Value)						
					this.cmbVendedores.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
		    	this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("select C.idCuenta, C.Descripcion from BodegaCuentasBancarias B  INNER JOIN Cuenta C ON C.idCuenta = B.idCuenta WHERE B.idBodega =  {0}", (int)this.cmbBodega.Value));
			    this.cmbFormaPago_1.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("SELECT idFormaPago AS idFormaPago_1, FormaPago AS  Descripcion FROM CompraFormaPago  "));
			
				
			if (this.cmbBodega.ActiveRow != null)					
				this.btnVer_Click(sender, e);				
		}
		
		bool bActualiza = false;

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "idPersonal")
			{	
				bActualiza = true;

				if ((int) ultraGrid1.ActiveRow.Cells["idOrigen"].Value == 1)
				{					
					if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar el Vendedor?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{						
						string sSQL = string.Format("UPDATE Compra SET idVendedor = {0} WHERE idCompra = {1} ", 
							(int)	this.ultraGrid1.ActiveRow.Cells["idPersonal"].Value, (int)	this.ultraGrid1.ActiveRow.Cells["idCompra"].Value );
						Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
					}
					else
					{						
						e.Cell.Row.Cells["idPersonal"].Value = (int)e.Cell.OriginalValue;
						//return;
					}					
				}	
			
				bActualiza = false;
			}	

			if (e.Cell.Column.ToString() == "idFormaPago_1")
			{	
				bActualiza = true;

				if ((int) ultraGrid1.ActiveRow.Cells["idOrigen"].Value == 1 )
				{	
					if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 3 || (int) ultraGrid1.ActiveRow.Cells["idFormaPago"].Value == 11 )
					{
						if ((int) ultraGrid1.ActiveRow.Cells["idFormaPago_1"].Value == 3 || (int) ultraGrid1.ActiveRow.Cells["idFormaPago_1"].Value == 11 )
						{
							if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar la Forma De Pago?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
							{						
								string sSQL = string.Format("UPDATE Compra SET idFormaPago = {0} WHERE idCompra = {1} ", 
									(int)	this.ultraGrid1.ActiveRow.Cells["idFormaPago_1"].Value, (int)	this.ultraGrid1.ActiveRow.Cells["idCompra"].Value );
								Funcion.EjecutaSQL(cdsSeteoF, sSQL);						
							}
							else
							{						
								e.Cell.Row.Cells["idFormaPago_1"].Value = (int)e.Cell.OriginalValue;
								//return;
							}					
						}
						else
						{						
							e.Cell.Row.Cells["idFormaPago_1"].Value = (int)e.Cell.OriginalValue;
							//return;
						}	
					}
					else
					{
						e.Cell.Row.Cells["idFormaPago_1"].Value = (int)e.Cell.OriginalValue;
					}
				}
				else
				{
					e.Cell.Row.Cells["idFormaPago_1"].Value = (int)e.Cell.OriginalValue;
				}
			
				bActualiza = false;
			}	

			
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Almacen")) return;

			using (frmBodegaCuentasBancarias BCB = new frmBodegaCuentasBancarias((int)this.cmbBodega.Value))
			{
				if (DialogResult.OK == BCB.ShowDialog())
				{
			
				}			
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		
		//bool bActualizaB = false;

		private void uGridDepositos_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			//			string sColumna = e.Cell.Column.ToString();
			//
			//			if (e.Cell.Column.ToString() == sColumna)
			//				if (e.Cell.Row.Cells["idAsiento"].Value != System.DBNull.Value)
			//					if ((int)e.Cell.Row.Cells["idAsiento"].Value > 0) e.Cancel = true;
			if (e.Cell.Row.Cells["FechaDeposito"].Value != System.DBNull.Value)
			{
				if (bActualiza) return;
				bActualiza = true;
				if (e.Cell.Column.ToString() == "FechaDeposito")
				{		
					FechaCajaBefore = (DateTime) e.Cell.Row.Cells["FechaDeposito"].Value ;
					bActualiza = false;
				}
			}
		}

		private void txtObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void uGridDepositos_AfterRowFilterChanged(object sender, Infragistics.Win.UltraWinGrid.AfterRowFilterChangedEventArgs e)
		{
		
		}

		private void uGridDepositos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{						
			if (e.Cell.Column.ToString() == "BtnDeposito")	
			{
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", true, "Bancos", cdsSeteoF)) return;
			
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Caja")) return;

				if (e.Cell.Row.Cells["Valor"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el valor del deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) == 0.00m)
				{
					MessageBox.Show("Ingrese el valor del deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
								
				#region Grids
				if ((decimal)this.spnCEfectivo.Value == 0.00m && (decimal)this.spnCCheque.Value == 0.00m && (decimal)this.spnC_Entrega.Value == 0.00m && (decimal)this.spnCTC.Value == 0.00m)
				{
					MessageBox.Show("No existen Valores para depositar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}			
					
				#region Validacion Detalle
				bool bNdcNDb = false;
				
				if (e.Cell.Row.Cells["idFormaPago"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la forma de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value)  == 4 || Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value)  == 16 || Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value)  == 3)
				{
					bNdcNDb = true;  
				}

				if (!bNdcNDb)
				{
					if (e.Cell.Row.Cells["idCuenta"].Value  == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione la cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					if ((int)e.Cell.Row.Cells["idCuenta"].Value  == 0)
					{
						MessageBox.Show("Seleccione la cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}		
				}
				if (e.Cell.Row.Cells["FechaDeposito"].Value  == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la fecha de depósito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToInt32(e.Cell.Row.Cells["idDetalleCaja"].Value) == 0)
				{		
					if (Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value) < Convert.ToDateTime(this.dtFecha.Value))
					{
						MessageBox.Show("La Fecha del Deposito No puede ser Anterior a la Fecha de la Caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if (Funcion.ValidaPeriodo((DateTime)e.Cell.Row.Cells["FechaDeposito"].Value, cdsSeteoF, "Bancos"))
				{						
					DateTime dtFechaVP = (DateTime)e.Cell.Row.Cells["FechaDeposito"].Value;
					MessageBox.Show(string.Format("No puede Registrar depositos en esta Fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
					return;
				}
				if (!bNdcNDb)
				{
					string sNumCB = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDepositoCajaConciliado {0}, {1}", Convert.ToInt32(e.Cell.Row.Cells["idAsiento"].Value), Convert.ToInt32(e.Cell.Row.Cells["idCuenta"].Value)));
					if (sNumCB.Length > 0)
					{
						MessageBox.Show("El deposito esta Conciliado, Consulte con el Departamento de Contabilidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				
					if (e.Cell.Row.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Escriba el número de depósito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if ((int) e.Cell.Row.Cells["idFormaPago"].Value  == 3)
				{
					if (e.Cell.Row.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Escriba el número de LOTE", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
				if (e.Cell.Row.Cells["Valor"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el valor del deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) == 0.00m)
				{
					MessageBox.Show("Ingrese el valor del deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if (!bNdcNDb && (int) e.Cell.Row.Cells["idFormaPago"].Value != 3)
				{
					if ((int)e.Cell.Row.Cells["idAsiento"].Value  == 0)
					{
						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeroDocumento {0}, '{1}', 2, {2}", (int)e.Cell.Row.Cells["idCuenta"].Value , e.Cell.Row.Cells["Numero"].Value.ToString(), (int) e.Cell.Row.Cells["idAsiento"].Value ), true))
						{
							MessageBox.Show("El numero de deposito ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}

				DateTime dFecha = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF); 
				if ((int)e.Cell.Row.Cells["idDetalleCaja"].Value  == 0)
				{
					if (MenuLatinium.stDirFacturacion != "DENTALIS")
					{
						//DateTime.Today;
						if (Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).Date != dFecha.Date)
						{					
							MessageBox.Show("La Fecha del DEPOSITO no puede ser diferente al dia actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
					}
					else
					{
						if (!miAcceso.CambiarFechaPosterior)
						{
							MessageBox.Show("La Fecha del DEPOSITO no puede ser diferente al dia actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							return;
						}
					}
				}

				#endregion Validacion Detalle
			
				#region Proyecto
				string sSQLProyecto = string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value);
				int iProyecto = Funcion.iEscalarSQL(cdsSeteoF, sSQLProyecto, true);
				#endregion Proyecto

				decimal dValDeposito = 0.00m;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
					dValDeposito = dValDeposito + Convert.ToDecimal(dr.Cells["Valor"].Value); 
			
				decimal dEfectivo = 0.00m;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					dEfectivo = dEfectivo + Convert.ToDecimal(dr.Cells["Efectivo"].Value); 

				if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 1 )
				{
					#region Guarda deposito de movimientos en efectivo
					using (frmReporteValidaDeposito Tarjetas = new frmReporteValidaDeposito ( Convert.ToDecimal ( e.Cell.Row.Cells["Valor"].Value ),(DateTime) this.dtFecha.Value, (int) this.cmbBodega.Value, (int)e.Cell.Row.Cells["idDetalleCaja"].Value, dEfectivo, idCaja))
					{
						if (DialogResult.OK == Tarjetas.ShowDialog())
						{
							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
									string sSQLMaestro = string.Format("Exec GuardaCaja {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}", 
										idCaja, (int)this.cmbBodega.Value, iProyecto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), "", "", 
										(decimal)this.spnCEfectivo.Value, (decimal)this.spnCCheque.Value, dValDeposito, 1, (bool)this.chkRevisado.Checked, this.txtObservaciones.Text.ToString(), (decimal)this.spnC_Entrega.Value, (decimal)this.spnCTC.Value);			
									oCmdActualiza.CommandText = sSQLMaestro;
									idCaja = (int)oCmdActualiza.ExecuteScalar();
			
									#region Detalle
					
									bool bNdcNDba = false;
									if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 4 || (int) (int)e.Cell.Row.Cells["idFormaPago"].Value == 16 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 3  ) bNdcNDba = true;  
									string sSQLDetalleEf = "";
									if(!bNdcNDba)
									{
										sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
											(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, (int)e.Cell.Row.Cells["idCuenta"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value, e.Cell.Row.Cells["Numero"].Value.ToString(),  
											Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value);
										oCmdActualiza.CommandText = sSQLDetalleEf;
										e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();
									}
									else
									{
										if ((int) e.Cell.Row.Cells["idFormaPago"].Value == 4  )  
										{
											sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
												(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, 0, (int)e.Cell.Row.Cells["idFormaPago"].Value, "",  
												Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value);
											oCmdActualiza.CommandText = sSQLDetalleEf;
											e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();
										}
										if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 3 )  
										{
											sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
												(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, 0, (int)e.Cell.Row.Cells["idFormaPago"].Value, e.Cell.Row.Cells["Numero"].Value.ToString(),  
												Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value);
											oCmdActualiza.CommandText = sSQLDetalleEf;
											e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();
										}
										if ((int) e.Cell.Row.Cells["idFormaPago"].Value == 16  )  
										{
											sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
												(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, 0, (int)e.Cell.Row.Cells["idFormaPago"].Value, "",  
												Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value * (-1));
											oCmdActualiza.CommandText = sSQLDetalleEf;
											e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();
										}
									}	
				
									bool bNdcNDbb = false;
									string sSQL = "";
									if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 4 ) bNdcNDbb = true;  
									if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 3 ) bNdcNDbb = true; 
									if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 16 ) bNdcNDbb = true; 
									//	if ( Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value) == 4  ||  Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value) == 16 ||  Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value) == 3   ) bNdcNDbb = true;  

									if(!bNdcNDbb)
									{
										//							if ((int)dr.Cells["idAsiento"].Value == 0)
										//							{
										sSQL =  string.Format("Exec CreaAsiCierreCaja {0}, {1}", (int)e.Cell.Row.Cells["idDetalleCaja"].Value , idCaja);						
										oCmdActualiza.CommandText = sSQL;
										oCmdActualiza.ExecuteNonQuery();
										//							}
									}
									else
									{
										if ((int) e.Cell.Row.Cells["idFormaPago"].Value != 3)
										{
											sSQL = string.Format("Exec CreaAsiCierreCajaNDCDB {0}, {1}, {2}", (int)e.Cell.Row.Cells["idDetalleCaja"].Value , idCaja, (int)e.Cell.Row.Cells["idFormaPago"].Value);						
											oCmdActualiza.CommandText = sSQL;
											oCmdActualiza.ExecuteNonQuery();
										}
									}		

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Tarjetas.uGridDepositos.Rows.All)
									{
										if ( (decimal)dr.Cells["Valor"].Value > 0 )
										{
											string  sSQLDeta = string.Format("Exec GuardaValidaDeposito {0}, {1}, {2}, {3}, {4}" ,
												0, 
												idCaja, (int)e.Cell.Row.Cells["idDetalleCaja"].Value, 
												(int)dr.Cells["idDetalleFormaPago"].Value, (decimal)dr.Cells["Valor"].Value);
											oCmdActualiza.CommandText = sSQLDeta;
											oCmdActualiza.ExecuteNonQuery();
										}
									}
													
									oCmdActualiza.CommandText = string.Format("Exec ConsultaDetalleCaja {0}", idCaja);
									SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
									DataTable oDTDetalle = new DataTable();
									oDADetalle.Fill(oDTDetalle);
									this.uGridDepositos.DataSource = oDTDetalle;
					
									oCmdActualiza.CommandText = string.Format("Exec MensajeCaja {0}", idCaja);
									this.lblEstadoDepositos.Text = oCmdActualiza.ExecuteScalar().ToString();
									#endregion Detalle

									#region Crea Asiento
									string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCaja {0}", idCaja);
									oCmdActualiza.CommandText = sSQLDiferencias;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Crea Asiento

									oTransaction.Commit();						

									oConexion.Close();								

									//MessageBox.Show(Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeCaja {0}", idCaja)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);		

									this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCaja {0}", idCaja));		
								}
								catch (Exception ex)
								{
									oTransaction.Rollback();

									oConexion.Close();								

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
								}
								finally
								{
									oConexion.Close();								
								}
							}
						}		
					}			
					#endregion Guarda deposito de movimientos en efectivo
				}
				else
				{
					#region guarda
					
					using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
							
							string sSQLMaestro = string.Format("Exec GuardaCaja {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}", 
								idCaja, (int)this.cmbBodega.Value, iProyecto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), "", "", 
								(decimal)this.spnCEfectivo.Value, (decimal)this.spnCCheque.Value, dValDeposito, 1, (bool)this.chkRevisado.Checked, this.txtObservaciones.Text.ToString(), (decimal)this.spnC_Entrega.Value, (decimal)this.spnCTC.Value);			
							oCmdActualiza.CommandText = sSQLMaestro;
							idCaja = (int)oCmdActualiza.ExecuteScalar();
			
							#region Detalle
					
							bool bNdcNDba = false;
							if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 4 || (int) (int)e.Cell.Row.Cells["idFormaPago"].Value == 16 || (int)e.Cell.Row.Cells["idFormaPago"].Value == 3  ) bNdcNDba = true;  
							string sSQLDetalleEf = "";
							if(!bNdcNDba)
							{
								//(int)e.Cell.Row.Cells["idFormaPago"].Value
								sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
									(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, (int)e.Cell.Row.Cells["idCuenta"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value, e.Cell.Row.Cells["Numero"].Value.ToString(),  
									Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value);

								oCmdActualiza.CommandText = sSQLDetalleEf;
								e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();
							}
							else
							{
								if ((int) e.Cell.Row.Cells["idFormaPago"].Value == 4  )  
								{
									sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
										(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, 0, (int)e.Cell.Row.Cells["idFormaPago"].Value, "",  
										Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value);
									oCmdActualiza.CommandText = sSQLDetalleEf;
									e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();

								}
								if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 3 )  
								{
									sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
										(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, 0, (int)e.Cell.Row.Cells["idFormaPago"].Value, e.Cell.Row.Cells["Numero"].Value.ToString(),  
										Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value);
									oCmdActualiza.CommandText = sSQLDetalleEf;
									e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();

								}
								if ((int) e.Cell.Row.Cells["idFormaPago"].Value == 16  )  
								{
									sSQLDetalleEf = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
										(int)e.Cell.Row.Cells["idDetalleCaja"].Value, idCaja, 0, (int)e.Cell.Row.Cells["idFormaPago"].Value, "",  
										Convert.ToDateTime(e.Cell.Row.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), (decimal)e.Cell.Row.Cells["Valor"].Value * (-1));
									oCmdActualiza.CommandText = sSQLDetalleEf;
									e.Cell.Row.Cells["idDetalleCaja"].Value = (int)oCmdActualiza.ExecuteScalar();

								}
							}	
				
							bool bNdcNDbb = false;
							string sSQL = "";
							if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 4 ) bNdcNDbb = true;  
							if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 3 ) bNdcNDbb = true; 
							if ( (int) e.Cell.Row.Cells["idFormaPago"].Value == 16 ) bNdcNDbb = true; 
							//	if ( Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value) == 4  ||  Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value) == 16 ||  Convert.ToInt32( e.Cell.Row.Cells["idFormaPago"].Value) == 3   ) bNdcNDbb = true;  

							if(!bNdcNDbb)
							{
								//							if ((int)dr.Cells["idAsiento"].Value == 0)
								//							{
								sSQL =  string.Format("Exec CreaAsiCierreCaja {0}, {1}", (int)e.Cell.Row.Cells["idDetalleCaja"].Value , idCaja);						
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();
								//							}
							}
							else
							{
								if ((int) e.Cell.Row.Cells["idFormaPago"].Value != 3)
								{
									sSQL = string.Format("Exec CreaAsiCierreCajaNDCDB {0}, {1}, {2}", (int)e.Cell.Row.Cells["idDetalleCaja"].Value , idCaja, (int)e.Cell.Row.Cells["idFormaPago"].Value);						
									oCmdActualiza.CommandText = sSQL;
									oCmdActualiza.ExecuteNonQuery();
								}
							}		
							oCmdActualiza.CommandText = string.Format("Exec ConsultaDetalleCaja {0}", idCaja);
							SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
							DataTable oDTDetalle = new DataTable();
							oDADetalle.Fill(oDTDetalle);
							this.uGridDepositos.DataSource = oDTDetalle;
					
							oCmdActualiza.CommandText = string.Format("Exec MensajeCaja {0}", idCaja);
							this.lblEstadoDepositos.Text = oCmdActualiza.ExecuteScalar().ToString();
							#endregion Detalle

							#region Crea Asiento

							string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCaja {0}", idCaja);
							oCmdActualiza.CommandText = sSQLDiferencias;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Crea Asiento

							oTransaction.Commit();						

							oConexion.Close();								

							//MessageBox.Show(Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeCaja {0}", idCaja)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);		

							this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCaja {0}", idCaja));		
						}
						catch (Exception ex)
						{
							oTransaction.Rollback();

							oConexion.Close();								

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
						finally
						{
							oConexion.Close();								
						}
					}
			
					#endregion guarda
				}
				Cursor = Cursors.Default;		
				#endregion Grids 

        int iContCajas = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drA in this.uGridDepositos.Rows.All)
				{
					if ((int) drA.Cells["idFormaPago"].Value == 1 && drA.Cells["idDetalleCaja"].Value != System.DBNull.Value && (int)drA.Cells["idDetalleCaja"].Value > 0)
					{
						iContCajas =  iContCajas + 1;
					}
				}

				int iCajaCapturadas = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dre in this.uGridDepositos.Rows.All)
				{
					if ((int) dre.Cells["idFormaPago"].Value == 1 && dre.Cells["idDetalleCaja"].Value != System.DBNull.Value && (int)dre.Cells["idDetalleCaja"].Value > 0)
					{
						int icapturadas = 0;
						icapturadas = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT count(1) FROM ValidaDeposito where idDetalleCaja  = {0}",
							(int) dre.Cells["idDetalleCaja"].Value));	
						if (icapturadas > 0)
						{
							iCajaCapturadas =  iCajaCapturadas + 1;
						}
					}
				}
				if (iContCajas > 0 &&  iCajaCapturadas > 0)
				{
					if ( iContCajas == iCajaCapturadas )
					{
						this.btnDepositar_Click (sender, e);
					}
				}
			}
		}
	}
}

