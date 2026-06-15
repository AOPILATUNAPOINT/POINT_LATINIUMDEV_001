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
	/// Summary description for frmCobranzas.
	/// </summary>
	
	
	public class frmCobranzas : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraLabel ultraLabel4;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.Misc.UltraLabel ultraLabel8;
		private Infragistics.Win.Misc.UltraLabel ultraLabel9;
		private Infragistics.Win.Misc.UltraLabel ultraLabel10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaVencida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoVencido;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProximaVencer;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSiguienteCuota;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCobro;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid GridFacturas;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridPagos;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;

		private C1DataRow drSeteoF;
		private Infragistics.Win.Misc.UltraLabel ultraLabel12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorAutorizado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAutorizacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstadoAutorizacion;
		private System.Windows.Forms.Button btCobro;
		private System.Windows.Forms.Button btnAutorizar;
		private System.Windows.Forms.Button btnSalir;
		private Acceso miAcceso;
		bool bMayorCobro = false;
		private System.Windows.Forms.Button btnActualizarEDsctoGM;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.Misc.UltraLabel lblFecha;
		private System.Windows.Forms.Button btnAnticipo;
		private System.Windows.Forms.Button btnReliquidar;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasVencidos;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		decimal dValorCobro = 0;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoAFavor;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.Misc.UltraLabel ultraLabel5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProntoPago;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel11;
		private Infragistics.Win.Misc.UltraLabel ultraLabel13;
		private Infragistics.Win.Misc.UltraLabel ultraLabel14;
		private System.Windows.Forms.Button btnHistorico;
		private System.Windows.Forms.Button btnRefinanciacion;
		private System.Windows.Forms.ContextMenu mnuRefinanciacion;
		private System.Windows.Forms.MenuItem mnuDocumentos;
		private System.Windows.Forms.MenuItem mnuGenerarCXC;
		private System.Windows.Forms.Button btnCondonar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel15;
		private Infragistics.Win.Misc.UltraLabel ultraLabel16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorARefinanciar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorACondonar;
		int iDiasVencidos = 0;
		int  id_Anticipo=0;
		private Infragistics.Win.Misc.UltraLabel ultraLabel17;
		private System.Windows.Forms.Button btnSNoDif;
		private System.Windows.Forms.Button btnAux;
		decimal Val_Cuota =0.00m; 
		public static int idBonoEntradaTipo25 = 1;
		string sArchivoBonoEntrada25 = "BonoEntrada25.rpt";
		int NumeracionLocalesBonoEntrada = 92;
		string ReciboImpresion = "ImpresionReciboCrediPOINT_2024_V1.rpt";

		public frmCobranzas()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		string sCedula ;
		bool bPantalla = false;
		
		public frmCobranzas(string SCedula , bool BPantalla)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			sCedula = SCedula;
			bPantalla = BPantalla;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeOperacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reliquido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ER");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asignado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeOperacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Reliquido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ER");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Asignado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vence");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuota");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InteresTotal");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalTotal");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABCapital");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gastos");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DsctoGastos");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DsctoIntMora");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retraso");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCancelacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Secuencial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalDiferimiento");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferimiento");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Abono", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Abono", 20, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCuota", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCuota", 6, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 8, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 11, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Gastos", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Gastos", 13, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Mora", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Mora", 15, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 17, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "InteresTotal", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "InteresTotal", 7, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalTotal", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalTotal", 10, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalDiferimiento", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalDiferimiento", 27, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferimiento", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferimiento", 28, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Retraso", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retraso", 18, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vence");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InteresTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABInteres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DsctoGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DsctoIntMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCancelacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Secuencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalDiferimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCobranzas));
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinEditors.EditorButton editorButton1 = new Infragistics.Win.UltraWinEditors.EditorButton();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			this.ultraLabel4 = new Infragistics.Win.Misc.UltraLabel();
			this.GridFacturas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.gridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel8 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel9 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel10 = new Infragistics.Win.Misc.UltraLabel();
			this.txtSaldoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaVencida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProximaVencer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSiguienteCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCobro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblFecha = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel12 = new Infragistics.Win.Misc.UltraLabel();
			this.txtValorAutorizado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAutorizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstadoAutorizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btCobro = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnActualizarEDsctoGM = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.btnAnticipo = new System.Windows.Forms.Button();
			this.btnReliquidar = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDiasVencidos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtSaldoAFavor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtProntoPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel5 = new Infragistics.Win.Misc.UltraLabel();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.ultraLabel11 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel13 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel14 = new Infragistics.Win.Misc.UltraLabel();
			this.btnHistorico = new System.Windows.Forms.Button();
			this.btnRefinanciacion = new System.Windows.Forms.Button();
			this.mnuRefinanciacion = new System.Windows.Forms.ContextMenu();
			this.mnuDocumentos = new System.Windows.Forms.MenuItem();
			this.mnuGenerarCXC = new System.Windows.Forms.MenuItem();
			this.btnCondonar = new System.Windows.Forms.Button();
			this.txtValorACondonar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel15 = new Infragistics.Win.Misc.UltraLabel();
			this.txtValorARefinanciar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraLabel16 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel17 = new Infragistics.Win.Misc.UltraLabel();
			this.btnSNoDif = new System.Windows.Forms.Button();
			this.btnAux = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.GridFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaVencida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProximaVencer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSiguienteCuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorAutorizado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAFavor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorACondonar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorARefinanciar)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraLabel4
			// 
			this.ultraLabel4.AutoSize = true;
			this.ultraLabel4.Location = new System.Drawing.Point(8, 60);
			this.ultraLabel4.Name = "ultraLabel4";
			this.ultraLabel4.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel4.TabIndex = 154;
			this.ultraLabel4.Text = "Cédula";
			// 
			// GridFacturas
			// 
			this.GridFacturas.Cursor = System.Windows.Forms.Cursors.Default;
			this.GridFacturas.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.GridFacturas.DisplayLayout.Appearance = appearance1;
			this.GridFacturas.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 95;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.Caption = "Documento";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 100;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.Caption = "Cuotas";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 140;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.Caption = "Lote";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 83;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 14;
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
																										 ultraGridColumn15});
			ultraGridBand1.Header.Caption = "Facturas";
			ultraGridBand1.HeaderVisible = true;
			this.GridFacturas.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.GridFacturas.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.GridFacturas.DisplayLayout.Override.CardAreaAppearance = appearance6;
			this.GridFacturas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.GridFacturas.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance8.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.GridFacturas.DisplayLayout.Override.RowAlternateAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.GridFacturas.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.GridFacturas.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.GridFacturas.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.GridFacturas.Location = new System.Drawing.Point(8, 128);
			this.GridFacturas.Name = "GridFacturas";
			this.GridFacturas.Size = new System.Drawing.Size(552, 160);
			this.GridFacturas.TabIndex = 167;
			this.GridFacturas.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.GridFacturas_InitializeRow);
			this.GridFacturas.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.GridFacturas_BeforeSelectChange);
			this.GridFacturas.Click += new System.EventHandler(this.GridFacturas_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn3.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn4.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			ultraDataColumn12.DataType = typeof(bool);
			ultraDataColumn12.DefaultValue = false;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
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
																																 ultraDataColumn15});
			// 
			// gridPagos
			// 
			this.gridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridPagos.CausesValidation = false;
			this.gridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.gridPagos.DataSource = this.ultraDataSource2;
			appearance11.BackColor = System.Drawing.Color.White;
			this.gridPagos.DisplayLayout.Appearance = appearance11;
			this.gridPagos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 105;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 76;
			ultraGridColumn18.Header.VisiblePosition = 2;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 53;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance12;
			ultraGridColumn19.Header.Caption = "Cuota";
			ultraGridColumn19.Header.VisiblePosition = 3;
			ultraGridColumn19.Width = 56;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 107;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.VisiblePosition = 5;
			ultraGridColumn21.Width = 67;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance13;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Valor";
			ultraGridColumn22.Header.VisiblePosition = 6;
			ultraGridColumn22.Width = 68;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance14;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Total Interes";
			ultraGridColumn23.Header.VisiblePosition = 7;
			ultraGridColumn23.Width = 72;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance15;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "Interés";
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn24.Width = 72;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 12;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance16;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "Total Capital";
			ultraGridColumn26.Header.VisiblePosition = 8;
			ultraGridColumn26.Width = 72;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance17;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 11;
			ultraGridColumn27.Width = 72;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 17;
			ultraGridColumn28.Hidden = true;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance18;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Gestión";
			ultraGridColumn29.Header.VisiblePosition = 13;
			ultraGridColumn29.MaskInput = "";
			ultraGridColumn29.PromptChar = ' ';
			ultraGridColumn29.Width = 74;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance19;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "% Dscto";
			ultraGridColumn30.Header.VisiblePosition = 14;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.PromptChar = ' ';
			ultraGridColumn30.Width = 66;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance20;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 15;
			ultraGridColumn31.MaskInput = "";
			ultraGridColumn31.PromptChar = ' ';
			ultraGridColumn31.Width = 74;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance21;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "% Dscto";
			ultraGridColumn32.Header.VisiblePosition = 18;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.PromptChar = ' ';
			ultraGridColumn32.Width = 72;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance22;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 19;
			ultraGridColumn33.Width = 74;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance23;
			ultraGridColumn34.Header.VisiblePosition = 22;
			ultraGridColumn34.Width = 60;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 21;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 57;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance24;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 20;
			ultraGridColumn36.Width = 74;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.Caption = "Pago";
			ultraGridColumn37.Header.VisiblePosition = 23;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 66;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.Caption = "Documento N.";
			ultraGridColumn38.Header.VisiblePosition = 24;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 87;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.Caption = "Forma de pago";
			ultraGridColumn39.Header.VisiblePosition = 26;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 107;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 27;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 108;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.Caption = "Local";
			ultraGridColumn41.Header.VisiblePosition = 25;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 183;
			ultraGridColumn42.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn42.Header.VisiblePosition = 28;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 207;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance25;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "T Diferimiento";
			ultraGridColumn43.Header.VisiblePosition = 9;
			ultraGridColumn43.Width = 78;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance26;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 16;
			ultraGridColumn44.Width = 74;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn45.Header.VisiblePosition = 29;
			ultraGridColumn45.Width = 32;
			ultraGridBand2.Columns.AddRange(new object[] {
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
			ultraGridBand2.Header.Caption = "";
			ultraGridBand2.HeaderVisible = true;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand2.Override.ActiveRowAppearance = appearance27;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance28;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance29;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance30;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance31;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance32;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance33;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance34;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance35;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance36;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance37;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance38;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings12.Appearance = appearance39;
			summarySettings12.DisplayFormat = "{0: #,##0}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings12});
			ultraGridBand2.SummaryFooterCaption = "";
			this.gridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance40.BackColor = System.Drawing.Color.Transparent;
			this.gridPagos.DisplayLayout.Override.CardAreaAppearance = appearance40;
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.gridPagos.DisplayLayout.Override.CellAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.FontData.BoldAsString = "True";
			appearance42.FontData.Name = "Arial";
			appearance42.FontData.SizeInPoints = 8F;
			appearance42.ForeColor = System.Drawing.Color.White;
			appearance42.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.gridPagos.DisplayLayout.Override.HeaderAppearance = appearance42;
			appearance43.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance43.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.RowAlternateAppearance = appearance43;
			appearance44.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance44.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance44.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance45;
			this.gridPagos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.gridPagos.Location = new System.Drawing.Point(8, 296);
			this.gridPagos.Name = "gridPagos";
			this.gridPagos.Size = new System.Drawing.Size(1040, 331);
			this.gridPagos.TabIndex = 169;
			this.gridPagos.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.gridPagos_InitializeRow);
			this.gridPagos.DoubleClick += new System.EventHandler(this.gridPagos_DoubleClick);
			this.gridPagos.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.gridPagos_BeforeCellUpdate);
			this.gridPagos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.gridPagos_KeyDown);
			this.gridPagos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.gridPagos_AfterCellUpdate);
			this.gridPagos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.gridPagos_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(System.Double);
			ultraDataColumn22.DefaultValue = 0;
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Double);
			ultraDataColumn24.DefaultValue = 0;
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn26.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn27.DefaultValue = 0;
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn28.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn29.DataType = typeof(System.Double);
			ultraDataColumn29.DefaultValue = 0;
			ultraDataColumn30.DataType = typeof(System.Double);
			ultraDataColumn30.DefaultValue = 0;
			ultraDataColumn31.DataType = typeof(System.Double);
			ultraDataColumn31.DefaultValue = 0;
			ultraDataColumn32.DataType = typeof(System.Double);
			ultraDataColumn32.DefaultValue = 0;
			ultraDataColumn33.DataType = typeof(System.Double);
			ultraDataColumn33.DefaultValue = 0;
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.Double);
			ultraDataColumn36.DefaultValue = 0;
			ultraDataColumn37.DataType = typeof(System.DateTime);
			ultraDataColumn37.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn43.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn44.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45});
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel6.Location = new System.Drawing.Point(608, 268);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(100, 15);
			this.ultraLabel6.TabIndex = 173;
			this.ultraLabel6.Text = "SIGUIENTE CUOTA";
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel7.Location = new System.Drawing.Point(608, 244);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(76, 15);
			this.ultraLabel7.TabIndex = 172;
			this.ultraLabel7.Text = "SALDO TOTAL";
			// 
			// ultraLabel8
			// 
			this.ultraLabel8.AutoSize = true;
			this.ultraLabel8.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel8.Location = new System.Drawing.Point(816, 220);
			this.ultraLabel8.Name = "ultraLabel8";
			this.ultraLabel8.Size = new System.Drawing.Size(102, 15);
			this.ultraLabel8.TabIndex = 171;
			this.ultraLabel8.Text = "CUOTAS VENCIDAS";
			// 
			// ultraLabel9
			// 
			this.ultraLabel9.AutoSize = true;
			this.ultraLabel9.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel9.Location = new System.Drawing.Point(608, 220);
			this.ultraLabel9.Name = "ultraLabel9";
			this.ultraLabel9.Size = new System.Drawing.Size(90, 15);
			this.ultraLabel9.TabIndex = 170;
			this.ultraLabel9.Text = "SALDO VENCIDO";
			// 
			// ultraLabel10
			// 
			this.ultraLabel10.AutoSize = true;
			this.ultraLabel10.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel10.Location = new System.Drawing.Point(816, 268);
			this.ultraLabel10.Name = "ultraLabel10";
			this.ultraLabel10.Size = new System.Drawing.Size(123, 15);
			this.ultraLabel10.TabIndex = 174;
			this.ultraLabel10.Text = "PRÓXIMA A VENCER(D)";
			// 
			// txtSaldoTotal
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoTotal.Appearance = appearance46;
			this.txtSaldoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoTotal.Enabled = false;
			this.txtSaldoTotal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoTotal.FormatString = "#,##0.00";
			this.txtSaldoTotal.Location = new System.Drawing.Point(720, 240);
			this.txtSaldoTotal.MinValue = 0;
			this.txtSaldoTotal.Name = "txtSaldoTotal";
			this.txtSaldoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoTotal.PromptChar = ' ';
			this.txtSaldoTotal.Size = new System.Drawing.Size(80, 22);
			this.txtSaldoTotal.TabIndex = 177;
			// 
			// txtCuotaVencida
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaVencida.Appearance = appearance47;
			this.txtCuotaVencida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaVencida.Enabled = false;
			this.txtCuotaVencida.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCuotaVencida.FormatString = "";
			this.txtCuotaVencida.Location = new System.Drawing.Point(952, 216);
			this.txtCuotaVencida.MinValue = 0;
			this.txtCuotaVencida.Name = "txtCuotaVencida";
			this.txtCuotaVencida.PromptChar = ' ';
			this.txtCuotaVencida.Size = new System.Drawing.Size(64, 22);
			this.txtCuotaVencida.TabIndex = 176;
			// 
			// txtSaldoVencido
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoVencido.Appearance = appearance48;
			this.txtSaldoVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoVencido.Enabled = false;
			this.txtSaldoVencido.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoVencido.FormatString = "#,##0.00";
			this.txtSaldoVencido.Location = new System.Drawing.Point(720, 216);
			this.txtSaldoVencido.MinValue = -2147483647;
			this.txtSaldoVencido.Name = "txtSaldoVencido";
			this.txtSaldoVencido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoVencido.PromptChar = ' ';
			this.txtSaldoVencido.Size = new System.Drawing.Size(80, 22);
			this.txtSaldoVencido.TabIndex = 175;
			// 
			// txtProximaVencer
			// 
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProximaVencer.Appearance = appearance49;
			this.txtProximaVencer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProximaVencer.Enabled = false;
			this.txtProximaVencer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtProximaVencer.FormatString = "";
			this.txtProximaVencer.Location = new System.Drawing.Point(952, 264);
			this.txtProximaVencer.MinValue = 0;
			this.txtProximaVencer.Name = "txtProximaVencer";
			this.txtProximaVencer.PromptChar = ' ';
			this.txtProximaVencer.Size = new System.Drawing.Size(64, 22);
			this.txtProximaVencer.TabIndex = 179;
			// 
			// txtSiguienteCuota
			// 
			appearance50.ForeColor = System.Drawing.Color.Black;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSiguienteCuota.Appearance = appearance50;
			this.txtSiguienteCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSiguienteCuota.Enabled = false;
			this.txtSiguienteCuota.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSiguienteCuota.FormatString = "#,##0.00";
			this.txtSiguienteCuota.Location = new System.Drawing.Point(720, 264);
			this.txtSiguienteCuota.MinValue = 0;
			this.txtSiguienteCuota.Name = "txtSiguienteCuota";
			this.txtSiguienteCuota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSiguienteCuota.PromptChar = ' ';
			this.txtSiguienteCuota.Size = new System.Drawing.Size(80, 22);
			this.txtSiguienteCuota.TabIndex = 178;
			// 
			// txtCobro
			// 
			appearance51.ForeColor = System.Drawing.Color.Red;
			appearance51.ForeColorDisabled = System.Drawing.Color.Red;
			this.txtCobro.Appearance = appearance51;
			this.txtCobro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCobro.Enabled = false;
			this.txtCobro.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCobro.FormatString = "#,##0.00";
			this.txtCobro.Location = new System.Drawing.Point(808, 8);
			this.txtCobro.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtCobro.MinValue = 0;
			this.txtCobro.Name = "txtCobro";
			this.txtCobro.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCobro.PromptChar = ' ';
			this.txtCobro.Size = new System.Drawing.Size(100, 22);
			this.txtCobro.TabIndex = 4;
			this.txtCobro.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCobro_KeyDown);
			this.txtCobro.Click += new System.EventHandler(this.txtCobro_Click);
			this.txtCobro.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCobro_KeyPress);
			this.txtCobro.ValueChanged += new System.EventHandler(this.txtCobro_ValueChanged);
			this.txtCobro.Enter += new System.EventHandler(this.txtCobro_Enter);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(8, 84);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(38, 15);
			this.ultraLabel2.TabIndex = 182;
			this.ultraLabel2.Text = "Cliente";
			// 
			// txtNombre
			// 
			appearance52.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance52.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance52;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(64, 80);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(300, 22);
			this.txtNombre.TabIndex = 3;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(308, -60);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.ReadOnly = true;
			this.txtIdCliente.Size = new System.Drawing.Size(32, 22);
			this.txtIdCliente.TabIndex = 241;
			this.txtIdCliente.Visible = false;
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
			// dtFecha
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance53;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.dtFecha.CausesValidation = false;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 0;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 35);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 243;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance54;
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 180;
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 22);
			this.cmbBodega.TabIndex = 1;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblFecha
			// 
			this.lblFecha.AutoSize = true;
			this.lblFecha.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblFecha.Location = new System.Drawing.Point(8, 11);
			this.lblFecha.Name = "lblFecha";
			this.lblFecha.Size = new System.Drawing.Size(34, 15);
			this.lblFecha.TabIndex = 246;
			this.lblFecha.Text = "Fecha";
			// 
			// ultraLabel12
			// 
			this.ultraLabel12.AutoSize = true;
			this.ultraLabel12.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel12.Location = new System.Drawing.Point(584, 180);
			this.ultraLabel12.Name = "ultraLabel12";
			this.ultraLabel12.Size = new System.Drawing.Size(111, 15);
			this.ultraLabel12.TabIndex = 247;
			this.ultraLabel12.Text = "VALOR AUTORIZADO";
			// 
			// txtValorAutorizado
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorAutorizado.Appearance = appearance55;
			this.txtValorAutorizado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorAutorizado.Enabled = false;
			this.txtValorAutorizado.FormatString = "#,##0.00";
			this.txtValorAutorizado.Location = new System.Drawing.Point(720, 176);
			this.txtValorAutorizado.MinValue = -2147483647;
			this.txtValorAutorizado.Name = "txtValorAutorizado";
			this.txtValorAutorizado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorAutorizado.PromptChar = ' ';
			this.txtValorAutorizado.Size = new System.Drawing.Size(80, 22);
			this.txtValorAutorizado.TabIndex = 6;
			this.txtValorAutorizado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValorAutorizado_KeyDown);
			this.txtValorAutorizado.Click += new System.EventHandler(this.txtValorAutorizado_Click);
			this.txtValorAutorizado.Validated += new System.EventHandler(this.txtValorAutorizado_Validated);
			this.txtValorAutorizado.Enter += new System.EventHandler(this.txtValorAutorizado_Enter);
			// 
			// txtIdAutorizacion
			// 
			this.txtIdAutorizacion.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.txtIdAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAutorizacion.Enabled = false;
			this.txtIdAutorizacion.Location = new System.Drawing.Point(388, -60);
			this.txtIdAutorizacion.Name = "txtIdAutorizacion";
			this.txtIdAutorizacion.PromptChar = ' ';
			this.txtIdAutorizacion.ReadOnly = true;
			this.txtIdAutorizacion.Size = new System.Drawing.Size(32, 22);
			this.txtIdAutorizacion.TabIndex = 250;
			this.txtIdAutorizacion.Visible = false;
			// 
			// txtEstadoAutorizacion
			// 
			this.txtEstadoAutorizacion.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.txtEstadoAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoAutorizacion.Enabled = false;
			this.txtEstadoAutorizacion.Location = new System.Drawing.Point(348, -60);
			this.txtEstadoAutorizacion.Name = "txtEstadoAutorizacion";
			this.txtEstadoAutorizacion.PromptChar = ' ';
			this.txtEstadoAutorizacion.ReadOnly = true;
			this.txtEstadoAutorizacion.Size = new System.Drawing.Size(32, 22);
			this.txtEstadoAutorizacion.TabIndex = 251;
			this.txtEstadoAutorizacion.Visible = false;
			// 
			// btCobro
			// 
			this.btCobro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btCobro.Enabled = false;
			this.btCobro.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btCobro.ImageIndex = 2;
			this.btCobro.ImageList = this.imageList1;
			this.btCobro.Location = new System.Drawing.Point(961, 7);
			this.btCobro.Name = "btCobro";
			this.btCobro.Size = new System.Drawing.Size(88, 23);
			this.btCobro.TabIndex = 5;
			this.btCobro.Text = "&Cobro";
			this.btCobro.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btCobro.Click += new System.EventHandler(this.btCobro_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAutorizar.Enabled = false;
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.ImageIndex = 3;
			this.btnAutorizar.ImageList = this.imageList1;
			this.btnAutorizar.Location = new System.Drawing.Point(865, 176);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.Size = new System.Drawing.Size(88, 23);
			this.btnAutorizar.TabIndex = 263;
			this.btnAutorizar.Text = "A&utorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 0;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(961, 176);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 262;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnActualizarEDsctoGM
			// 
			this.btnActualizarEDsctoGM.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizarEDsctoGM.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizarEDsctoGM.Location = new System.Drawing.Point(961, 128);
			this.btnActualizarEDsctoGM.Name = "btnActualizarEDsctoGM";
			this.btnActualizarEDsctoGM.Size = new System.Drawing.Size(88, 23);
			this.btnActualizarEDsctoGM.TabIndex = 265;
			this.btnActualizarEDsctoGM.Text = "&...";
			this.btnActualizarEDsctoGM.Click += new System.EventHandler(this.btnActualizarEDsctoGM_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.ImageIndex = 1;
			this.btnCancelar.ImageList = this.imageList1;
			this.btnCancelar.Location = new System.Drawing.Point(961, 152);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 266;
			this.btnCancelar.Text = "C&ancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(672, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(110, 17);
			this.label1.TabIndex = 267;
			this.label1.Text = "MONTO A COBRAR";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnAnticipo
			// 
			this.btnAnticipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAnticipo.Enabled = false;
			this.btnAnticipo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnticipo.Location = new System.Drawing.Point(961, 56);
			this.btnAnticipo.Name = "btnAnticipo";
			this.btnAnticipo.Size = new System.Drawing.Size(88, 23);
			this.btnAnticipo.TabIndex = 268;
			this.btnAnticipo.Text = "&Anticipos/NDC";
			this.btnAnticipo.Click += new System.EventHandler(this.btnAnticipo_Click);
			// 
			// btnReliquidar
			// 
			this.btnReliquidar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnReliquidar.Enabled = false;
			this.btnReliquidar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnReliquidar.ImageList = this.imageList1;
			this.btnReliquidar.Location = new System.Drawing.Point(961, 32);
			this.btnReliquidar.Name = "btnReliquidar";
			this.btnReliquidar.Size = new System.Drawing.Size(88, 23);
			this.btnReliquidar.TabIndex = 269;
			this.btnReliquidar.Text = "&Reliquidar";
			this.btnReliquidar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnReliquidar.Click += new System.EventHandler(this.btnReliquidar_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 14F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Red;
			this.lblEstado.Location = new System.Drawing.Point(528, 80);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 270;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdentificacion
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdentificacion.Appearance = appearance56;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Center;
			editorButton1.Appearance = appearance57;
			editorButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			editorButton1.Text = "...";
			this.txtIdentificacion.ButtonsRight.Add(editorButton1);
			this.txtIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacion.Location = new System.Drawing.Point(64, 56);
			this.txtIdentificacion.MaxLength = 10;
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(150, 22);
			this.txtIdentificacion.TabIndex = 2;
			this.txtIdentificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIdentificacion_KeyDown);
			this.txtIdentificacion.EditorButtonClick += new Infragistics.Win.UltraWinEditors.EditorButtonEventHandler(this.txtIdentificacion_EditorButtonClick);
			this.txtIdentificacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtIdentificacion_KeyPress);
			this.txtIdentificacion.Validated += new System.EventHandler(this.txtIdentificacion_Validated);
			this.txtIdentificacion.ValueChanged += new System.EventHandler(this.txtIdentificacion_ValueChanged);
			// 
			// txtDiasVencidos
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasVencidos.Appearance = appearance58;
			this.txtDiasVencidos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasVencidos.Enabled = false;
			this.txtDiasVencidos.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDiasVencidos.FormatString = "";
			this.txtDiasVencidos.Location = new System.Drawing.Point(952, 240);
			this.txtDiasVencidos.MinValue = -2147483647;
			this.txtDiasVencidos.Name = "txtDiasVencidos";
			this.txtDiasVencidos.PromptChar = ' ';
			this.txtDiasVencidos.Size = new System.Drawing.Size(64, 22);
			this.txtDiasVencidos.TabIndex = 272;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel1.Location = new System.Drawing.Point(816, 244);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(86, 15);
			this.ultraLabel1.TabIndex = 271;
			this.ultraLabel1.Text = "DIAS VENCIDOS";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 112);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1056, 8);
			this.groupBox1.TabIndex = 273;
			this.groupBox1.TabStop = false;
			// 
			// txtSaldoAFavor
			// 
			appearance59.ForeColor = System.Drawing.Color.Black;
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoAFavor.Appearance = appearance59;
			this.txtSaldoAFavor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoAFavor.Enabled = false;
			this.txtSaldoAFavor.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoAFavor.FormatString = "#,##0.00";
			this.txtSaldoAFavor.Location = new System.Drawing.Point(808, 56);
			this.txtSaldoAFavor.MinValue = -2147483647;
			this.txtSaldoAFavor.Name = "txtSaldoAFavor";
			this.txtSaldoAFavor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoAFavor.PromptChar = ' ';
			this.txtSaldoAFavor.Size = new System.Drawing.Size(100, 22);
			this.txtSaldoAFavor.TabIndex = 275;
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel3.Location = new System.Drawing.Point(672, 60);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(70, 15);
			this.ultraLabel3.TabIndex = 274;
			this.ultraLabel3.Text = "POR CRUZAR";
			// 
			// txtProntoPago
			// 
			appearance60.ForeColor = System.Drawing.Color.Black;
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProntoPago.Appearance = appearance60;
			this.txtProntoPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProntoPago.Enabled = false;
			this.txtProntoPago.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtProntoPago.FormatString = "#,##0.00";
			this.txtProntoPago.Location = new System.Drawing.Point(808, 32);
			this.txtProntoPago.MinValue = -2147483647;
			this.txtProntoPago.Name = "txtProntoPago";
			this.txtProntoPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProntoPago.PromptChar = ' ';
			this.txtProntoPago.Size = new System.Drawing.Size(100, 22);
			this.txtProntoPago.TabIndex = 277;
			// 
			// ultraLabel5
			// 
			this.ultraLabel5.AutoSize = true;
			this.ultraLabel5.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel5.Location = new System.Drawing.Point(672, 36);
			this.ultraLabel5.Name = "ultraLabel5";
			this.ultraLabel5.Size = new System.Drawing.Size(126, 15);
			this.ultraLabel5.TabIndex = 276;
			this.ultraLabel5.Text = "VALOR RELIQUIDACIÓN";
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(568, 200);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(480, 8);
			this.groupBox2.TabIndex = 278;
			this.groupBox2.TabStop = false;
			// 
			// ultraLabel11
			// 
			appearance61.ForeColor = System.Drawing.Color.White;
			appearance61.ForeColorDisabled = System.Drawing.Color.White;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Center;
			this.ultraLabel11.Appearance = appearance61;
			this.ultraLabel11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(192)), ((System.Byte)(0)));
			this.ultraLabel11.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraLabel11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.ultraLabel11.Location = new System.Drawing.Point(536, 8);
			this.ultraLabel11.Name = "ultraLabel11";
			this.ultraLabel11.Size = new System.Drawing.Size(88, 15);
			this.ultraLabel11.TabIndex = 279;
			this.ultraLabel11.Text = "AL DÍA";
			// 
			// ultraLabel13
			// 
			appearance62.ForeColor = System.Drawing.Color.White;
			appearance62.ForeColorDisabled = System.Drawing.Color.White;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Center;
			this.ultraLabel13.Appearance = appearance62;
			this.ultraLabel13.BackColor = System.Drawing.Color.Red;
			this.ultraLabel13.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraLabel13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.ultraLabel13.Location = new System.Drawing.Point(536, 24);
			this.ultraLabel13.Name = "ultraLabel13";
			this.ultraLabel13.Size = new System.Drawing.Size(88, 15);
			this.ultraLabel13.TabIndex = 280;
			this.ultraLabel13.Text = "VENCIDO";
			// 
			// ultraLabel14
			// 
			appearance63.ForeColor = System.Drawing.Color.White;
			appearance63.ForeColorDisabled = System.Drawing.Color.White;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Center;
			this.ultraLabel14.Appearance = appearance63;
			this.ultraLabel14.BackColor = System.Drawing.Color.Blue;
			this.ultraLabel14.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraLabel14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.ultraLabel14.Location = new System.Drawing.Point(536, 40);
			this.ultraLabel14.Name = "ultraLabel14";
			this.ultraLabel14.Size = new System.Drawing.Size(88, 15);
			this.ultraLabel14.TabIndex = 281;
			this.ultraLabel14.Text = "CANCELADO";
			// 
			// btnHistorico
			// 
			this.btnHistorico.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnHistorico.Enabled = false;
			this.btnHistorico.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnHistorico.ImageList = this.imageList1;
			this.btnHistorico.Location = new System.Drawing.Point(961, 80);
			this.btnHistorico.Name = "btnHistorico";
			this.btnHistorico.Size = new System.Drawing.Size(88, 23);
			this.btnHistorico.TabIndex = 282;
			this.btnHistorico.Text = "&Historico";
			this.btnHistorico.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnHistorico.Click += new System.EventHandler(this.btnHistorico_Click);
			// 
			// btnRefinanciacion
			// 
			this.btnRefinanciacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRefinanciacion.Enabled = false;
			this.btnRefinanciacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRefinanciacion.ImageList = this.imageList1;
			this.btnRefinanciacion.Location = new System.Drawing.Point(865, 128);
			this.btnRefinanciacion.Name = "btnRefinanciacion";
			this.btnRefinanciacion.Size = new System.Drawing.Size(88, 23);
			this.btnRefinanciacion.TabIndex = 283;
			this.btnRefinanciacion.Text = "&Refinanciar";
			this.btnRefinanciacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRefinanciacion.Click += new System.EventHandler(this.btnRefinanciacion_Click);
			// 
			// mnuRefinanciacion
			// 
			this.mnuRefinanciacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																											this.mnuDocumentos,
																																											this.mnuGenerarCXC});
			// 
			// mnuDocumentos
			// 
			this.mnuDocumentos.Index = 0;
			this.mnuDocumentos.Text = "Generar";
			this.mnuDocumentos.Click += new System.EventHandler(this.mnuDocumentos_Click);
			// 
			// mnuGenerarCXC
			// 
			this.mnuGenerarCXC.Index = 1;
			this.mnuGenerarCXC.Text = "Procesar";
			this.mnuGenerarCXC.Click += new System.EventHandler(this.mnuGenerarCXC_Click);
			// 
			// btnCondonar
			// 
			this.btnCondonar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCondonar.Enabled = false;
			this.btnCondonar.Image = ((System.Drawing.Image)(resources.GetObject("btnCondonar.Image")));
			this.btnCondonar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCondonar.Location = new System.Drawing.Point(865, 152);
			this.btnCondonar.Name = "btnCondonar";
			this.btnCondonar.Size = new System.Drawing.Size(88, 23);
			this.btnCondonar.TabIndex = 286;
			this.btnCondonar.Text = "C&ondonar";
			this.btnCondonar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCondonar.Click += new System.EventHandler(this.btnCondonar_Click);
			// 
			// txtValorACondonar
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorACondonar.Appearance = appearance64;
			this.txtValorACondonar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorACondonar.Enabled = false;
			this.txtValorACondonar.FormatString = "#,##0.00";
			this.txtValorACondonar.Location = new System.Drawing.Point(720, 152);
			this.txtValorACondonar.MinValue = -2147483647;
			this.txtValorACondonar.Name = "txtValorACondonar";
			this.txtValorACondonar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorACondonar.PromptChar = ' ';
			this.txtValorACondonar.Size = new System.Drawing.Size(80, 22);
			this.txtValorACondonar.TabIndex = 284;
			// 
			// ultraLabel15
			// 
			this.ultraLabel15.AutoSize = true;
			this.ultraLabel15.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel15.Location = new System.Drawing.Point(584, 156);
			this.ultraLabel15.Name = "ultraLabel15";
			this.ultraLabel15.Size = new System.Drawing.Size(112, 15);
			this.ultraLabel15.TabIndex = 285;
			this.ultraLabel15.Text = "VALOR A CONDONAR";
			// 
			// txtValorARefinanciar
			// 
			appearance65.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorARefinanciar.Appearance = appearance65;
			this.txtValorARefinanciar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorARefinanciar.Enabled = false;
			this.txtValorARefinanciar.FormatString = "#,##0.00";
			this.txtValorARefinanciar.Location = new System.Drawing.Point(720, 128);
			this.txtValorARefinanciar.MinValue = -2147483647;
			this.txtValorARefinanciar.Name = "txtValorARefinanciar";
			this.txtValorARefinanciar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorARefinanciar.PromptChar = ' ';
			this.txtValorARefinanciar.Size = new System.Drawing.Size(80, 22);
			this.txtValorARefinanciar.TabIndex = 287;
			// 
			// ultraLabel16
			// 
			this.ultraLabel16.AutoSize = true;
			this.ultraLabel16.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraLabel16.Location = new System.Drawing.Point(584, 132);
			this.ultraLabel16.Name = "ultraLabel16";
			this.ultraLabel16.Size = new System.Drawing.Size(123, 15);
			this.ultraLabel16.TabIndex = 288;
			this.ultraLabel16.Text = "VALOR A REFINANCIAR";
			// 
			// ultraLabel17
			// 
			appearance66.ForeColor = System.Drawing.Color.White;
			appearance66.ForeColorDisabled = System.Drawing.Color.White;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Center;
			this.ultraLabel17.Appearance = appearance66;
			this.ultraLabel17.BackColor = System.Drawing.Color.Brown;
			this.ultraLabel17.BorderStyleOuter = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraLabel17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.ultraLabel17.Location = new System.Drawing.Point(536, 56);
			this.ultraLabel17.Name = "ultraLabel17";
			this.ultraLabel17.Size = new System.Drawing.Size(88, 15);
			this.ultraLabel17.TabIndex = 289;
			this.ultraLabel17.Text = "DIFERIDO";
			this.ultraLabel17.Visible = false;
			// 
			// btnSNoDif
			// 
			this.btnSNoDif.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSNoDif.Enabled = false;
			this.btnSNoDif.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSNoDif.ImageList = this.imageList1;
			this.btnSNoDif.Location = new System.Drawing.Point(352, 8);
			this.btnSNoDif.Name = "btnSNoDif";
			this.btnSNoDif.Size = new System.Drawing.Size(152, 23);
			this.btnSNoDif.TabIndex = 290;
			this.btnSNoDif.Text = "&Solicitar NO DIFERIMIENTO";
			this.btnSNoDif.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSNoDif.Visible = false;
			this.btnSNoDif.Click += new System.EventHandler(this.btnSNoDif_Click);
			// 
			// btnAux
			// 
			this.btnAux.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAux.Enabled = false;
			this.btnAux.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAux.Location = new System.Drawing.Point(416, 48);
			this.btnAux.Name = "btnAux";
			this.btnAux.Size = new System.Drawing.Size(88, 23);
			this.btnAux.TabIndex = 291;
			this.btnAux.Text = "&...";
			this.btnAux.Visible = false;
			this.btnAux.Click += new System.EventHandler(this.btnAux_Click);
			// 
			// frmCobranzas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1057, 635);
			this.ControlBox = false;
			this.Controls.Add(this.btnAux);
			this.Controls.Add(this.btnSNoDif);
			this.Controls.Add(this.ultraLabel17);
			this.Controls.Add(this.txtValorARefinanciar);
			this.Controls.Add(this.ultraLabel16);
			this.Controls.Add(this.btnCondonar);
			this.Controls.Add(this.txtValorACondonar);
			this.Controls.Add(this.ultraLabel15);
			this.Controls.Add(this.btnRefinanciacion);
			this.Controls.Add(this.btnHistorico);
			this.Controls.Add(this.ultraLabel14);
			this.Controls.Add(this.ultraLabel13);
			this.Controls.Add(this.ultraLabel11);
			this.Controls.Add(this.txtProntoPago);
			this.Controls.Add(this.ultraLabel5);
			this.Controls.Add(this.txtSaldoAFavor);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtDiasVencidos);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtIdentificacion);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnReliquidar);
			this.Controls.Add(this.btnAnticipo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnActualizarEDsctoGM);
			this.Controls.Add(this.btCobro);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtEstadoAutorizacion);
			this.Controls.Add(this.txtIdAutorizacion);
			this.Controls.Add(this.txtValorAutorizado);
			this.Controls.Add(this.ultraLabel12);
			this.Controls.Add(this.lblFecha);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtProximaVencer);
			this.Controls.Add(this.txtSiguienteCuota);
			this.Controls.Add(this.txtSaldoTotal);
			this.Controls.Add(this.txtCuotaVencida);
			this.Controls.Add(this.txtSaldoVencido);
			this.Controls.Add(this.ultraLabel10);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.ultraLabel8);
			this.Controls.Add(this.ultraLabel9);
			this.Controls.Add(this.gridPagos);
			this.Controls.Add(this.GridFacturas);
			this.Controls.Add(this.ultraLabel4);
			this.Controls.Add(this.txtCobro);
			this.Controls.Add(this.groupBox2);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCobranzas";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cobranza";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCobranzas_KeyDown);
			this.Load += new System.EventHandler(this.frmCobranzas_Load);
			((System.ComponentModel.ISupportInitialize)(this.GridFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaVencida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProximaVencer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSiguienteCuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorAutorizado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAFavor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorACondonar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorARefinanciar)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iBodega = 0;
		decimal dInteres = 0.00m;
		DateTime dtUltimaFechaPago = DateTime.Today;
		DateTime Vence = DateTime.Today;
		int idanticipo2;
		decimal dSaldoSD = 0.00m;
		int EstadoSorteo = 0;
		int EstadoCelular = 0;
		bool bFiniquitos = false;

		private void Seguridades()
		{			
			miAcceso = new Acceso(cdsSeteoF, "0801");
			
			if (!Funcion.Permiso("133", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.CambiarFecha) 
			{						
				this.dtFecha.Enabled = true;
				if (miAcceso.B15DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-5);
				else if (miAcceso.B30DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-30);
				else if (miAcceso.B60DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-500);
				else if (miAcceso.BCobranzasSinLimiteFechas) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddMonths(-12); //DateTime.Parse("01/01/2022");				
				else this.dtFecha.CalendarInfo.MinDate = DateTime.Today;
			}

			if (miAcceso.BCobrar)
			{
				this.txtCobro.Enabled = true;
				this.btCobro.Enabled = true;
			}			

			if (miAcceso.BReliquidar || miAcceso.BReliquidarCruzandoCuentas) this.btnReliquidar.Enabled = true;

			if (miAcceso.BCruzaAnticiposCP) 
			{
				this.btnAnticipo.Enabled = true;
				this.btnAux.Enabled = true;
			}

			if (miAcceso.BHistorico) this.btnHistorico.Enabled = true;

			if (miAcceso.BAutCobro)
			{
				this.txtValorAutorizado.Enabled = true;
				this.btnAutorizar.Enabled = true;								
			}

			#region Modificar Gastos
			if (!miAcceso.BModificarGastos) 
			{
				this.gridPagos.DisplayLayout.Bands[0].Columns["Gastos"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.gridPagos.DisplayLayout.Bands[0].Columns["DsctoGastos"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
			#endregion Modificar Gastos

			#region Modificar Mora
			if (!miAcceso.BModificarMora) 
			{
				this.gridPagos.DisplayLayout.Bands[0].Columns["Mora"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.gridPagos.DisplayLayout.Bands[0].Columns["DsctoIntMora"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}
			#endregion Modificar Mora

			#region Capital - Interes
			if (!miAcceso.BVerCapitalInteres)
			{
				this.gridPagos.DisplayLayout.Bands[0].Columns["CapitalTotal"].Hidden = true;
				this.gridPagos.DisplayLayout.Bands[0].Columns["InteresTotal"].Hidden = true;				
			}		

			if (miAcceso.BVerCapitalInteres && miAcceso.BModificarCapitalInteres)
			{
				this.gridPagos.DisplayLayout.Bands[0].Columns["CapitalTotal"].CellActivation = Activation.AllowEdit;
				this.gridPagos.DisplayLayout.Bands[0].Columns["InteresTotal"].CellActivation = Activation.AllowEdit;
			}			
			#endregion Capital - Interes
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCobranzas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void frmCobranzas_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);
						
			Seguridades();
			
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0) this.cmbBodega.Value = iBodega; 
			else 
			{
				this.cmbBodega.Enabled = true;
				this.cmbBodega.Value = 57;
			}
			if(bPantalla)
			{
				this.txtIdentificacion.Text = sCedula;
				this.txtIdentificacion_Validated ( sender, e);
			}

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void Cobranza()
		{
			if (this.GridFacturas.ActiveRow != null)
			{		
				try
				{
					LimpiaControles(false, true);
											
					#region Valores Estado de Cuenta
					string sSQL = string.Format("Exec ValoresCobranza {0}, '{1}'", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					dr.Read();
					if(dr.HasRows)
					{
						//this.txtSaldoVencido.Value = dr.GetDecimal(0);
						this.txtCuotaVencida.Value = dr.GetInt32(1);
						this.txtSiguienteCuota.Value = dr.GetDecimal(3);
						this.txtProximaVencer.Value = dr.GetInt32(4);
						iDiasVencidos = (int)dr.GetInt32(5); 
						this.txtSaldoAFavor.Value = dr.GetDecimal(6);
					}
					dr.Close();
				
					if ((int)this.txtCuotaVencida.Value > 0)
					{
						this.txtSaldoVencido.Appearance.ForeColor = Color.Red;
						this.txtSaldoVencido.Appearance.ForeColorDisabled = Color.Red;
						this.txtCuotaVencida.Appearance.ForeColor = Color.Red;
						this.txtCuotaVencida.Appearance.ForeColorDisabled = Color.Red;
						this.txtDiasVencidos.Appearance.ForeColor = Color.Red;
						this.txtDiasVencidos.Appearance.ForeColorDisabled = Color.Red;
					}
					#endregion Valores Estado de Cuenta
					
					if (iDiasVencidos >= 0) this.txtDiasVencidos.Value = iDiasVencidos; 
					else this.txtDiasVencidos.Value = 0;

					string sSQLUFP = string.Format("Exec CobranzasUltimaFechaPago {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					dtUltimaFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, sSQLUFP);

					//					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCovidAlDia {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value), true) > 0)
					//						bDescCovid = true;
					//					else					
					//						bDescCovid = false;
					
					ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);				

					#region Valor re liquidacion
					decimal dGC = 0.00m;
					decimal dIM = 0.00m;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPagos in this.gridPagos.Rows.All)
					{
						dGC = dGC + Convert.ToDecimal(drPagos.Cells["Gastos"].Value);
						dIM = dIM + Convert.ToDecimal(drPagos.Cells["Mora"].Value);
					}

					sSQL = string.Format("Exec ConsultaValorReliquidacion {0}, '{1}', {2}, {3}, 1", 
						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), dGC, dIM);
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
					
						if (Convert.ToDateTime(this.dtFecha.Value) > Vence)
							this.txtSaldoVencido.Value = 
								Math.Round(Convert.ToDecimal(this.txtSaldoVencido.Value), 2) + Convert.ToDecimal(dRow.Cells["Saldo"].Value);							
						
						if (Convert.ToDateTime(this.dtFecha.Value) == Vence)
							this.txtSaldoVencido.Value = 
								Math.Round(Convert.ToDecimal(this.txtSaldoVencido.Value), 2) + Convert.ToDecimal(dRow.Cells["Saldo"].Value) 
								- Convert.ToDecimal(dRow.Cells["Diferimiento"].Value);						
					}
					#endregion Valida Saldo Total
									
					#region Valida Refinanciacion
					if (miAcceso.BRefinanciar)
					{
						if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value)
						{
							if ((int)this.txtCuotaVencida.Value >= 3) 
							{
								if ((int)this.GridFacturas.ActiveRow.Cells["ER"].Value == 0)
								{
									this.txtValorARefinanciar.Value = Convert.ToDecimal(this.txtSaldoVencido.Value);
									this.btnRefinanciacion.Enabled = true;
									this.mnuDocumentos.Enabled = true;
								}
					
								if ((int)this.GridFacturas.ActiveRow.Cells["ER"].Value == 1)
								{								
									this.btnRefinanciacion.Enabled = true;
									this.mnuGenerarCXC.Enabled = true;
								}
							}
						}
					}
					#endregion Valida Refinanciacion

					#region Condonacion
					//					if (miAcceso.BCondonar)
					//					{
					//						if ((bool)this.GridFacturas.ActiveRow.Cells["Asignado"].Value)
					//						{
					//							if (Convert.ToDecimal(this.txtSaldoVencido.Value) > 0.00m)
					//							{
					//								this.txtValorACondonar.Value = Convert.ToDecimal(this.txtSaldoVencido.Value);
					//								this.btnCondonar.Enabled = true;
					//							}
					//						}
					//					}
					#endregion Condonacion

					#region Valida Autorizacion de Cobro
//					SqlDataReader drAut = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select idAutorizacionCobro, ValorAutorizado, Estado From HistorialAutorizacionesCobro Where idCompra = {0} And Estado = 0", 
//						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value), true);
//					drAut.Read();
//					if (drAut.HasRows)
//					{					
//						this.txtIdAutorizacion.Value = drAut.GetInt32(0);
//						this.txtCobro.Value = drAut.GetDecimal(1);
//						this.txtValorAutorizado.Value = drAut.GetDecimal(1);
//						this.txtEstadoAutorizacion.Value = drAut.GetInt32(2);
//
//						drAut.Close();
//
//						bMayorCobro = true;
//						dValorCobro = Convert.ToDecimal(this.txtCobro.Value);
//
//						if (miAcceso.BCobrar)
//						{
//							this.txtCobro.Enabled = true;
//							this.btCobro.Enabled = true;
//						}
//
//						return;
//					}				
//					drAut.Close();
					#endregion Valida Autorizacion de Cobro

					#region Valor a cobrar
					if (Convert.ToDecimal(this.txtSaldoVencido.Value) > 0.00m)
					{
						this.txtCobro.Value = Convert.ToDecimal(this.txtSaldoVencido.Value);
						bMayorCobro = true;
						dValorCobro = Convert.ToDecimal(this.txtCobro.Value);						
					}
					
					if (miAcceso.BCobrar)
					{
						this.txtCobro.Enabled = true;
						this.btCobro.Enabled = true;
					}
					#endregion Valor a cobrar
				}
				catch (Exception exc)
				{
					MessageBox.Show(string.Format("Exception Type: {0} {1}", exc.GetType(), exc.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
			}
		}

		private void gridPagos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico

			#region Decimales para el Saldo Abono Gastos Mora
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decPrecio"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decPrecio"]);

			e.Layout.Bands[0].Columns["CapitalTotal"].Format = stFormato;
			e.Layout.Bands[0].Columns["CapitalTotal"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["InteresTotal"].Format = stFormato;
			e.Layout.Bands[0].Columns["InteresTotal"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["Gastos"].Format = stFormato;
			e.Layout.Bands[0].Columns["Gastos"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["DsctoGastos"].Format = stFormato;
			e.Layout.Bands[0].Columns["DsctoGastos"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["Mora"].Format = stFormato;
			e.Layout.Bands[0].Columns["Mora"].MaskInput = stInput;
			
			e.Layout.Bands[0].Columns["DsctoIntMora"].Format = stFormato;
			e.Layout.Bands[0].Columns["DsctoIntMora"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["Mora"].Format = stFormato;
			e.Layout.Bands[0].Columns["Mora"].MaskInput = stInput;
			
			e.Layout.Bands[0].Columns["Saldo"].Format = stFormato;
			e.Layout.Bands[0].Columns["Saldo"].MaskInput = stInput;

			e.Layout.Bands[0].Columns["Abono"].Format = stFormato;
			e.Layout.Bands[0].Columns["Abono"].MaskInput = stInput;

			this.txtCobro.FormatString = stFormato;
			#endregion Decimales para el Saldo Abono Gastos Mora	
		}

		private void gridPagos_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["Estado"].Value.ToString().Trim() == "0") 
			{
				e.Row.Appearance.ForeColor = Color.Green; //AL DIA
				e.Row.Appearance.ForeColorDisabled = Color.Green;
			}
			if (e.Row.Cells["Estado"].Value.ToString().Trim() == "1") 
			{
				e.Row.Appearance.ForeColor = Color.Red; // EN MORA
				e.Row.Appearance.ForeColorDisabled = Color.Red;
			}
			if (e.Row.Cells["Estado"].Value.ToString().Trim() == "2")
			{
				e.Row.Appearance.ForeColor = Color.Blue; // PAGO
				e.Row.Appearance.ForeColorDisabled = Color.Blue;
			}
			if (e.Row.Cells["Estado"].Value.ToString().Trim() == "3")
			{
				e.Row.Appearance.ForeColor = Color.Fuchsia; // ABONO
				e.Row.Appearance.ForeColorDisabled = Color.Fuchsia;
			}
			if (e.Row.Cells["Estado"].Value.ToString().Trim() == "4")
			{
				e.Row.Appearance.ForeColor = Color.Brown; // DIFERIMIENTO
				e.Row.Appearance.ForeColorDisabled = Color.Brown;
			}
		}
		
		private void txtCobro_Click(object sender, System.EventArgs e)
		{
			this.txtCobro.SelectAll();
		}

		private void txtCobro_Enter(object sender, System.EventArgs e)
		{
			this.txtCobro.SelectAll();
		}
					
		bool bActualiza = false;

		private void gridPagos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Variables
			decimal dGasto = 0.00m;
			decimal dMora = 0.00m;
			int iGastoMora = 0;
			int iCompra = 0;
			int iCreAmortizacion = 0;
			decimal OrigenValor = 0.00m;
			decimal CambioValor = 0.00m;
			decimal dValorRecargo = 0.00m;
			int iCuota = 0;
			int iGestion = 0;
			int iDiasMora = 0;
			#endregion Variables

			if (bActualiza) return;

			#region Capital
			decimal dSaldo = 0.00m;
			if (e.Cell.Column.ToString() == "CapitalTotal")
			{
				decimal dCapital = Convert.ToDecimal(e.Cell.Row.Cells["CapitalTotal"].Value);				
				decimal dABCapital = Convert.ToDecimal(e.Cell.Row.Cells["ABCapital"].Value);
				
				dSaldo = Convert.ToDecimal(e.Cell.Row.Cells["ValorCuota"].Value) - (Convert.ToDecimal(e.Cell.Row.Cells["CapitalTotal"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ABCapital"].Value))
					+ (Convert.ToDecimal(e.Cell.Row.Cells["InteresTotal"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ABInteres"].Value));

				string sSQLG = string.Format("Update Cre_TablaDeAmortizacion Set Capital = {0}, Saldo = {2} Where idCre_TablaDeAmortizacion = {1}", dCapital, (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value, dSaldo);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLG, true);			
			
				return;
			}
			#endregion Capital

			#region Interes
			if (e.Cell.Column.ToString() == "InteresTotal")
			{
				decimal dInteres = Convert.ToDecimal(e.Cell.Row.Cells["InteresTotal"].Value);				
				decimal dABInteres = Convert.ToDecimal(e.Cell.Row.Cells["ABInteres"].Value);				
				
				dSaldo = Convert.ToDecimal(e.Cell.Row.Cells["ValorCuota"].Value) - (Convert.ToDecimal(e.Cell.Row.Cells["CapitalTotal"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ABCapital"].Value))
					+ (Convert.ToDecimal(e.Cell.Row.Cells["InteresTotal"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["ABInteres"].Value));

				string sSQLG = string.Format("Update Cre_TablaDeAmortizacion Set Interes = {0}, Saldo = {2} Where idCre_TablaDeAmortizacion = {1}", dInteres, (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value, dSaldo);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLG, true);
			
				return;
			}
			#endregion Interes
						
			#region valida moficiacion gastos mora
			//			if (e.Cell.Column.ToString() == "Mora")
			//			{	
			//				bActualiza = true;
			//		
			//				if (!bDescCovid    ) // miAcceso.GrabaCero
			//				{		
			//					dsctoMora = (Decimal)e.Cell.OriginalValue ;
			//					dsctoMora	 = dsctoMora - (Decimal)e.Cell.Row.Cells["Mora"].Value ;
			//					if (dsctoMora == 0) dsctoMora = (Decimal)e.Cell.Row.Cells["Mora"].Value ; 
			//					if (dsctoMora < 0) dsctoMora = 0.00m;											
			//				}			
			//				else
			//				{
			//					e.Cell.Row.Cells["Mora"].Value = (Decimal)e.Cell.OriginalValue;					
			//				}
			//
			//				bActualiza = false;
			//			}	
			//  
			//			if (e.Cell.Column.ToString() == "Gastos")
			//			{	
			//				bActualiza = true;
			//		
			//				if (!bDescCovid  || miAcceso.GrabaCero)
			//				{		
			//					DescGat = (Decimal)e.Cell.OriginalValue ;
			//					DescGat =   DescGat - (Decimal)e.Cell.Row.Cells["Gastos"].Value ;
			//					if (DescGat == 0) DescGat = (Decimal)e.Cell.Row.Cells["Gastos"].Value ; 
			//					if (DescGat < 0) DescGat =  0.00m;											
			//				}	
			//				else
			//				{
			//					e.Cell.Row.Cells["Gastos"].Value = (Decimal)e.Cell.OriginalValue;
			//				}
			//				bActualiza = false;
			//			}	
			#endregion valida modificacion gastos mora

			#region Gastos
			if (e.Cell.Column.ToString() == "Gastos")
			{
				//				if (!bDescCovid || miAcceso.GrabaCero )
				//				{
				#region Modifica Valores Locales
				bActualiza = true;
									
				dGasto = Convert.ToDecimal(e.Cell.Row.Cells["Gastos"].Value);
				iGastoMora = 1;
				iCompra = (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value;
				iCreAmortizacion = (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value;
				CambioValor = (decimal)e.Cell.Row.Cells["Gastos"].Value;
				OrigenValor = (decimal)e.Cell.OriginalValue;
				iCuota = (int)e.Cell.Row.Cells["NumeroCuota"].Value;
				iDiasMora = (int)e.Cell.Row.Cells["Retraso"].Value;
				
				if (iDiasMora <= 0)
				{
					MessageBox.Show("No puede modificar valores de Gastos si fecha de Vencimiento es menor a la fecha actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cell.Row.Cells["Gastos"].Value = (decimal)e.Cell.Row.Cells["Gastos"].OriginalValue;
					bActualiza = false;
					return;
				}

				if(!miAcceso.BGestion)
				{
					if (iDiasMora > 0 && iDiasMora < 61)
					{
						dValorRecargo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec Cre_RetornaValorGestionCobranza {0}", (decimal)e.Cell.Row.Cells["ValorCuota"].Value));

						dValorRecargo = dValorRecargo * 50.00m/100.00m;

						if (CambioValor < dValorRecargo)
						{
							MessageBox.Show("El descuento del gasto no puede ser superior al 50%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cell.Row.Cells["Gastos"].Value = (decimal)e.Cell.Row.Cells["Gastos"].OriginalValue;
							bActualiza = false;
							return;
						}
					}
				}
												
				string sSQLG = string.Format("Update Cre_TablaDeAmortizacion Set Gastos = {0}, DsctGasManual = 1 , DescGestion = {2} Where idCre_TablaDeAmortizacion = {1}", 
					dGasto, (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value, dGasto);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLG, true);
				
				string sSQLGuarda = string.Format("Exec Cre_ActualizaGestionMora {0}, {1}, {2}, {3}, {4}, {5}",
					iCompra, iCreAmortizacion, OrigenValor, CambioValor, iCuota, iGastoMora);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLGuarda, true);

				this.ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
				bActualiza = false;
				return;
				
				#endregion Modifica Valores Locales
			}
			#endregion Gastos			

			#region Mora
			if (e.Cell.Column.ToString() == "Mora")
			{		
				//				if (!bDescCovid )//|| miAcceso.GrabaCero
				//				{				
				bActualiza = true;
				dMora = Convert.ToDecimal(e.Cell.Row.Cells["Mora"].Value);
				iGastoMora = 2;
				iCompra = (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value;
				iCreAmortizacion = (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value;
				CambioValor = (decimal)e.Cell.Row.Cells["Mora"].Value;
				OrigenValor = (decimal)e.Cell.OriginalValue;
				iCuota = (int)e.Cell.Row.Cells["NumeroCuota"].Value;
				
				#region Modifica Valores Locales
				if(miAcceso.BMora)
				{
					string sSQLFe = string.Format("Exec ValidaFechaGestion {0}, {1}", iCompra, iCuota);
					iGestion = Funcion.iEscalarSQL(cdsSeteoF, sSQLFe);

					if(iGestion == 1)
					{
						string sSQLMo = string.Format("Update Cre_TablaDeAmortizacion Set Mora = {0}, DsctMorManual = 1, DesMora = {2} Where idCre_TablaDeAmortizacion = {1}", 
							dMora, (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value, dMora);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLMo, true);
				
						this.ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);

						string sSQLGuarda = string.Format("Exec Cre_ActualizaGestionMora {0}, {1}, {2}, {3}, {4}, {5}",
							iCompra, iCreAmortizacion, OrigenValor, CambioValor, iCuota, iGastoMora);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLGuarda, true);	

						bActualiza = false;
						return;
					}
					else
					{
						MessageBox.Show("No puede modificar valores de Mora si fecha de Vencimiento es menor a la fecha actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cell.Row.Cells["Mora"].Value = (decimal)e.Cell.Row.Cells["Mora"].OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Modifica Valores Locales

				string sSQLM = string.Format("Update Cre_TablaDeAmortizacion Set Mora = {0}, DsctMorManual = 1, DesMora = {2} Where idCre_TablaDeAmortizacion = {1}", 
					dMora, (int)e.Cell.Row.Cells["idCre_TablaDeAmortizacion"].Value, dMora);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLM, true);

				string sSQLGuardar = string.Format("Exec Cre_ActualizaGestionMora {0}, {1}, {2}, {3}, {4}, {5}",
					iCompra, iCreAmortizacion, OrigenValor, CambioValor, iCuota, iGastoMora);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLGuardar, true);
				
				this.ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
				bActualiza = false;
				return;
				//				}
			}
			#endregion Mora

			bActualiza = false;
		}

		private void gridPagos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.gridPagos);			
		}
		
		private void gridPagos_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			#region Gastos
			if (e.Cell.Column.ToString() == "Gastos")
			{ 				
				if ((int)e.Cell.Row.Cells["Estado"].Value == 2) e.Cancel = true;
				else if ((int)e.Cell.Row.Cells["Estado"].Value != 2 && (decimal)e.Cell.Row.Cells["Capital"].Value == 0.00m) e.Cancel = true;
				else if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 2 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
			}
			
			if (e.Cell.Column.ToString() == "DsctoGastos")
			{ 
				if ((int)e.Cell.Row.Cells["Estado"].Value == 0)	e.Cancel = true;
				if (Convert.ToDecimal(e.Cell.Row.Cells["DsctoGastos"].Value) > 100.00m || Convert.ToDecimal(e.Cell.Row.Cells["DsctoGastos"].Value) < 0.00m)	e.Cancel = true;				
			}
			#endregion Gastos

			#region Mora
			if (e.Cell.Column.ToString() == "Mora")
			{ 				
				if ((int)e.Cell.Row.Cells["Estado"].Value == 2) e.Cancel = true;
				if (!miAcceso.ModificarMoraCartera)
				{
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 2 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 15 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 16 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 19 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
				}
				else if ((int)this.txtDiasVencidos.Value < 100)
				{				
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 2 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 15 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 16 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 19 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) e.Cancel = true;
				}
			}
			
			if (e.Cell.Column.ToString() == "DsctoIntMora")
			{ 
				if ((int)e.Cell.Row.Cells["Estado"].Value == 0)	e.Cancel = true;
				if (Convert.ToDecimal(e.Cell.Row.Cells["DsctoIntMora"].Value) > 100.00m || Convert.ToDecimal(e.Cell.Row.Cells["DsctoIntMora"].Value) < 0.00m)	e.Cancel = true;
			}
			#endregion Mora
		}

		private void gridPagos_DoubleClick(object sender, System.EventArgs e)
		{
			if (this.gridPagos.ActiveRow.Cells["idCre_TablaDeAmortizacion"].Value == System.DBNull.Value) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From dbo.PagosCuotasCredito Where idCre_TablaDeAmortizacion = {0}", (int)this.gridPagos.ActiveRow.Cells["idCre_TablaDeAmortizacion"].Value)) > 0)
			{
				frmResumenPagosCrediPoint Resumen = new frmResumenPagosCrediPoint((int)this.gridPagos.ActiveRow.Cells["idCre_TablaDeAmortizacion"].Value, (int)this.gridPagos.ActiveRow.Cells["NumeroCuota"].Value, this.txtNombre.Text.ToString().Trim(), this.GridFacturas.ActiveRow.Cells["Numero"].Value.ToString());
				Resumen.ShowDialog();
			}
			else
			{
				MessageBox.Show("No existen cobros para esta Cuota", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}

		private void GridFacturas_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			bDescCovid = true;

			Cobranza();	
		}
		
		private bool ValidaSaldosRefinanciacion()
		{
			int iCuotasPendientesR = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CobranzaValidaRefinancion {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));

			if (iCuotasPendientesR > 0)
			{
				MessageBox.Show("No puede generar Cobro a documentos de refinanciacion si el credito inicial no ha sido cancelado en su totalidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}

			return true;
		}

		private bool ValidaSaldosAFavor()
		{
			bool bValidaSaldoPendiente = true;

			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && !(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && iDiasVencidos <= 0) bValidaSaldoPendiente = false;
			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 15 && !(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && iDiasVencidos <= 0) bValidaSaldoPendiente = false;
			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 16 && !(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && iDiasVencidos <= 0) bValidaSaldoPendiente = false;

//						if (bValidaSaldoPendiente && Convert.ToDecimal(this.txtSaldoAFavor.Value) > 0.00m)
//						{
//							MessageBox.Show("No puede registrar abonos hasta que se crucen los saldos pendientes\n\nComuniquese con el Departamento de Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);									
//							this.txtCobro.Focus();					
//							return false;
//						}

			//			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Compra c Inner Join MotivosDevolucion mt On mt.idMotivo = c.idMotivo Where c.idTipoFactura = 5 And Borrar = 0 And c.Estado < 8 And c.idOrigen = {0} And mt.Cruza = 1", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value)) > 0)
			//			{
			//				MessageBox.Show("El Documento tiene saldos por cruzar\n\nComuniquese con el Departamento de Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.txtIdentificacion.Focus();			
			//				return false;
			//			}

			return true;
		}

		private bool ValidaCobro()
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha", true, "Cobros", cdsSeteoF)) return false;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return false;

			if (!Validacion.vbIdentificacion(this.txtIdentificacion, 5, cdsSeteoF)) return false;

			if (!Validacion.vbClienteExiste(this.txtIdentificacion, false, 46, 9, cdsSeteoF)) return false;		
			
			if (GridFacturas.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtIdentificacion.Focus();
				return false;
			}
			if (GridFacturas.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.txtIdentificacion.Focus();
				return false;
			}
			if (this.gridPagos.Rows.Count == 0)
			{
				MessageBox.Show("Seleccione un documento para cargar la tabla de amortización", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.txtIdentificacion.Focus();
				return false;
			}

			if (Math.Round(Convert.ToDecimal(this.txtSaldoTotal.Value), 2) == 0.00m)
			{
				MessageBox.Show("El Documento esta Cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.txtIdentificacion.Focus();
				return false;
			}		
				

			return true;
		}	

		private void btCobro_Click(object sender, System.EventArgs e)
		{	
			string sIdentificacion = this.txtIdentificacion.Text.ToString();
			this.txtIdentificacion.Text = sIdentificacion;
			AbreactualizaClientes(sIdentificacion);
			if (!this.ValidaCobro()) return;

			int IApp = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT COUNT(1) FROM Cob_APPCobrosEfectivo WHERE idCompra = {0} AND Estado = 1  ", 
				(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value ));

			if( IApp > 0)
			{
				string NumeroS = Funcion.sEscalarSQL(cdsSeteoF, string.Format("SELECT top 1 numero FROM Cob_APPCobrosEfectivo WHERE idCompra = {0} AND Estado = 1  ", 
					(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value ));

				MessageBox.Show(string.Format("El Cliente tiene un cobro pendiente (Aplicación movil)\n N°: '{0}'", NumeroS ), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.dtFecha.Focus();		
				return;
			}

			if (Convert.ToDateTime(this.dtFecha.Value) < dtUltimaFechaPago)
			{
				MessageBox.Show(string.Format("La Fecha de Cobro no puede ser menor a '{0}'", dtUltimaFechaPago.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.dtFecha.Focus();		
				return;
			}
			
			if ((bool)this.GridFacturas.ActiveRow.Cells["Reliquido"].Value)
			{
				MessageBox.Show("El Documento esta reliquidado\n\nComuniquese con el Departamento de Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();
				return;
			}
						
			int iDiasUltimoCobro = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CobranzaValidaDiasUltimoPago {0}, {1}, {2}", 
				(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, (int)this.txtIdCliente.Value, (int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value));

			string sMensaje = "";
			if (iDiasUltimoCobro == 0) sMensaje = "El Cliente tiene un pago registrado el dia de hoy\n\n¿Esta seguro de volver a cobrar?";
			if (iDiasUltimoCobro > 0 && iDiasUltimoCobro <= 10) sMensaje = string.Format("El Cliente tiene un pago registrado hace {0} dias\n\n¿Esta seguro de volver a cobrar?", iDiasUltimoCobro);

			if (sMensaje.Length > 0 && DialogResult.No == MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				this.txtIdentificacion.Focus();					
				return;
			}
		
			if (!Validacion.vbCampoDecimalVacio(this.txtCobro, "Ingrese el monto a cobrar", 0.01m, Convert.ToDecimal(this.txtSaldoTotal.Value))) return;					
			
			if (!this.ValidaSaldosAFavor()) return;
			
			if (!ValidaSaldosRefinanciacion()) return;

			#region Comentado
			//			FuncionesProcedimientos.AbreFacturacionClientes(5, this.txtIdentificacion.Text.ToString(), false);

			//			decimal dDiferencia = Convert.ToDecimal(this.txtCobro.Value) - Decimal.Round(dValorCobro, 2);
			//			if ((Math.Round(Convert.ToDecimal(this.txtCobro.Value), 2) < Decimal.Round(dValorCobro, 2)))// && iDiasVencidos >= 0)
			//			{
			//				MessageBox.Show(string.Format("No puede cobrar menos de : {0}.", Math.Round(dValorCobro, 2)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);									
			//				this.txtCobro.Value = Decimal.Round(dValorCobro, 2);
			//				this.txtCobro.Focus();					
			//				return;
			//			}
			#endregion Comentado
		
			ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
			
			//FuncionesProcedimientos.AbreFacturacionClientes(5, this.txtIdentificacion.Text.ToString(), false, miAcceso.EditarCuentaNombre);

			bool bCobroManual = false;
			bool bCovid = false;
			int iTipo = 12;

			#region Cobro Manual
			/*
			//if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && !(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && iDiasVencidos < 0) bCobroManual = true;
			//if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 13 && !(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && iDiasVencidos < 0) bCobroManual = true;
			//if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 14 && !(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && iDiasVencidos < 0) bCobroManual = true;
			//if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && (int)this.txtDiasVencidos.Value > 7200) bCobroManual = true;

			bCobroManual = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaCobroManual {0}, {1}", 
				(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDecimal(this.txtCobro.Value)), true);

			if (!Validacion.vbCarteraEnTransito((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value, 
				(int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value, (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value)) bCobroManual = true;

			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 2) 
			{
				bCobroManual = true;
				iTipo = 5;
			}
      			
			if (bCobroManual)
			{
				using (frmAnticiposClientes Anticipos = new frmAnticiposClientes (true, this.txtIdentificacion.Text.ToString(), (int)this.cmbBodega.Value, Convert.ToDecimal(this.txtCobro.Value), iTipo,
								 (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value))
				{
					if (DialogResult.OK == Anticipos.ShowDialog())
					{
						Facturacion.ImpresionRecibos((int)Anticipos.txtIdAnticipo.Value, iTipo, cdsSeteoF);

						if ((int)this.txtDiasVencidos.Value < 1)
						{
							#region Boletos
//							using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
//							{
//								oConexion.Open();
//
//								SqlTransaction oTransaction = null;
//								SqlCommand oCmdActualiza = oConexion.CreateCommand();
//								oCmdActualiza.Connection = oConexion;
//								oCmdActualiza.CommandType = CommandType.Text;
//								oCmdActualiza.CommandTimeout = 0;
//						
//								try
//								{							
//									oTransaction = oConexion.BeginTransaction();
//									oCmdActualiza.Transaction = oTransaction;
//
//									string sSQL = string.Format("Exec ValidaCuotasPagadasAnticipoSorteo {0}, {1}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDecimal(this.txtCobro.Value));
//									oCmdActualiza.CommandText = sSQL;
//									int iCuotasPagadas = (int)oCmdActualiza.ExecuteScalar();
//
//									#region Valida Cuotas Pagadas
//									if (iCuotasPagadas > 0)
//									{
//										iCuotasPagadas = iCuotasPagadas * 3;
//							
//										int j=1;	
//	
//										for(j=1;j<=iCuotasPagadas;j++)
//										{
//											string ssSQL = string.Format("Exec NumeracionLocales 68, 0, 0");									
//											oCmdActualiza.CommandText = ssSQL;
//											string num_sorteo = oCmdActualiza.ExecuteScalar().ToString();
//									
//											Val_Cuota = 0.00m;
//
//											string datosSQL = string .Format("Insert into sorteo_ticket(numero,id_Compra,id_anticipo,Estado,bodega,num_cuota,Valor_Cuota) values ('{0}',{1},{2},1,{3},{4},{5})", 
//												num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,(int)Anticipos.txtIdAnticipo.Value,(int)this.cmbBodega.Value,0,0.00m);
//											oCmdActualiza.CommandText = datosSQL;
//											oCmdActualiza.ExecuteNonQuery();
//										}
//								
//										oTransaction.Commit();
//
//										#region Impresion Sorteo
//										ParameterFields paramFields = new ParameterFields();
//
//										ParameterField pfIdOrigen2 = new ParameterField();
//										ParameterDiscreteValue ValIdOrigen2 = new ParameterDiscreteValue();
//										pfIdOrigen2.ParameterFieldName = "@idAnticipo";
//										ValIdOrigen2.Value = (int)Anticipos.txtIdAnticipo.Value;
//										pfIdOrigen2.CurrentValues.Add(ValIdOrigen2);
//										paramFields.Add(pfIdOrigen2);
//
//										Reporte miRep = new Reporte("sorteo_ticket1.rpt", "");
//										miRep.MdiParent = this.MdiParent;
//										miRep.Titulo("Cupones de  Sorte");
//										miRep.crVista.ParameterFieldInfo = paramFields;
//										miRep.ShowDialog();
//										#endregion Impresion Sorteo
//									}
//									#endregion Valida Cuotas Pagadas
//
//									return;
//								}
//								catch (Exception ex)
//								{
//									oTransaction.Rollback();
//
//									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//								}
//								finally
//								{
//									oConexion.Close();
//								}
//							}
						#endregion Boletos
						}
					}				
				}

				return;
			}
			*/
			#endregion Cobro Manual

			#region Registro de Cobro
			int idCompraPago = 0;
			bool bNormalAtrasado = false;
      
			using (VentaFormaDePago Caja = new VentaFormaDePago (idCompraPago, 1, Convert.ToDecimal(this.txtCobro.Value), 3, (DateTime)this.dtFecha.Value, (int)this.cmbBodega.Value, miAcceso.BFiniquitos))
			{
				if (DialogResult.OK == Caja.ShowDialog())
				{
					this.Cursor = Cursors.WaitCursor;

					using (SqlConnection oConCobranzas = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
					{
						oConCobranzas.Open();

						SqlTransaction oTransaction = null;
						SqlCommand oCmdActualiza = oConCobranzas.CreateCommand();
						oCmdActualiza.Connection = oConCobranzas;
						oCmdActualiza.CommandType = CommandType.Text;
						oCmdActualiza.CommandTimeout = 120;
						
						try
						{							
							oTransaction = oConCobranzas.BeginTransaction();
							oCmdActualiza.Transaction = oTransaction;

							#region Variables
							decimal Cobro = Convert.ToDecimal(this.txtCobro.Value);
							decimal Diferimiento = 0.00m;
							decimal ABDiferimiento = 0.00m;
							decimal DescDiferimiento = 0.00m;
							decimal Gasto = 0.00m;
							decimal IntMora = 0.00m;
							decimal IntFin = 0.00m;
							decimal Capital = 0.00m;
							decimal Abono = 0.00m;
							decimal Saldo = Convert.ToDecimal(this.txtCobro.Value);
							decimal NSaldo = 0.00m;							
							decimal AbGastos = 0.00m;
							decimal AbIntMora = 0.00m;
							decimal AbInteres = 0.00m;
							decimal AbCapital = 0.00m;							
							int Estado = 0;
							string sConcepto = "";
							int idCompraOrigen = (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value;							
							#endregion Variables
														
							#region Genera Abono Factura 24
							string sSQLNumero = string.Format("Exec NumeracionLocales 24, {0}, 0", (int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							string sNumero = oCmdActualiza.ExecuteScalar().ToString(); 

							if (Cobro < Convert.ToDecimal(this.txtSaldoVencido.Value)) bNormalAtrasado = true;

							string sSQLPago = string.Format("Exec GuardaPagoCobranza {0}, {1}, {2}, 0, '{3}', '{4}', {5}", 
								idCompraOrigen, (int)this.cmbBodega.Value, Math.Round(Convert.ToDecimal(this.txtCobro.Value), 2), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumero, bNormalAtrasado);
							oCmdActualiza.CommandText = sSQLPago;								
							idCompraPago = (int)oCmdActualiza.ExecuteScalar(); 
							#endregion Genera Abono Factura 24

							#region Guarda Cobro de la factura
							if (Caja.ultraGrid1.Rows.Count > 0)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Caja.ultraGrid1.Rows.All)
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
									if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
									if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = Convert.ToDecimal(dr.Cells["ValorUsoTarjeta"].Value);

									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
										idDetFormaPago, 
										idCompraPago, 3, 
										(int)dr.Cells["idFormaPago"].Value, Math.Round(Convert.ToDecimal(dr.Cells["Valor"].Value), 2),
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
							#endregion Guarda Cobro de la factura  										

							#region Calculo de Cobranza
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
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
									//EstadoSorteo = 0;									

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
											ABDiferimiento = ABDiferimiento + Diferimiento;
											DescDiferimiento = DescDiferimiento + Diferimiento;											
										}
										else
										{											
											Abono = Abono + Diferimiento;
											ABDiferimiento = ABDiferimiento + Diferimiento;
											Saldo = Saldo - Diferimiento;
										}
									}									
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
									{
										Diferimiento = Diferimiento - Saldo;
										Abono = Abono + Saldo;
										ABDiferimiento = ABDiferimiento + Saldo;
										Saldo = 0.00m;
									}
									#endregion Diferimiento
									
									#region Gestión de cobranza
									if (Math.Round(Gasto, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Gasto, 2))
									{								
										Abono = Abono + Gasto;
										AbGastos = AbGastos + Gasto;
										Saldo = Saldo - Gasto;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Gasto, 2) > 0.00m && Math.Round(Gasto, 2) >= Math.Round(Saldo, 2)) 
									{
										Gasto = Gasto - Saldo;
										Abono = Abono + Saldo;
										AbGastos = AbGastos + Saldo;
										Saldo = 0.00m;
									}
									#endregion Gestión de cobranza

									#region Interés por mora
									if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
									{
										Abono = Abono + IntMora;
										AbIntMora = AbIntMora + IntMora;
										Saldo = Saldo - IntMora;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
									{
										IntMora = IntMora - Saldo;
										Abono = Abono + Saldo;
										AbIntMora = AbIntMora + Saldo;
										Saldo = 0.00m;
									}
									#endregion Interés por mora
						
									#region Interés por financiamiento
									if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
									{
										Abono = Abono + IntFin;
										AbInteres = AbInteres + IntFin;
										Saldo = Saldo - IntFin;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
									{
										IntFin = IntFin - Saldo;
										Abono = Abono + Saldo;
										AbInteres = AbInteres + Saldo;
										Saldo = 0.00m;
									}
									#endregion Interés por financiamiento

									#region Capital
									if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
									{
										Abono = Abono + Capital;
										AbCapital = AbCapital + Capital;
										Saldo = Saldo - Capital;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
									{
										Capital = Capital - Saldo;
										Abono = Abono + Saldo;
										AbCapital = AbCapital + Saldo;
										Saldo = 0.00m;								
									}
									#endregion Capital

									#region Actualiza tabla Cre_TablaDeAmortización
									if (Abono != 0.00m)
									{										
										decimal dSaldoPrb = 0.00m;										
										dSaldoPrb = Convert.ToDecimal(dr.Cells["Saldo"].Value);										
										NSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value) - Abono - ABDiferimiento;

										string sSQLAC = string.Format("Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}", 
											(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value, 
											Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idCompraPago, AbInteres, AbCapital, AbGastos, AbIntMora, ABDiferimiento, DescDiferimiento);
										oCmdActualiza.CommandText = sSQLAC;
										oCmdActualiza.ExecuteNonQuery();

										#region Estado Sorteo
										sConcepto = "";

										if (Math.Round(NSaldo, 2) <= 0.00m)
										{					
											Estado = 2;
											EstadoSorteo = 1;
											EstadoCelular=1;
											sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();	
												 
											int num_cuota  = Convert.ToInt32(dr.Cells["NumeroCuota"].Value.ToString());											

											//											#region Boletos de sorteo
											//											if (Estado == 2)//EstadoSorteo=1;
											//											{												
											//												if (Convert.ToDateTime(this.dtFecha.Value) <= Convert.ToDateTime(dr.Cells["Vence"].Value )) 
											//												{
											//													int j=1;		
											//													for(j=1;j<=3;j++)
											//													{
											//														string ssSQL = string.Format("Exec NumeracionLocales 68, 0, 0");
											//														oCmdActualiza.CommandText = ssSQL;														
											//														string num_sorteo  = oCmdActualiza.ExecuteScalar().ToString();														
											//														Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);														
											//														string datosSQL = string .Format("Insert into sorteo_ticket(numero,id_Compra,id_anticipo,Estado,bodega,num_cuota,Valor_Cuota) values ('{0}',{1},{2},1,{3},{4},{5})", num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,(int)this.cmbBodega.Value,num_cuota,Val_Cuota);
											//														oCmdActualiza.CommandText = datosSQL;
											//														oCmdActualiza.ExecuteNonQuery();
											//													}
											//												}
											//												else 	if ( Convert.ToDateTime(this.dtFecha.Value) > Convert.ToDateTime(dr.Cells["Vence"].Value ))
											//												{
											//													string ssSQL = string.Format("Exec NumeracionLocales 68, 0, 0");
											//													oCmdActualiza.CommandText = ssSQL;
											//													string num_sorteo  = oCmdActualiza.ExecuteScalar().ToString();	
											//													
											//													Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
											//
											//													string datosSQL = string .Format("Insert into sorteo_ticket(numero,id_Compra,id_anticipo,Estado,bodega,num_cuota,Valor_Cuota)values ('{0}',{1},{2},1,{3},{4},{5})", num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,(int)this.cmbBodega.Value,num_cuota,Val_Cuota);
											//													oCmdActualiza.CommandText = datosSQL;
											//													oCmdActualiza.ExecuteNonQuery();
											//												}												
											//											}
											//											#endregion Boletos de sorteo
											//
											//											#region Impresion Sorteo 
											//											//												ParameterFields paramFields = new ParameterFields();
											//											//
											//											//												ParameterField pfIdOrigen2 = new ParameterField();
											//											//												ParameterDiscreteValue ValIdOrigen2 = new ParameterDiscreteValue();
											//											//												pfIdOrigen2.ParameterFieldName = "@idAnticipo";
											//											//												ValIdOrigen2.Value = idCompraPago;
											//											//												pfIdOrigen2.CurrentValues.Add(ValIdOrigen2);
											//											//												paramFields.Add(pfIdOrigen2);
											//											//
											//											//												Reporte miRep = new Reporte("sorteo_ticket1.rpt", "");
											//											//												miRep.MdiParent = this.MdiParent;
											//											//												miRep.Titulo("Cupones de  Sorteo");
											//											//												miRep.crVista.ParameterFieldInfo = paramFields;
											//											//												miRep.Show();
											//											#endregion Impresion Sorteo
										}
										else if (Math.Round(NSaldo, 2) > 0.00m) 
										{
											Estado = 3;
											sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
										}
										#endregion Estado Sorteo

										string sSQLTipoPago = string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
										oCmdActualiza.CommandText = sSQLTipoPago;
										int iPago = (int)oCmdActualiza.ExecuteScalar();

										#region  Descuento COVID 
										decimal dDescGestion = 0.00m;
										//										string sSQLDescGestion = string.Format("Select ISNULL (DescGestion ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
										//										oCmdActualiza.CommandText = sSQLDescGestion;
										//										decimal dDescGestion = (decimal)oCmdActualiza.ExecuteScalar();
										//
										//										string sSQLDesMora = string.Format("Select ISNULL (DesMora ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
										//										oCmdActualiza.CommandText = sSQLDesMora;
										//										decimal dDesMora = (decimal)oCmdActualiza.ExecuteScalar();
										decimal dDesMora = 0.00m;
										//										if (Diferimiento == 0.00m)
										//										{											
										//											oCmdActualiza.CommandText = string.Format("Select ISNULL (DescDiferimiento, 0) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
										//											dDescuentoDiferimiento = (decimal)oCmdActualiza.ExecuteScalar();										
										//											
										//											if (dDescuentoDiferimiento > 0.00m) ABDiferimiento = dDescuentoDiferimiento;
										//										}
										#endregion  Descuento COVID 

										decimal dDebitado = AbIntMora + AbInteres + AbCapital + ABDiferimiento;
										
										decimal dRescate = 0.00m;
										
										string sSQLRescate = string.Format("Exec Cre_ValidaSaldoCuotaRescatada {0}, {1}, '{2}'", 
											(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, dDebitado, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
										oCmdActualiza.CommandText = sSQLRescate;
										dRescate = (decimal)oCmdActualiza.ExecuteScalar();

										string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}, {17}, {18}, {19}",
											idCompraPago, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, AbGastos, AbIntMora, AbInteres, AbCapital, Abono, //6
											Math.Round(NSaldo, 2), Estado, sConcepto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iPago, dDebitado, this.GridFacturas.ActiveRow.Cells["NumeroDeOperacion"].Value.ToString(), //14
											0, dDescGestion, dDesMora, ABDiferimiento, DescDiferimiento, dRescate);
										oCmdActualiza.CommandText = sSQLPCC;
										oCmdActualiza.ExecuteNonQuery();

										#region Actualiza Pago en gestiones
										decimal dValorPagadoGestiones = AbCapital + AbInteres;
										string sSQLAPG = string.Format("Exec Cre_ActualizaValorPagoGestiones {0}, '{1}', '{2}', {3}, {4}",
											idCompraOrigen, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(dr.Cells["Vence"].Value).ToString("yyyyMMdd"), dValorPagadoGestiones, Abono);
										oCmdActualiza.CommandText = sSQLAPG;
										oCmdActualiza.ExecuteNonQuery();
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
								oCmdActualiza.CommandText = sSQLHA;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Historial de autorizaciones

							#region Asiento - Saldos
							//							string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, 0, {1}", idCompraPago, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
							//							oCmdActualiza.CommandText = sSQLAsiento;
							//							oCmdActualiza.ExecuteNonQuery();
							
							//							oCmdActualiza.CommandText = string.Format("Exec AsientoCuentasDeOrdenCobranzaPichincha {0}, 0", idCompraPago);
							//							oCmdActualiza.ExecuteNonQuery();

							string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 9 Exec ActualizaSaldoFacturas {1}, 1, 9", idCompraPago, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
							oCmdActualiza.CommandText = sSQLSaldo;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLCarteraExterna = string.Format("Exec GeneraMovimientosCarteraExterna {0}, {1}", idCompraPago, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
							oCmdActualiza.CommandText = sSQLCarteraExterna;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Asiento - Saldos
							
							#region Actualiza Compromisos de Pago - Estado de Cuenta 
							string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}", 
								(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDecimal(this.txtCobro.Value));
							oCmdActualiza.CommandText = sSQLActCompromisoPago;
							oCmdActualiza.ExecuteNonQuery();								

							string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
							oCmdActualiza.CommandText = sSQLAEC;
							oCmdActualiza.ExecuteNonQuery();
							#endregion Actualiza Compromisos de Pago - Estado de Cuenta

							#region Notifiacion Pago
							string sSQLNotificaPago = string.Format("Exec GuardaNotifiacionCuotaPagos {0}, {1}, {2},{3}, {4}",(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,0,(int)this.cmbBodega.Value,1);
							oCmdActualiza.CommandText = sSQLNotificaPago;	
							oCmdActualiza.ExecuteNonQuery();	
													
							#endregion Notifiacion Pago
							
							//							#region Actualiza Valor Cobrado Gestiones de Credito_Cobranza
							//							string sSQLActCompromisoCredito_Cobranza = string.Format("Exec Cbo_ActualizaCompromisoPago {0}, '{1}', {2}", 
							//								(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDecimal(this.txtCobro.Value));
							//							oCmdActualiza.CommandText = sSQLActCompromisoCredito_Cobranza;
							//							oCmdActualiza.ExecuteNonQuery();
							//							#endregion Actualiza Valor Cobrado Gestiones de Credito_Cobranza

							#region Libera celulares	
							//int iValEstadoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select  COUNT(1)From dbo.Cre_TablaDeAmortizacion Where Estado Not In (2, 4)  and idCompra = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
							oCmdActualiza.CommandText = string.Format("Select COUNT(1) From dbo.Cre_TablaDeAmortizacion Where Estado Not In (2, 4) and idCompra = {0}", 
								(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
							int iValEstadoCel = (int) oCmdActualiza.ExecuteScalar();

							if (EstadoCelular==1)
							{															
								string sSQLBLO = string.Format("Exec LiberacionDeCelulares {0}",(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
								oCmdActualiza.CommandText = sSQLBLO;
								oCmdActualiza.ExecuteNonQuery();									
							}
							#endregion Libera celulares

							oTransaction.Commit();	
							oConCobranzas.Close();
							Cobranza();	
							int idBonoEntradaTipo = 0;

							//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
							DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");

							if (dvResultado != null)
							{
									foreach (DataRowView rowView in dvResultado)
									{
										DataRow row = rowView.Row;
										idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
										//Validar si el bono se encuentra en estado activo
										string query = String.Format("Exec BonoEntradaTipoEstado {0}", idBonoEntradaTipo);
										bool bBonoEntradaTipoEstado = Funcion.bEscalarSQL(cdsSeteoF, query, false);
										if(bBonoEntradaTipoEstado)
										{
											#region Comprobar Bono 25
													int iBonoEntradaIdCompra = 0;
													int PromedioMora = 0;//Es un valor por default, debería calcularse si es la última cuota
													bool PagoUltimaCuota = true;
													string squery = "";
													iBonoEntradaIdCompra = (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value;
													string Numero = "";
													int getIdBonoEntrada = 0;
													bool TieneBono = false;
													int idBonoEntrada = 0;

													//Validar si todo ya está cancelado
													squery = String.Format("Exec BonoEntradaCancelacionTotal {0}, {1}", iBonoEntradaIdCompra, idBonoEntradaTipo);
													PagoUltimaCuota = Funcion.bEscalarSQL(cdsSeteoF, squery, true);
			
													//Entra si ya todo está cancelado
													if(PagoUltimaCuota)
													{
														squery = String.Format("Exec BonoEntradaPromedioMora {0}", iBonoEntradaIdCompra);
														PromedioMora = Funcion.iEscalarSQL(cdsSeteoF, squery);
														//Entra si cumple la condición que la mora sea menor o igual a 0
														squery = String.Format("Exec BonoEntradaTipoModel '{0}', {1}", "CondicionMoraDias", idBonoEntradaTipo);
														int CondicionMoraDias = Funcion.iEscalarSQL(cdsSeteoF, squery);
														if(PromedioMora <= CondicionMoraDias)
														{
															//Validar si tiene un bono entrada activo
															if((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value > 0)
															{
																squery = String.Format("Exec BonoEntradaTieneBonoXIdCompra {0}, {1}", iBonoEntradaIdCompra, idBonoEntradaTipo);
																TieneBono = Funcion.bEscalarSQL(cdsSeteoF, squery, false);
																if(TieneBono)
																{
																	squery = String.Format("Exec BonoEntradaConId {0}, {1}", iBonoEntradaIdCompra, idBonoEntradaTipo);
																	idBonoEntrada = Funcion.iEscalarSQL(cdsSeteoF, squery);

																	if(idBonoEntrada > 0)
																	{
																		#region Actualización del bono
																		#region BonoEntradaCrear
																		try
																		{
																			//Guardar Bono Entrada 25
																			//BONO DE 25$ AL MOMENTO DE IMPRIMIR EL ULTIMO PAGO DE LOS CLIENTES CUYA MORA PROMEDIO 
																			//HAYA SIDO DE MENOR O IGUAL A CERO, ADICIONAL, ESTE BONO SE PODRA USAR EN UN LAPSO DE 
																			//1 MES SOLO A CREDITO DIRECTO PARA EL MISMO CLIENTE Y CUYO NUEVO CREDITO SEA MAYOR A $ 800
																			squery = String.Format("EXEC BonoEntradaCrear {0}, {1}, {2}, {3}, '{4}', {5}", 
																				idBonoEntrada, 
																				idBonoEntradaTipo, 
																				iBonoEntradaIdCompra,
																				1,
																				"",
																				idCompraPago);
																			getIdBonoEntrada = Funcion.iEscalarSQL(cdsSeteoF, squery);
																		}
																		catch(Exception ex)
																		{
																			Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
																			MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
																			return;
																		}
																		#endregion BonoEntradaCrear
																		#endregion Actualización del bono
																	}
																}
																else
																{
																	#region Creación del bono
																	#region Numeracion
																	try
																	{
																		//Fijarse bien en el id de la numeración en producción porque puede ser otra.
																		squery = string.Format("Exec NumeracionLocales {0}, 0, 0", NumeracionLocalesBonoEntrada);
																		Numero = Funcion.sEscalarSQL(cdsSeteoF, squery);
																	}
																	catch(Exception ex)
																	{
																		Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
																		MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
																		return;
																	}
																	#endregion Numeracion
																	#region BonoEntradaCrear
																	try
																	{
																		//Guardar Bono Entrada 25
																		//BONO DE 25$ AL MOMENTO DE IMPRIMIR EL ULTIMO PAGO DE LOS CLIENTES CUYA MORA PROMEDIO 
																		//HAYA SIDO DE MENOR O IGUAL A CERO, ADICIONAL, ESTE BONO SE PODRA USAR EN UN LAPSO DE 
																		//1 MES SOLO A CREDITO DIRECTO PARA EL MISMO CLIENTE Y CUYO NUEVO CREDITO SEA MAYOR A $ 800
																		squery = String.Format("EXEC BonoEntradaCrear {0}, {1}, {2}, {3}, '{4}', {5}", 
																			0, 
																			idBonoEntradaTipo, 
																			iBonoEntradaIdCompra,
																			1,
																			Numero,
																			idCompraPago);
																		getIdBonoEntrada = Funcion.iEscalarSQL(cdsSeteoF, squery);
																	}
																	catch(Exception ex)
																	{
																		Funcion.LogSistema(cdsSeteoF, this.Name + ": "+ this.Text, string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), squery, Funcion.slogtipoerror);
																		MessageBox.Show(Funcion.smensajeerror, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
																		return;
																	}
																	#endregion BonoEntradaCrear
																	#endregion Creación del bono
																}
															}
														}
													}
													#endregion Comprobar Bono 25
										}
									}
							}
							else
							{
									Console.WriteLine("El resultado del procedimiento es nulo.");
							}

							#region Impresion
							using (frmNumeracionRecibo miMsje = new frmNumeracionRecibo (Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Secuencial, '') From Anticipos Where idAnticipo = {0}", idCompraPago))))
							{
								if (DialogResult.OK == miMsje.ShowDialog())
									Facturacion.ImpresionRecibos(idCompraPago, 2, cdsSeteoF);
							}
							#endregion Impresion

							#region Impresion Sorteo
							//							if (EstadoSorteo == 1 ){ 
							//
							//								ParameterFields paramFields = new ParameterFields();
							//
							//								ParameterField pfIdOrigen2 = new ParameterField();
							//								ParameterDiscreteValue ValIdOrigen2 = new ParameterDiscreteValue();
							//								pfIdOrigen2.ParameterFieldName = "@idAnticipo";
							//								ValIdOrigen2.Value = idCompraPago;
							//								pfIdOrigen2.CurrentValues.Add(ValIdOrigen2);
							//								paramFields.Add(pfIdOrigen2);
							//								
							//								Reporte miRep = new Reporte("sorteo_ticket1.rpt", "");
							//								miRep.MdiParent = this.MdiParent;
							//								miRep.Titulo("Cupones de  Sorteo");
							//								miRep.crVista.ParameterFieldInfo = paramFields;
							//								miRep.Show();	
							//								EstadoSorteo = 0;
							//							}
							//							
							#endregion Impresion Sorteo
								
							#region ImpresionCompromisoPago
							//							bool iComprPago =  Funcion.bEscalarSQL(cdsSeteoF, string.Format("select isnull (CompromisoPago, 0) from Compra where convert (date, Fecha) <= '20190418' and idCompra = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value), true);
							//							if (iComprPago == false)
							//							{
							//								Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE Compra set CompromisoPago = 1 where idCompra  = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
							//		            
							//							
							//
							//								//			ParameterFields paramFields = new ParameterFields ();
							//
							//								ParameterField pfIdOrigen = new ParameterField();
							//								ParameterDiscreteValue ValIdOrigen = new ParameterDiscreteValue();
							//								pfIdOrigen.ParameterFieldName = "@idcompra";
							//								ValIdOrigen.Value = Convert.ToInt32((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value );
							//								pfIdOrigen.CurrentValues.Add(ValIdOrigen);
							//								paramFields.Add(pfIdOrigen);
							//				
							//								Reporte Reporte = new Reporte("Cre_CompromisoLugarDePagoActualizado.rpt", "");
							//								Reporte.MdiParent = this.MdiParent;			
							//								Reporte.crVista.ParameterFieldInfo = paramFields;
							//								Reporte.Show();
							//								 
							//							}
							#endregion ImpresionCompromisoPago											

							Cobranza();	

							this.Cursor = Cursors.Default;
						}
						catch (Exception ex)
						{
							try
							{
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

								oConCobranzas.Close();
							}
							catch (Exception ex2)
							{
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							

								oConCobranzas.Close();
							}
						}
						finally
						{
							oConCobranzas.Close();
						}
					}
				}
			}
			#endregion Registro de Cobro
								
			this.txtCobro.Enabled = false;
			this.btCobro.Enabled = false;
			this.txtCobro.Value = 0.00m;
			bMayorCobro = false;
			bCovid = false;
			dValorCobro = 0.00m;

			this.txtIdAutorizacion.Value = 0;
			this.txtValorAutorizado.Value = 0.00m;
			this.txtEstadoAutorizacion.Value = 0;
			this.btnCancelar.Focus();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnActualizarEDsctoGM_Click(object sender, System.EventArgs e)
		{
			if (this.GridFacturas.ActiveRow == null) return;

			if (this.GridFacturas.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value) return;
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cre_TablaDeAmortizacion Set DsctGasManual = 0, DsctMorManual = 0, DsctoIntMora = 0 Where idCompra = {0} And Vendida = 0", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.dtFecha.Value = DateTime.Today;

			LimpiaControles(true, true);
		}
		
		private void btnAnticipo_Click(object sender, System.EventArgs e)
		{	
			#region Validacion
			if (!this.ValidaCobro()) return;

			if (Convert.ToDateTime(this.dtFecha.Value) < dtUltimaFechaPago)
			{
				MessageBox.Show(string.Format("La Fecha de Cobro no puede ser menor a '{0}'", dtUltimaFechaPago.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.dtFecha.Focus();
				return;
			}
//
			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 2)
			{
				MessageBox.Show("Cartera de Solidario no puede cruzar saldos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();
				return;
			}
//			
			if (!Validacion.vbCarteraEnTransito((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value, 
				(int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value, (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value)) return;
			
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaAnticiposACruzarCrediPoint '{0}', '{1}', {2}", this.txtIdentificacion.Text.ToString(),  
				Convert.ToDateTime(this.GridFacturas.ActiveRow.Cells["Fecha"].Value).ToString("yyyyMMdd"), (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value)) == 0)
			{
				MessageBox.Show(string.Format("No Existen Anticipos o Notas De Credito Pendientes del Cliente '{0}'", this.txtNombre.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.txtIdentificacion.Focus();
				return;
			}

			if (((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 || 
				(int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 15 || 
				(int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 16) && 
				!(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value 
				&& iDiasVencidos < 0 && !miAcceso.BCruzarPagosAdelantados && 
				!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Reliquido From Compra Where idCompra = {0}", 
				(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value), true))
			{
				MessageBox.Show("El documento no esta en Mora\n\nEn Documentos pendientes de envio al Banco se cruzan pagos cuando estan en Mora", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();	
				return;
			}
//
//			//			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && (int)this.txtDiasVencidos.Value > 290)
//			//			{
//			//				MessageBox.Show("Documento en Fondo de Garantia.\n\nProcese el debito del Banco para cruzar el documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//			//				this.txtIdentificacion.Focus();			
//			//				return;
//			//			}
//
//			if (!ValidaSaldosRefinanciacion()) return;	
//
////			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCobroManualCOVID {0} ", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value), true) > 0)
////			{
////				MessageBox.Show("Documento en proceso de diferimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
////				this.txtIdentificacion.Focus();			
////				return;
////			}
			#endregion Validacion

			#region Variables
			decimal Gasto = 0.00m;
			decimal IntMora = 0.00m;
			decimal IntFin = 0.00m;
			decimal Capital = 0.00m;
			decimal Abono = 0.00m;				
			decimal NSaldo = 0.00m;				
			decimal Diferimiento = 0.00m;
			decimal ABDiferimiento = 0.00m;	
			decimal DescDiferimiento = 0.00m;	
			decimal AbGastos = 0.00m;
			decimal AbIntMora = 0.00m;
			decimal AbInteres = 0.00m;
			decimal AbCapital = 0.00m;
								
			int Estado = 0;
			DateTime dtFecha = DateTime.Today;
			string sSQL = "";
			string sConcepto = "";

			int iDescDif = Funcion.iEscalarSQL(cdsSeteoF, 
				string.Format("Select COUNT(1) From dbo.Temporal_Operaciones_Dif_Final Where idCompra = {0} And Estado = 3 And EstadoPoint = 3", 
				(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));

			if (dSaldoSD <= Convert.ToDecimal(this.txtSaldoAFavor.Value)) iDescDif++;
			#endregion Variables

			#region Registro de Cobro
			using (frmCruceAnticiposCrediPoint Cruces = new frmCruceAnticiposCrediPoint(this.txtIdentificacion.Text.ToString(), Convert.ToDecimal(this.txtSaldoTotal.Value), Convert.ToDateTime(this.GridFacturas.ActiveRow.Cells["Fecha"].Value), Convert.ToDateTime(this.dtFecha.Value), Vence, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value))										
			{
				if (DialogResult.OK == Cruces.ShowDialog())
				{
					this.Cursor = Cursors.WaitCursor;
					decimal Cobro = Convert.ToDecimal(Cruces.uGrid.ActiveRow.Cells["Saldo"].Value);
					decimal Saldo = Convert.ToDecimal(Cruces.uGrid.ActiveRow.Cells["Saldo"].Value);
					
					idanticipo2 = Convert.ToInt32(Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value);			

					#region Calculo de Cobranza
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
					{
						if (Saldo > 0.00m)
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
							Gasto = Convert.ToDecimal(dr.Cells["Gastos"].Value);
							IntMora = Convert.ToDecimal(dr.Cells["Mora"].Value);
							IntFin = Convert.ToDecimal(dr.Cells["Interes"].Value);
							Capital = Convert.ToDecimal(dr.Cells["Capital"].Value);							
							Abono = 0.00m;
							
							#region Diferimiento
							if (Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2))
							{
								if (Convert.ToDateTime(dr.Cells["Vence"].Value) >= Convert.ToDateTime(this.dtFecha.Value) && (Capital + IntFin + IntMora + Gasto) <= Saldo)
								{
									ABDiferimiento = ABDiferimiento + Diferimiento;
									DescDiferimiento = DescDiferimiento + Diferimiento;											
								}
								else
								{											
									Abono = Abono + Diferimiento;
									ABDiferimiento = ABDiferimiento + Diferimiento;
									Saldo = Saldo - Diferimiento;
								}
							}									
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
							{
								Diferimiento = Diferimiento - Saldo;
								Abono = Abono + Saldo;
								ABDiferimiento = ABDiferimiento + Saldo;
								Saldo = 0.00m;
							}
							#endregion Diferimiento

							#region Diferimiento Comentado
							//							if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2) && Math.Round(Diferimiento, 2) != 0)
							//							{
							//								if (iDescDif > 0)
							//								{
							//									if (Diferimiento < 0.00m)
							//									{
							//										decimal dValorCuotas = Convert.ToDecimal(this.GridFacturas.ActiveRow.Cells["ValorCuotas"].Value);
							//										decimal dAbonoCD = Convert.ToDecimal(dr.Cells["Abono"].Value);										
							//										DescDiferimiento = IntFin +  Capital + dAbonoCD - Math.Abs(Diferimiento) - dValorCuotas;
							//									}
							//									else DescDiferimiento = Diferimiento;
							//								}
							//								//Diferimiento = Diferimiento - DescDiferimiento;
							//								Abono = Abono + Diferimiento; //- Math.Abs(DescDiferimiento);
							//								ABDiferimiento = ABDiferimiento + Diferimiento;								
							//								Saldo = Saldo - Diferimiento + DescDiferimiento;
							//							}
							//							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
							//							{
							//								if (iDescDif > 0) DescDiferimiento = Saldo;
							//								Diferimiento = Diferimiento - Saldo;
							//								Abono = Abono + Saldo;// - DescDiferimiento;
							//								ABDiferimiento = ABDiferimiento + Saldo;								
							//								if (iDescDif == 0) Saldo = 0.00m;
							//							}
							#endregion Diferimiento Comentado

							#region Gestión de cobranza
							if (Math.Round(Gasto, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Gasto, 2))
							{								
								Abono = Abono + Gasto;
								AbGastos = AbGastos + Gasto;
								Saldo = Saldo - Gasto;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Gasto, 2) > 0.00m && Math.Round(Gasto, 2) >= Math.Round(Saldo, 2)) 
							{
								Gasto = Gasto - Saldo;
								Abono = Abono + Saldo;
								AbGastos = AbGastos + Saldo;
								Saldo = 0.00m;
							}
							#endregion Gestión de cobranza

							#region Interés por mora
							if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
							{
								Abono = Abono + IntMora;
								AbIntMora = AbIntMora + IntMora;
								Saldo = Saldo - IntMora;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
							{
								IntMora = IntMora - Saldo;
								Abono = Abono + Saldo;
								AbIntMora = AbIntMora + Saldo;
								Saldo = 0.00m;
							}
							#endregion Interés por mora
						
							#region Interés por financiamiento
							if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
							{
								Abono = Abono + IntFin;
								AbInteres = AbInteres + IntFin;
								Saldo = Saldo - IntFin;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
							{
								IntFin = IntFin - Saldo;
								Abono = Abono + Saldo;
								AbInteres = AbInteres + Saldo;
								Saldo = 0.00m;
							}
							#endregion Interés por financiamiento

							#region Capital
							if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
							{
								Abono = Abono + Capital;
								AbCapital = AbCapital + Capital;
								Saldo = Saldo - Capital;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
							{
								Capital = Capital - Saldo;
								Abono = Abono + Saldo;
								AbCapital = AbCapital + Saldo;
								Saldo = 0.00m;								
							}
							#endregion Capital
							
							#region Actualiza tabla Cre_TablaDeAmortización
							if (Abono != 0.00m)
							{
								dtFecha = Convert.ToDateTime(this.dtFecha.Value); 
								decimal dSaldoPrb = 0;
								dSaldoPrb = Convert.ToDecimal(dr.Cells["Saldo"].Value);
								NSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value) - Abono - ABDiferimiento;

								sSQL = string.Format("Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}", 
									(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value, dtFecha.ToString("yyyyMMdd"), 
									(int) Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value, AbInteres, AbCapital, AbGastos, AbIntMora, ABDiferimiento, DescDiferimiento);
								Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							
								sConcepto = "";								

								id_Anticipo = Convert.ToInt32(Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value);

								if (Math.Round(NSaldo, 2) <= 0.00m) 
								{
									Estado = 2;
									EstadoSorteo = 1;
									sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();		
									#region sorteo_botonanticipos
//									int num_cuota;
//									num_cuota  = Convert.ToInt32(dr.Cells["NumeroCuota"].Value.ToString());
//
//									if ((int) Cruces.uGrid.ActiveRow.Cells["Tipo"].Value == 1)
//									{																						
//										if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idMotivo From Compra Where idCompra = {0}", Convert.ToInt32(Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value))) == 33)
//										{
//											if (Convert.ToDateTime(this.dtFecha.Value) <= Convert.ToDateTime(dr.Cells["Vence"].Value )) 
//											{
//												int j=1;
//												for(j=1;j<=3;j++)
//												{
//													string sSQLNum = string.Format("Exec NumeracionLocales 68, 0, 0");
//													string num_sorteo = Funcion.sEscalarSQL(cdsSeteoF,sSQLNum);
//										
//													Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
//												
//													string datosSQL = string.Format("Insert into sorteo_ticket (numero,id_compra,id_anticipo,estado,num_cuota,bodega,Valor_cuota)values ('{0}',{1},{2},1,{3},{4},{5})",num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idanticipo2,Convert.ToInt32( dr.Cells["NumeroCuota"].Value),(int)this.cmbBodega.Value,Val_Cuota);												  
//													Funcion.EjecutaSQL(cdsSeteoF, datosSQL);
//												}
//											}
//											else if (Convert.ToDateTime(this.dtFecha.Value) > Convert.ToDateTime(dr.Cells["Vence"].Value)) 
//											{
//												string sSQLNum = string.Format("Exec NumeracionLocales 68, 0, 0");
//												string num_sorteo =Funcion.sEscalarSQL(cdsSeteoF,sSQLNum);
//										
//												Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
//												
//												string datosSQL = string.Format("Insert into sorteo_ticket (numero,id_compra,id_anticipo,estado,num_cuota,bodega,Valor_cuota)values ('{0}',{1},{2},1,{3},{4},{5})",num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idanticipo2,Convert.ToInt32( dr.Cells["NumeroCuota"].Value),(int)this.cmbBodega.Value,Val_Cuota);
//												  
//												Funcion.EjecutaSQL(cdsSeteoF, datosSQL);
//											}
//										}
//									}
									#endregion sorteo_botonanticipos

								}
								else if (Math.Round(NSaldo, 2) > 0.00m) 
								{
									Estado = 3;
									sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
								}
							
								DateTime dtFechaCruce = (DateTime)this.dtFecha.Value;
									
								decimal dDebitado = AbIntMora + AbInteres + AbCapital + ABDiferimiento;

								int iPago = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
								
								#region  Descuento COVID
								decimal dDescGestion = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL (DescGestion ,0 )  From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
								decimal dDesMora = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL (DesMora ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
								#endregion  Descuento COVID 

								decimal dRescate = 0.00m;

								string sSQLRescate = string.Format("Exec Cre_ValidaSaldoCuotaRescatada {0}, {1}, '{2}'", 
									(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, dDebitado, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								dRescate = Funcion.decEscalarSQL(cdsSeteoF, sSQLRescate);
								
								string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}, {17}, {18}, {19}",
									(int)Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, 
									AbGastos, AbIntMora, AbInteres, AbCapital, Abono, Math.Round(NSaldo, 2), Estado, sConcepto, 
									Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iPago, dDebitado, this.GridFacturas.ActiveRow.Cells["NumeroDeOperacion"].Value.ToString(), (int) Cruces.uGrid.ActiveRow.Cells["Tipo"].Value, dDescGestion, dDesMora, ABDiferimiento, DescDiferimiento, dRescate);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLPCC);

								if ((int) Cruces.uGrid.ActiveRow.Cells["Tipo"].Value == 0)
								{
									#region Actualiza Pago en gestiones
									decimal dValorPagadoGestiones = AbCapital + AbInteres;
									string sSQLAPG = string.Format("Exec Cre_ActualizaValorPagoGestiones {0}, '{1}', '{2}', {3}, {4}",
										(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(dr.Cells["Vence"].Value).ToString("yyyyMMdd"), dValorPagadoGestiones, Abono);
									Funcion.EjecutaSQL(cdsSeteoF, sSQLAPG, true);
									#endregion Actualiza Pago en gestiones
								}
							}
							#endregion Actualiza tabla Cre_TablaDeAmortización
						}
					}	
					#endregion Calculo de Cobranza
						
					#region Asiento
					bool bPichincha = false;
					//					if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) bPichincha = true;

					string sSQLAsiento = string.Format("Exec AsientoCruceAnticipoNCCrediPoint {0}, 0, {1}, 0, {2}, '{3}', {4}", 
						(int) Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value, (int) Cruces.uGrid.ActiveRow.Cells["Tipo"].Value, bPichincha, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
					#endregion Asiento

					#region Actualiza Saldo Estado En Documentos
					string sSQLSaldos = string.Format("Exec CobranzasActualizaSaldoCruceCuentas {0}, {1}, {2}", 
						(int)Cruces.uGrid.ActiveRow.Cells["Tipo"].Value, (int)Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldos);
					#endregion Actualiza Saldo Estado En Documentos 

					#region Depositos pendientes 
					string sSQDepositosPendientes = string.Format("Exec ActualizaEstadoDepositosPendientes {0}, {1}",(int)Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQDepositosPendientes);
					#endregion Depositos pendientes 

					#region Actualiza Compromisos de Pago - Estado de Cuenta - Historial de autorizaciones
					string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}", 
						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, dtFecha.ToString("yyyyMMdd"), Convert.ToDecimal(Cruces.uGrid.ActiveRow.Cells["Abono"].Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActCompromisoPago, true);
					
					string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAEC, true);

					#region Actualiza Valor Cobrado Gestiones de Credito_Cobranza
							 
//					string sSQLActCompromisoCredito_Cobranza= string.Format("Exec Cbo_ActualizaCompromisoPago {0}, '{1}', {2}", 
//						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(dtFecha).ToString("yyyyMMdd"),  Convert.ToDecimal(Cruces.uGrid.ActiveRow.Cells["Abono"].Value));
//					Funcion.EjecutaSQL(cdsSeteoF, sSQLActCompromisoCredito_Cobranza, true);

					#endregion Actualiza Valor Cobrado Gestiones de Credito_Cobranza

					if ((int)this.txtIdAutorizacion.Value > 0) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CobranzaActualizaEstadoAutorizaciones {0}", this.txtIdAutorizacion.Value));						
					#endregion Actualiza Compromisos de Pago - Estado de Cuenta - Historial de autorizaciones				


					Cobranza();

					this.Cursor = Cursors.Default;
				}
			}
			#endregion Registro de Cobro
			
			this.txtCobro.Enabled = false;
			bMayorCobro = false;
			dValorCobro = 0.00m;

			this.txtIdAutorizacion.Value = 0;
			this.txtValorAutorizado.Value = 0.00m;
			this.txtEstadoAutorizacion.Value = 0;
//
//			#region Impresion Sorteo
//			ParameterFields paramFields = new ParameterFields();
//
//			ParameterField pfIdOrigen2 = new ParameterField();
//			ParameterDiscreteValue ValIdOrigen2 = new ParameterDiscreteValue();
//			pfIdOrigen2.ParameterFieldName = "@idAnticipo";
//			ValIdOrigen2.Value = idCompraPago;
//			pfIdOrigen2.CurrentValues.Add(ValIdOrigen2);
//			paramFields.Add(pfIdOrigen2);
//			if (Estado == 3 )
//			{
//				return;
//			}
//			if (EstadoSorteo == 1 )
//			{ 
//								
//				Reporte miRep = new Reporte("sorteo_ticket1.rpt", "");
//				miRep.MdiParent = this.MdiParent;
//				miRep.Titulo("Cupones de  Sorteo");
//				miRep.crVista.ParameterFieldInfo = paramFields;
//				miRep.Show();	
//				EstadoSorteo = 0;
//			}
//							
//			#endregion Impresion Sorteo
		}

		private void btnReliquidar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!this.ValidaCobro()) return;

			if (Convert.ToDateTime(this.dtFecha.Value) < dtUltimaFechaPago)
			{
				MessageBox.Show(string.Format("La Fecha de Cobro no puede ser menor a '{0}'", dtUltimaFechaPago.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.dtFecha.Focus();				
			}

			if ((bool)this.GridFacturas.ActiveRow.Cells["Reliquido"].Value)
			{
				MessageBox.Show("El Documento esta reliquidado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}			

			#region Valida Ultima Cuota
			int iCuotasPendientes = 0;
			decimal dSaldoCI = 0.00m;
			decimal dValorCuota = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)				
				if ((int)dr.Cells["Estado"].Value != 2) iCuotasPendientes++;

			if (iCuotasPendientes == 1 && iDiasVencidos >= 0)
			{
				MessageBox.Show("No puede reliquidar ultima cuota con saldo vencido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}

			if (iCuotasPendientes == 1)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)				
				{
					if ((int)dr.Cells["Estado"].Value != 2)
					{
						dSaldoCI = Convert.ToDecimal(dr.Cells["Interes"].Value) + Convert.ToDecimal(dr.Cells["Capital"].Value);
						dValorCuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
					}
				}
						
				if (dSaldoCI < dValorCuota)
				{
					MessageBox.Show("No puede reliquidar el saldo es menor al valor de la cuota", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtIdentificacion.Focus();
					return;
				}
			}
			#endregion Valida Ultima Cuota

			int iContDocPend = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDocumentosPendientes {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
			if (iContDocPend > 0 && !miAcceso.BReliquidarCruzandoCuentas)
			{
				MessageBox.Show("El Documento tiene saldos por cruzar\n\nComuniquese con el Departamento de Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}

			if (!Validacion.vbCarteraEnTransito((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value, 
				(int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value, (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value)) return;

			#region Comentado

			//			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && (bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && (int)this.txtDiasVencidos.Value > 103)
			//			{
			//				string sMensaje = "Documento en Fondo de Garantia.\n\nProcese el debito del Banco para re liquidar el documento";
			//				if (!miAcceso.BReliquidarCruzandoCuentas) sMensaje = "Para re liquidar el documento comuniquese con el Departamento de Cobranzas";
			//
			//				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.txtIdentificacion.Focus();			
			//				return;
			//			}

			

			//			if (Convert.ToDecimal(this.txtSaldoAFavor.Value) > 0.00m) 
			//				if (!Validacion.vbComparaDecimales(Convert.ToDecimal(this.txtProntoPago.Value), Convert.ToDecimal(this.txtSaldoAFavor.Value), "El cliente tiene saldo por cruzar\n\nComuniquese con el Departamento de Cobranzas", false)) return;
			#endregion Comentado
					
			if (!ValidaSaldosRefinanciacion()) return;

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCobroManualCOVID {0} ",(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value ), true) > 0)
			{
				MessageBox.Show("Documento en proceso de diferimiento, Ingrese el valor de la re liquidacion desde el boton de Cobro, en breve se cerrara la deuda del cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}
			#endregion Validacion

			if (DialogResult.No == MessageBox.Show("Esta Seguro De\n\n\n R E L I Q U I D A R?.\n\n\nEste Proceso No Puede Ser Revertido", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				this.txtIdentificacion.Focus();
				return;
			}

			if (iContDocPend == 0)
			{
				#region Registro de Cobro
				int idCompraPago = 0;

				using (VentaFormaDePago Caja = new VentaFormaDePago (idCompraPago, 1, Convert.ToDecimal(this.txtProntoPago.Value), 3, (DateTime)this.dtFecha.Value, (int)this.cmbBodega.Value, miAcceso.BFiniquitos))
				{
					if (DialogResult.OK == Caja.ShowDialog())
					{
						string sSQLRel = string.Format("Exec Reliquidacion {0}, '{1}', {2}, {3}", 
							(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 1, dInteres);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLRel);
					
						this.ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);

						this.Cursor = Cursors.WaitCursor;

						using (SqlConnection oConCobranzas = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
						{
							oConCobranzas.Open();

							SqlTransaction oTransaction = null;
							SqlCommand oCmdActualiza = oConCobranzas.CreateCommand();
							oCmdActualiza.Connection = oConCobranzas;
							oCmdActualiza.CommandType = CommandType.Text;
							oCmdActualiza.CommandTimeout = 120;
						
							try
							{							
								oTransaction = oConCobranzas.BeginTransaction();
								oCmdActualiza.Transaction = oTransaction;
						
								#region Variables
								decimal Cobro = Convert.ToDecimal(this.txtProntoPago.Value);
								decimal Saldo = Convert.ToDecimal(this.txtProntoPago.Value);
								decimal Diferimiento = 0.00m;
								decimal ABDiferimiento = 0.00m;
								decimal DescDiferimiento = 0.00m;	
								decimal Gasto = 0.00m;
								decimal IntMora = 0.00m;
								decimal IntFin = 0.00m;
								decimal Capital = 0.00m;
								decimal Abono = 0.00m;							
								decimal NSaldo = 0.00m;							
								decimal AbGastos = 0.00m;
								decimal AbIntMora = 0.00m;
								decimal AbInteres = 0.00m;
								decimal AbCapital = 0.00m;
								int Estado = 0;
								string sConcepto = "";
								int idCompraOrigen = (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value;

								int iDescDif = 1;
								//Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Temporal_Operaciones_Dif_Final Where idCompra = {0} And Estado = 3 And EstadoPoint = 3", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
								#endregion Variables
														
								#region Genera Abono Factura 24
								string sSQLNumero = string.Format("Exec NumeracionLocales 24, {0}, 0", (int)this.cmbBodega.Value);
								oCmdActualiza.CommandText = sSQLNumero;
								string sNumero = oCmdActualiza.ExecuteScalar().ToString(); 

								string sSQLPago = string.Format("Exec GuardaPagoCobranza {0}, {1}, {2}, 0, '{3}', '{4}'", 
									idCompraOrigen, (int)this.cmbBodega.Value, Math.Round(Convert.ToDecimal(this.txtProntoPago.Value), 2), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), sNumero);
								oCmdActualiza.CommandText = sSQLPago;								
								idCompraPago = (int)oCmdActualiza.ExecuteScalar(); 
								#endregion Genera Abono Factura 24

								#region Guarda Cobro de la factura	
								if (Caja.ultraGrid1.Rows.Count > 0)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Caja.ultraGrid1.Rows.All)
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
										if ((int)dr.Cells["idFormaPago"].Value == 3) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
										if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = Convert.ToDecimal(dr.Cells["ValorUsoTarjeta"].Value);

										string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
											idDetFormaPago, 
											idCompraPago, 3, 
											(int)dr.Cells["idFormaPago"].Value, Math.Round(Convert.ToDecimal(dr.Cells["Valor"].Value), 2),
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
								#endregion Guarda Cobro de la factura  										

								#region Calculo de Cobranza
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
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
										Gasto = Convert.ToDecimal(dr.Cells["Gastos"].Value);
										IntMora = Convert.ToDecimal(dr.Cells["Mora"].Value);
										IntFin = Convert.ToDecimal(dr.Cells["Interes"].Value);
										Capital = Convert.ToDecimal(dr.Cells["Capital"].Value);									
										Abono = 0.00m;
										
										#region Diferimiento
										if (Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2))
										{
											if (Convert.ToDateTime(dr.Cells["Vence"].Value) >= Convert.ToDateTime(this.dtFecha.Value) && (Capital + IntFin + IntMora + Gasto) <= Saldo)
											{
												ABDiferimiento = ABDiferimiento + Diferimiento;
												DescDiferimiento = DescDiferimiento + Diferimiento;											
											}
											else
											{											
												Abono = Abono + Diferimiento;
												ABDiferimiento = ABDiferimiento + Diferimiento;
												Saldo = Saldo - Diferimiento;
											}
										}									
										else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
										{
											Diferimiento = Diferimiento - Saldo;
											Abono = Abono + Saldo;
											ABDiferimiento = ABDiferimiento + Saldo;
											Saldo = 0.00m;
										}
										#endregion Diferimiento

										#region Diferimiento Comentado
										//										if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2) && Math.Round(Diferimiento, 2) != 0)
										//										{
										//											if (iDescDif > 0)
										//											{
										//												if (Diferimiento < 0.00m)
										//												{
										//													decimal dValorCuotas = Convert.ToDecimal(this.GridFacturas.ActiveRow.Cells["ValorCuotas"].Value);
										//													decimal dAbonoCD = Convert.ToDecimal(dr.Cells["Abono"].Value);										
										//													DescDiferimiento = IntFin +  Capital + dAbonoCD - Math.Abs(Diferimiento) - dValorCuotas;
										//												}
										//												else DescDiferimiento = Diferimiento;
										//											}
										//											//Diferimiento = Diferimiento - DescDiferimiento;
										//											Abono = Abono + Diferimiento; //- Math.Abs(DescDiferimiento);
										//											ABDiferimiento = ABDiferimiento + Diferimiento;								
										//											Saldo = Saldo - Diferimiento + DescDiferimiento;
										//										}
										//										else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
										//										{
										//											if (iDescDif > 0) DescDiferimiento = Saldo;
										//											Diferimiento = Diferimiento - Saldo;
										//											Abono = Abono + Saldo;// - DescDiferimiento;
										//											ABDiferimiento = ABDiferimiento + Saldo;								
										//											if (iDescDif == 0) Saldo = 0.00m;
										//										}
										#endregion Diferimiento Comentado

										#region Gestión de cobranza
										if (Math.Round(Gasto, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Gasto, 2))
										{								
											Abono = Abono + Gasto;
											AbGastos = AbGastos + Gasto;
											Saldo = Saldo - Gasto;
										}
										else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Gasto, 2) > 0.00m && Math.Round(Gasto, 2) >= Math.Round(Saldo, 2)) 
										{
											Gasto = Gasto - Saldo;
											Abono = Abono + Saldo;
											AbGastos = AbGastos + Saldo;
											Saldo = 0.00m;
										}
										#endregion Gestión de cobranza

										#region Interés por mora
										if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
										{
											Abono = Abono + IntMora;
											AbIntMora = AbIntMora + IntMora;
											Saldo = Saldo - IntMora;
										}
										else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
										{
											IntMora = IntMora - Saldo;
											Abono = Abono + Saldo;
											AbIntMora = AbIntMora + Saldo;
											Saldo = 0.00m;
										}
										#endregion Interés por mora
						
										#region Interés por financiamiento
										if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
										{
											Abono = Abono + IntFin;
											AbInteres = AbInteres + IntFin;
											Saldo = Saldo - IntFin;
										}
										else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
										{
											IntFin = IntFin - Saldo;
											Abono = Abono + Saldo;
											AbInteres = AbInteres + Saldo;
											Saldo = 0.00m;
										}
										#endregion Interés por financiamiento

										#region Capital
										if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
										{
											Abono = Abono + Capital;
											AbCapital = AbCapital + Capital;
											Saldo = Saldo - Capital;
										}
										else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
										{
											Capital = Capital - Saldo;
											Abono = Abono + Saldo;
											AbCapital = AbCapital + Saldo;
											Saldo = 0.00m;								
										}
										#endregion Capital
										
										#region Actualiza tabla Cre_TablaDeAmortización
										if (Abono > 0.00m)
										{										
											decimal dSaldoPrb = 0.00m;
											dSaldoPrb = Convert.ToDecimal(dr.Cells["Saldo"].Value);
											NSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value) - Abono - ABDiferimiento;

											string sSQLAC = string.Format("Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}", 
												(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idCompraPago, AbInteres, AbCapital, AbGastos, AbIntMora, ABDiferimiento, DescDiferimiento);
											oCmdActualiza.CommandText = sSQLAC;
											oCmdActualiza.ExecuteNonQuery();
							
											sConcepto = "";
											if (Math.Round(NSaldo, 2) == 0.00m) 
											{
												Estado = 2;
												sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
												int num_cuota;
												num_cuota  = Convert.ToInt32(dr.Cells["NumeroCuota"].Value.ToString());											
//												if (Estado ==2 )
//												{												
//													if (Convert.ToDateTime(this.dtFecha.Value) < Convert.ToDateTime(dr.Cells["Vence"].Value )) 
//													{
//														int j=1;		
//														for( j=1;j<=3;j++)
//														{
//															string ssSQL = string.Format("Exec NumeracionLocales 68, 0, 0");
//															oCmdActualiza.CommandText = ssSQL;
//															string num_sorteo  = oCmdActualiza.ExecuteScalar().ToString();	
//															Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
//															string datosSQL = string .Format("Insert into sorteo_ticket(numero,id_Compra,id_anticipo,Estado,bodega,num_cuota,Valor_Cuota)values ('{0}',{1},{2},1,{3},{4},{5})", num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,(int)this.cmbBodega.Value,num_cuota,Val_Cuota);
//															oCmdActualiza.CommandText = datosSQL;
//															oCmdActualiza.ExecuteNonQuery();
//														}
//													}
//													else 	if ( Convert.ToDateTime(this.dtFecha.Value) > Convert.ToDateTime(dr.Cells["Vence"].Value ))
//													{
//														string ssSQL = string.Format("Exec NumeracionLocales 68, 0, 0");
//														oCmdActualiza.CommandText = ssSQL;
//														string num_sorteo  = oCmdActualiza.ExecuteScalar().ToString();	
//														Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
//														string datosSQL = string .Format("Insert into sorteo_ticket(numero,id_Compra,id_anticipo,Estado,bodega,num_cuota,Valor_Cuota)values ('{0}',{1},{2},1,{3},{4},{5})", num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,(int)this.cmbBodega.Value,num_cuota,Val_Cuota);
//														oCmdActualiza.CommandText = datosSQL;
//														oCmdActualiza.ExecuteNonQuery();
//													}
//												}								
											}
											else if (Math.Round(NSaldo, 2) > 0.00m) 
											{
												Estado = 3;
												sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
											}
							
											string sSQLTipoPago = string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
											oCmdActualiza.CommandText = sSQLTipoPago;
											int iPago = (int)oCmdActualiza.ExecuteScalar();
											
											#region  Descuento COVID
											string sSQLDescGestion = string.Format("Select ISNULL (DescGestion ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
											oCmdActualiza.CommandText = sSQLDescGestion;
											decimal dDescGestion = (decimal)oCmdActualiza.ExecuteScalar();

											string sSQLDesMora = string.Format("Select ISNULL (DesMora ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
											oCmdActualiza.CommandText = sSQLDesMora;
											decimal  dDesMora = (decimal)oCmdActualiza.ExecuteScalar();

											//											decimal dDescuentoDiferimiento = 0.00m;
											//
											//											if (Diferimiento == 0.00m)
											//											{											
											//												oCmdActualiza.CommandText = string.Format("Select ISNULL (DescDiferimiento, 0) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
											//												dDescuentoDiferimiento = (decimal)oCmdActualiza.ExecuteScalar();										
											//											
											//												if (dDescuentoDiferimiento > 0.00m) ABDiferimiento = dDescuentoDiferimiento;
											//											}
											#endregion  Descuento COVID 

											decimal dDebitado = AbIntMora + AbInteres + AbCapital + ABDiferimiento;
											
											decimal dRescate = 0.00m;

											string sSQLRescate = string.Format("Exec Cre_ValidaSaldoCuotaRescatada {0}, {1}, '{2}'", 
													(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, dDebitado, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
												oCmdActualiza.CommandText = sSQLRescate;
												dRescate = (decimal)oCmdActualiza.ExecuteScalar();											
											
											string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}, {17}, {18}, {19}",
												idCompraPago, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, 
												AbGastos, AbIntMora, AbInteres, AbCapital, Abono, Math.Round(NSaldo, 2), Estado, sConcepto, 
												Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iPago, dDebitado, this.GridFacturas.ActiveRow.Cells["NumeroDeOperacion"].Value.ToString(), 0, dDescGestion, dDesMora, ABDiferimiento, DescDiferimiento, dRescate);
											oCmdActualiza.CommandText = sSQLPCC;
											oCmdActualiza.ExecuteNonQuery();

											#region Actualiza Pago en gestiones
											decimal dValorPagadoGestiones = AbCapital + AbInteres;
											string sSQLAPG = string.Format("Exec Cre_ActualizaValorPagoGestiones {0}, '{1}', '{2}', {3}, {4}",
												idCompraOrigen, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(dr.Cells["Vence"].Value).ToString("yyyyMMdd"), dValorPagadoGestiones, Abono);
											oCmdActualiza.CommandText = sSQLAPG;
											oCmdActualiza.ExecuteNonQuery();
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
									oCmdActualiza.CommandText = sSQLHA;
									oCmdActualiza.ExecuteNonQuery();
								}
								#endregion Historial de autorizaciones

								#region Asiento - Saldos
//								string sSQLAsiento = string.Format("Exec AsientoCobranzaCrediPoint {0}, 0, {1}", idCompraPago, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
//								oCmdActualiza.CommandText = sSQLAsiento;
//								oCmdActualiza.ExecuteNonQuery();

								string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 9 Exec ActualizaSaldoFacturas {1}, 1, 9", idCompraPago, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
								oCmdActualiza.CommandText = sSQLSaldo;
								oCmdActualiza.ExecuteNonQuery();
								#endregion Asiento - Saldos
															
								#region Saldo por pre cancelacion temprana
								//								oCmdActualiza.CommandText = string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", 
								//									Convert.ToDateTime(this.GridFacturas.ActiveRow.Cells["Fecha"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								//								int iDias = (int)oCmdActualiza.ExecuteScalar();
								//
								//								if (!(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value)
								//								{
								//									if (iDias <= 8)
								//									{
								//										oCmdActualiza.CommandText = "Exec NumeracionLocales 5, 1, 1";
								//										string sNumeroNC = oCmdActualiza.ExecuteScalar().ToString();
								//										
								//										//@idCompra Int, @Numero Varchar(10), @Valor Decimal (18, 2), @Fecha Datetime, @Autorizacion Varchar(60)
								//										oCmdActualiza.CommandText = string.Format("Exec CreaNotasDeCreditoReliquidacion {0}, '{1}', '{2}'", 
								//											(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, sNumeroNC, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
								//										int idDocumento = (int) oCmdActualiza.ExecuteScalar();
								//
								//										oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", idDocumento);;
								//										string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();
								//
								//										string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
								//										string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
								//										string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;
								//
								//										oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idDocumento);
								//										oCmdActualiza.ExecuteNonQuery();
								//									}
								//								}
								#endregion Saldo por pre cancelacion temprana

								#region Actualiza Compromisos de Pago - Estado de Cuenta 
								string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}", 
									(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDecimal(this.txtProntoPago.Value));
								oCmdActualiza.CommandText = sSQLActCompromisoPago;
								oCmdActualiza.ExecuteNonQuery();								

								string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
								oCmdActualiza.CommandText = sSQLAEC;
								oCmdActualiza.ExecuteNonQuery();
								#endregion Actualiza Compromisos de Pago - Estado de Cuenta

								#region Actualiza Valor Cobrado Gestiones de Credito_Cobranza
							 
//								string sSQLActCompromisoCredito_Cobranza= string.Format("Exec Cbo_ActualizaCompromisoPago {0}, '{1}', {2}", 
//									(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDecimal(this.txtCobro.Value));
//								oCmdActualiza.CommandText = sSQLActCompromisoCredito_Cobranza;
//								oCmdActualiza.ExecuteNonQuery();	

								#endregion Actualiza Valor Cobrado Gestiones de Credito_Cobranza

								#region Notificacion Pago
								string sSQLNotificaPago = string.Format("Exec GuardaNotifiacionCuotaPagos {0}, {1}, {2},{3}, {4}",(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,0,(int)this.cmbBodega.Value,1);
								oCmdActualiza.CommandText = sSQLNotificaPago;	
								oCmdActualiza.ExecuteNonQuery();													
								#endregion Notifiacion Pago

								oTransaction.Commit();	

								oConCobranzas.Close();

								this.Cobranza();
				
								#region Impresion
								using (frmNumeracionRecibo miMsje = new frmNumeracionRecibo (Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Secuencial, '') From Anticipos Where idAnticipo = {0}", idCompraPago))))
								{
									if (DialogResult.OK == miMsje.ShowDialog())
										Facturacion.ImpresionRecibos(idCompraPago, 2, cdsSeteoF);
								}

								#endregion Impresion					

								#region Impresion Sorteo COMENTADO
//								ParameterFields paramFields = new ParameterFields();
//
//								ParameterField pfIdOrigen2 = new ParameterField();
//								ParameterDiscreteValue ValIdOrigen2 = new ParameterDiscreteValue();
//								pfIdOrigen2.ParameterFieldName = "@idAnticipo";
//								ValIdOrigen2.Value = idCompraPago;
//								pfIdOrigen2.CurrentValues.Add(ValIdOrigen2);
//								paramFields.Add(pfIdOrigen2);
//
//								Reporte miRep = new Reporte("sorteo_ticket1.rpt", "");
//								miRep.MdiParent = this.MdiParent;
//								miRep.Titulo("Cupones de  Sorte");
//								miRep.crVista.ParameterFieldInfo = paramFields;
//								miRep.Show();
								#endregion Impresion Sorteo

								Cobranza();	

								this.Cursor = Cursors.Default;
							}
							catch (Exception ex)
							{
								oTransaction.Rollback();

								oConCobranzas.Close();		

								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);													
							}
							finally
							{
								oConCobranzas.Close();
							}
						}
					}
				}
				#endregion Registro de Cobro			
			}
			else if (miAcceso.BReliquidarCruzandoCuentas)
			{
				#region Cruce de cuentas
				decimal dGC = 0.00m;
				decimal dIM = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
				{
					dGC = dGC + Convert.ToDecimal(dr.Cells["Gastos"].Value);
					dIM = dIM + Convert.ToDecimal(dr.Cells["Mora"].Value);
				}

				using (frmReliquidacion Reliquidar = new frmReliquidacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, (DateTime)this.dtFecha.Value, dGC, dIM, this.txtIdentificacion.Text.ToString(), Convert.ToDateTime(this.GridFacturas.ActiveRow.Cells["Fecha"].Value)))
				{					
					if (DialogResult.OK == Reliquidar.ShowDialog())
					{						
						#region Variables
						decimal Gasto = 0.00m;
						decimal IntMora = 0.00m;
						decimal IntFin = 0.00m;
						decimal Capital = 0.00m;
						decimal Abono = 0.00m;				
						decimal NSaldo = 0.00m;				
						decimal Diferimiento = 0.00m;
						decimal ABDiferimiento = 0.00m;
						decimal DescDiferimiento = 0.00m;	
						decimal AbGastos = 0.00m;
						decimal AbIntMora = 0.00m;
						decimal AbInteres = 0.00m;
						decimal AbCapital = 0.00m;
								
						int Estado = 0;
						DateTime dtFecha = DateTime.Today;
						string sSQL = "";
						string sConcepto = "";

						int iDescDif = 1; 
						//Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Temporal_Operaciones_Dif_Final Where idCompra = {0} And Estado = 3 And EstadoPoint = 3", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
						#endregion Variables

						string sSQLRel = string.Format("Exec Reliquidacion {0}, '{1}', {2}, {3}", 
							(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 1, Reliquidar.dInteres);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLRel);

						this.ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);

						this.Cursor = Cursors.WaitCursor;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drReliquidar in Reliquidar.uGrid.Rows.All)
						{
							this.ConsultaTablaDeAmortizacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);

							decimal Cobro = Convert.ToDecimal(drReliquidar.Cells["Saldo"].Value);
							decimal Saldo = Convert.ToDecimal(drReliquidar.Cells["Saldo"].Value);

							#region Calculo de Cobranza
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
							{
								if (Saldo > 0.00m)
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
									Gasto = Convert.ToDecimal(dr.Cells["Gastos"].Value);
									IntMora = Convert.ToDecimal(dr.Cells["Mora"].Value);
									IntFin = Convert.ToDecimal(dr.Cells["Interes"].Value);
									Capital = Convert.ToDecimal(dr.Cells["Capital"].Value);							
									Abono = 0.00m;
									
									#region Diferimiento
									if (Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2))
									{
										if (Convert.ToDateTime(dr.Cells["Vence"].Value) >= Convert.ToDateTime(this.dtFecha.Value) && (Capital + IntFin + IntMora + Gasto) <= Saldo)
										{
											ABDiferimiento = ABDiferimiento + Diferimiento;
											DescDiferimiento = DescDiferimiento + Diferimiento;											
										}
										else
										{											
											Abono = Abono + Diferimiento;
											ABDiferimiento = ABDiferimiento + Diferimiento;
											Saldo = Saldo - Diferimiento;
										}
									}									
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
									{
										Diferimiento = Diferimiento - Saldo;
										Abono = Abono + Saldo;
										ABDiferimiento = ABDiferimiento + Saldo;
										Saldo = 0.00m;
									}
									#endregion Diferimiento

									#region Diferimiento Comentado
									//									if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Diferimiento, 2) && Math.Round(Diferimiento, 2) != 0)
									//									{
									//										if (iDescDif > 0)
									//										{
									//											if (Diferimiento < 0.00m)
									//											{
									//												decimal dValorCuotas = Convert.ToDecimal(this.GridFacturas.ActiveRow.Cells["ValorCuotas"].Value);
									//												decimal dAbonoCD = Convert.ToDecimal(dr.Cells["Abono"].Value);										
									//												DescDiferimiento = IntFin +  Capital + dAbonoCD - Math.Abs(Diferimiento) - dValorCuotas;
									//											}
									//											else DescDiferimiento = Diferimiento;
									//										}
									//										//Diferimiento = Diferimiento - DescDiferimiento;
									//										Abono = Abono + Diferimiento; //- Math.Abs(DescDiferimiento);
									//										ABDiferimiento = ABDiferimiento + Diferimiento;								
									//										Saldo = Saldo - Diferimiento + DescDiferimiento;
									//									}
									//									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Diferimiento, 2) >= Math.Round(Saldo, 2)) 
									//									{
									//										if (iDescDif > 0) DescDiferimiento = Saldo;
									//										Diferimiento = Diferimiento - Saldo;
									//										Abono = Abono + Saldo;// - DescDiferimiento;
									//										ABDiferimiento = ABDiferimiento + Saldo;								
									//										if (iDescDif == 0) Saldo = 0.00m;
									//									}
									#endregion Diferimiento Comentado

									#region Gestión de cobranza
									if (Math.Round(Gasto, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Gasto, 2))
									{								
										Abono = Abono + Gasto;
										AbGastos = AbGastos + Gasto;
										Saldo = Saldo - Gasto;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Gasto, 2) > 0.00m && Math.Round(Gasto, 2) >= Math.Round(Saldo, 2)) 
									{
										Gasto = Gasto - Saldo;
										Abono = Abono + Saldo;
										AbGastos = AbGastos + Saldo;
										Saldo = 0.00m;
									}
									#endregion Gestión de cobranza

									#region Interés por mora
									if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
									{
										Abono = Abono + IntMora;
										AbIntMora = AbIntMora + IntMora;
										Saldo = Saldo - IntMora;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
									{
										IntMora = IntMora - Saldo;
										Abono = Abono + Saldo;
										AbIntMora = AbIntMora + Saldo;
										Saldo = 0.00m;
									}
									#endregion Interés por mora
						
									#region Interés por financiamiento
									if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
									{
										Abono = Abono + IntFin;
										AbInteres = AbInteres + IntFin;
										Saldo = Saldo - IntFin;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
									{
										IntFin = IntFin - Saldo;
										Abono = Abono + Saldo;
										AbInteres = AbInteres + Saldo;
										Saldo = 0.00m;
									}
									#endregion Interés por financiamiento

									#region Capital
									if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
									{
										Abono = Abono + Capital;
										AbCapital = AbCapital + Capital;
										Saldo = Saldo - Capital;
									}
									else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
									{
										Capital = Capital - Saldo;
										Abono = Abono + Saldo;
										AbCapital = AbCapital + Saldo;
										Saldo = 0.00m;								
									}
									#endregion Capital
									
									#region Actualiza tabla Cre_TablaDeAmortización
									if (Abono != 0)
									{
										dtFecha = Convert.ToDateTime(this.dtFecha.Value);
										decimal dSaldoPrb = 0;
										dSaldoPrb = Convert.ToDecimal(dr.Cells["Saldo"].Value);
										NSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value) - Abono;

										sSQL = string.Format("Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}", 
											(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value, 
											Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int) drReliquidar.Cells["idAnticipo"].Value, AbInteres, AbCapital, AbGastos, AbIntMora, ABDiferimiento, DescDiferimiento);
										Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
							
										#region Estado Cuota
										sConcepto = "";
										if (Math.Round(NSaldo, 2) == 0.00m) 
										{
											Estado = 2;
											sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();

//											#region sorteo_botonreliquidar
//											int num_cuota;
//											num_cuota  = Convert.ToInt32(dr.Cells["NumeroCuota"].Value.ToString());
//											if (Estado ==2 )
//											{																																						
//												if (Convert.ToDateTime(this.dtFecha.Value) < Convert.ToDateTime(dr.Cells["Vence"].Value )) 
//												{
//													int j=1;
//													for( j=1;j<=3;j++)
//													{
//														string sSQLNum = string.Format("Exec NumeracionLocales 68, 0, 0");
//														string num_sorteo =Funcion.sEscalarSQL(cdsSeteoF,sSQLNum);
//										
//														Val_Cuota = Convert.ToDecimal(dr.Cells["ValorCuota"].Value);
//														//idanticipo2 = Convert.ToInt32((int)Cruces.uGrid.ActiveRow.Cells["idAnticipo"].Value);
//
//
//														string datosSQL = string.Format("Insert into sorteo_ticket (numero,id_compra,id_anticipo,estado,num_cuota,bodega,Valor_cuota)values ('{0}',{1},{2},1,{3},{4},{5})",num_sorteo,(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idanticipo2,Convert.ToInt32( dr.Cells["NumeroCuota"].Value),(int)this.cmbBodega.Value,Val_Cuota);
//												  
//														Funcion.EjecutaSQL(cdsSeteoF, datosSQL);
//													}
//
//
//												}
//
//											}
//											//AY3
//
//											#endregion sorteo_botonreliquidar

										}
										else if (Math.Round(NSaldo, 2) > 0.00m) 
										{
											Estado = 3;
											sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
										}
										#endregion Estado Cuota

										DateTime dtFechaCruce = (DateTime)this.dtFecha.Value;
									
										decimal dDebitado = AbIntMora + AbInteres + AbCapital + ABDiferimiento;

										int iPago = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
										
										#region  Descuento COVID
										decimal dDescGestion = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL (DescGestion ,0 )  From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
										decimal dDesMora = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL (DesMora ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
										decimal dDescuentoDiferimiento = 0.00m;

										if (Diferimiento == 0.00m)
										{	
											dDescuentoDiferimiento = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL (DescDiferimiento, 0) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value));
											
											if (dDescuentoDiferimiento > 0.00m) ABDiferimiento = dDescuentoDiferimiento;
										}
										#endregion  Descuento COVID 

										decimal dRescate = 0.00m;
									
											string sSQLRescate = string.Format("Exec Cre_ValidaSaldoCuotaRescatada {0}, {1}, '{2}'", 
												(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, dDebitado, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
											dRescate = Funcion.decEscalarSQL(cdsSeteoF, sSQLRescate);
										
										string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}, {17}, {18}, {19}",
											(int)drReliquidar.Cells["idAnticipo"].Value, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, 
											AbGastos, AbIntMora, AbInteres, AbCapital, Abono, Math.Round(NSaldo, 2), Estado, sConcepto, 
											Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iPago, dDebitado, this.GridFacturas.ActiveRow.Cells["NumeroDeOperacion"].Value.ToString(), 
											(int)Reliquidar.uGrid.ActiveRow.Cells["Tipo"].Value, dDescGestion, dDesMora, ABDiferimiento, DescDiferimiento, dRescate);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLPCC);
									}
									#endregion Actualiza tabla Cre_TablaDeAmortización
								}
							}	
							#endregion Calculo de Cobranza
						
							#region Asiento
							bool bPichincha = false;
							if (((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 || 
								(int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 15 || 
								(int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 16) && 
								(int)this.GridFacturas.ActiveRow.Cells["NumeroLote"].Value > 0 && 
								(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) bPichincha = true;

							string sSQLAsiento = string.Format("Exec AsientoCruceAnticipoNCCrediPoint {0}, 0, {1}, 0, {2}, '{3}'", 
								(int) drReliquidar.Cells["idAnticipo"].Value, (int) drReliquidar.Cells["Tipo"].Value, bPichincha, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")); 					
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
							#endregion Asiento
						
							#region Actualiza Saldo Estado En Documentos
							string sSQLSaldos = string.Format("Exec CobranzasActualizaSaldoCruceCuentas {0}, {1}, {2}", 
								(int)drReliquidar.Cells["Tipo"].Value, (int)drReliquidar.Cells["idAnticipo"].Value, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldos);
							#endregion Actualiza Saldo Estado En Documentos
						}

						#region Actualiza Compromisos de Pago - Estado de Cuenta - Historial de autorizaciones
						string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}", 
							(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Reliquidar.dSaldo);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLActCompromisoPago, true);
					
						string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAEC, true);

					
						
						#region Actualiza Valor Cobrado Gestiones de Credito_Cobranza
							 
//						string sSQLActCompromisoCredito_Cobranza= string.Format("Exec Cbo_ActualizaCompromisoPago {0}, '{1}', {2}", 
//							(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Reliquidar.dSaldo);
//						Funcion.EjecutaSQL(cdsSeteoF, sSQLActCompromisoCredito_Cobranza, true);

						#endregion Actualiza Valor Cobrado Gestiones de Credito_Cobranza

						if ((int)this.txtIdAutorizacion.Value > 0) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CobranzaActualizaEstadoAutorizaciones {0}", this.txtIdAutorizacion.Value));						
						#endregion Actualiza Compromisos de Pago - Estado de Cuenta - Historial de autorizaciones		
		
						#region Notifiacion Pago
//						string sSQLNotificaPago = string.Format("Exec GuardaNotifiacionCuotaPagos {0}, {1}, {2},{3}, {4}",(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,idCompraPago,0,(int)this.cmbBodega.Value,1);
//						oCmdActualiza.CommandText = sSQLNotificaPago;	
//						oCmdActualiza.ExecuteNonQuery();														
						#endregion Notifiacion Pago

						this.Cobranza();
						this.Cursor = Cursors.Default;	
					}
				}
				#endregion Cruce de cuentas
			}
		}		
		
		private void ConsultaFacturas(string sIdentificacion)
		{
			string sSQL = string.Format("Exec ConsultaFacturasCobranza '{0}'", sIdentificacion);
			FuncionesProcedimientos.dsGeneral(sSQL, this.GridFacturas);
		}

		private void ConsultaTablaDeAmortizacion(int idCompra)
		{
			string sSQL = string.Format("Exec Cobranzas {0}, '{1}'", 
				idCompra, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.gridPagos);			
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 46, this.txtIdentificacion.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);			
			if (sParametros[3] != null) this.txtIdentificacion.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];			

			if ((int)this.txtIdCliente.Value > 0) 
				ConsultaFacturas(this.txtIdentificacion.Text.ToString());
		}
		bool bDescCovid = true;

		private void LimpiaControles(bool bCliente, bool bDocumentos)
		{
			#region Limpia Informacion
			if (bCliente)
			{				
				this.txtIdCliente.Value = 0;
				this.txtIdentificacion.Text = "";
				this.txtNombre.Text = "";				
				ConsultaFacturas("");
			}
			
			bMayorCobro = false;
			dValorCobro = 0.00m;
			iDiasVencidos = 0;
			dInteres = 0.00m;
			bDescCovid = true;
			dSaldoSD = 0.00m;

			this.txtCobro.Value = 0.00m;				
			this.txtProntoPago.Value = 0.00m;
			this.txtSaldoAFavor.Value = 0.00m;
			this.txtValorARefinanciar.Value = 0.00m;
			this.txtValorACondonar.Value = 0.00m;
			this.txtValorAutorizado.Value = 0.00m;

			this.txtSaldoVencido.Value = 0.00m;
			this.txtCuotaVencida.Value = 0.00m;
			this.txtDiasVencidos.Value = 0;
			this.txtSaldoTotal.Value = 0.00m;
			this.txtSiguienteCuota.Value = 0.00m;
			this.txtProximaVencer.Value = 0.00m;						
			this.txtIdAutorizacion.Value = 0;
			this.txtEstadoAutorizacion.Value = 0;
			
			this.txtSaldoVencido.Appearance.ForeColor = Color.Black;
			this.txtSaldoVencido.Appearance.ForeColorDisabled = Color.Black;
			this.txtCuotaVencida.Appearance.ForeColor = Color.Black;
			this.txtCuotaVencida.Appearance.ForeColorDisabled = Color.Black;
			this.txtDiasVencidos.Appearance.ForeColor = Color.Black;
			this.txtDiasVencidos.Appearance.ForeColorDisabled = Color.Black;
			#endregion Limpia Informacion

			this.txtCobro.Enabled = false;
			this.btCobro.Enabled = false;
			this.btnRefinanciacion.Enabled = false;
			this.mnuDocumentos.Enabled = false;
			this.mnuGenerarCXC.Enabled = false;

			this.txtValorACondonar.Enabled = false;
			this.btnCondonar.Enabled = false;

			if (bDocumentos)
				ConsultaTablaDeAmortizacion(0);
		}

		private void txtIdentificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{					
				LimpiaControles(true, true);

				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 46);
				if (sIdentificacion.Length > 0)
				{						
					this.txtIdentificacion.Text = sIdentificacion;
					
					CargaCliente();				
				}				
			}
		}
		
		private void txtIdentificacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);

			if (e.KeyChar == 13) 
			{				
				string sIdentificacion = this.txtIdentificacion.Text.ToString();
			
				LimpiaControles(true, true);
				
				this.txtIdentificacion.Text = sIdentificacion;
				//AbreactualizaClientes(sIdentificacion);
				CargaCliente();
			}
		}

		private void txtIdentificacion_EditorButtonClick(object sender, Infragistics.Win.UltraWinEditors.EditorButtonEventArgs e)
		{
			LimpiaControles(true, true);

			string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 46);

			if (sIdentificacion.Length > 0)
			{						
				this.txtIdentificacion.Text = sIdentificacion;				
				//AbreactualizaClientes(sIdentificacion);
				CargaCliente();
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 46, (DateTime)this.dtFecha.Value, false, cdsSeteoF);

			if (this.gridPagos.Rows.Count > 0) 
			{
				LimpiaControles(false, true);
			}			
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) if (this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.txtIdentificacion.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}
		
		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{	
			if (!this.ValidaCobro()) return;
			
			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 2)
			{
				MessageBox.Show("Cartera Solidario no aplica autorización de cobros", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}

			if ((bool)this.GridFacturas.ActiveRow.Cells["Reliquido"].Value)
			{
				MessageBox.Show("El Documento esta reliquidado\n\nComuniquese con el Departamento de Cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}

			if (!Validacion.vbCampoDecimalVacio(this.txtValorAutorizado, "El valor autorizado no puede ser Cero", 0.01m, Convert.ToDecimal(this.txtSaldoTotal.Value))) return;

			if (Convert.ToDecimal(this.txtSaldoVencido.Value) == 0.00m)
			{
				MessageBox.Show("El saldo vencido es 0.00", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}

			if (!Validacion.vbComparaDecimales(Convert.ToDecimal(this.txtValorAutorizado.Value), Convert.ToDecimal(this.txtSaldoVencido.Value), "El valor a autorizar no puede ser mayor o igual al saldo vencido del documento", true)) return;
			
			if (!Validacion.vbComparaDecimales(Convert.ToDecimal(this.txtValorAutorizado.Value), Convert.ToDecimal(this.txtSaldoTotal.Value), "El valor a autorizar no puede ser mayor o igual al saldo del documento", true)) return;			
			
			if (!this.ValidaSaldosAFavor()) return;

			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value == 4 && 
				(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value && (int)this.txtDiasVencidos.Value > 97)
			{
				string sMensaje = "Documento en Fondo de Garantia.\n\nProcese el debito del Banco para autorizar pagos";
				if (!miAcceso.BReliquidarCruzandoCuentas) sMensaje = "Para autorizar pagos comuniquese con el Departamento de Cobranzas";

				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtIdentificacion.Focus();			
				return;
			}

			using (frmNotasSistema Nota = new frmNotasSistema())
			{
				if (DialogResult.OK == Nota.ShowDialog())
				{
					DateTime dFecUltPag = DateTime.Today;

					string sSQL = string.Format("Exec GrabaHistAutCobro {0}, {1}, {2}, {3}, {4}, {5}, '{6}', '{7}'", (int)this.txtIdAutorizacion.Value, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value,
						Convert.ToDecimal(this.txtValorAutorizado.Value), Convert.ToDecimal(this.txtCobro.Value), 
						(int)this.txtDiasVencidos.Value, this.txtCuotaVencida.Value, 
						dFecUltPag.ToString("yyyyMMdd"), Nota.txtNotas.Text.ToString());
					this.txtIdAutorizacion.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					MessageBox.Show(string.Format("Pago Autorizado {0}", Convert.ToDecimal(this.txtValorAutorizado.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					Cobranza();	
				}
			}			
		}

		private void txtCobro_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void txtCobro_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btCobro_Click(sender, e);
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void txtValorAutorizado_Click(object sender, System.EventArgs e)
		{
			this.txtValorAutorizado.SelectAll();
		}

		private void txtValorAutorizado_Enter(object sender, System.EventArgs e)
		{
			this.txtValorAutorizado.SelectAll();
		}

		private void txtValorAutorizado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAutorizar_Click(sender, e);
		}

		private void txtValorAutorizado_Validated(object sender, System.EventArgs e)
		{
			if (this.txtValorAutorizado.Value == System.DBNull.Value) this.txtValorAutorizado.Value = 0.00m;
		}



		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtIdentificacion.Focus();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnHistorico_Click(object sender, System.EventArgs e)
		{
			if (GridFacturas.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.txtIdentificacion.Focus();
				return;
			}
			if (GridFacturas.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);        				
				this.txtIdentificacion.Focus();
				return;
			}

			frmEstadoCuentaPorOperacion ECPO = new frmEstadoCuentaPorOperacion((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
			ECPO.ShowDialog();
		}



		private void txtIdentificacion_Validated(object sender, System.EventArgs e)
		{
			string sIdentificacion = this.txtIdentificacion.Text.ToString();

			LimpiaControles(true, true);
	
			this.txtIdentificacion.Text = sIdentificacion;

			CargaCliente();	
//			using (frmGestionPorCedula Nota = new frmGestionPorCedula(this.txtIdentificacion.Text))
//			{
//				if (DialogResult.OK == Nota.ShowDialog())
//				{
//					AbreactualizaClientes(sIdentificacion);
//				}
//			}

		}
		public static void AbreactualizaClientes(string sIdentificacion)
		{
			using (frmActualizadatosCliente AFCL = new frmActualizadatosCliente(sIdentificacion))
			{
				AFCL.ShowDialog();
			}
		}


		private void mnuDocumentos_Click(object sender, System.EventArgs e)
		{
			if (!this.ValidaCobro()) return;

			if ((int)this.GridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value != 4)
			{
				MessageBox.Show("Refinanciación aplica para cartera Pichincha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (!(bool)this.GridFacturas.ActiveRow.Cells["Confirmado"].Value) 
			{
				MessageBox.Show("La operacion seleccionada no aplica para refinanciación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtCuotaVencida.Value < 3) 
			{
				MessageBox.Show("El minimo de cuotas vencidas para refinanciar debe ser 3", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Convert.ToDecimal(this.txtSaldoVencido.Value) <= 0) 
			{
				MessageBox.Show("La operacion seleccionada no tiene saldo vencido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			DateTime dtFechaGenera = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Select FechaVencimiento From Compra Where idCompra = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));

			dtFechaGenera = dtFechaGenera.AddMonths(1);

			frmDocsRefinanciacion Dctos = new frmDocsRefinanciacion ((int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Convert.ToDecimal(this.txtSaldoVencido.Value), (int)this.txtIdCliente.Value, dtFechaGenera, Convert.ToDateTime(this.dtFecha.Value));
			Dctos.ShowDialog();	

			this.ConsultaFacturas(this.txtIdentificacion.Text.ToString());
		}

		private void btnRefinanciacion_Click(object sender, System.EventArgs e)
		{
			this.mnuRefinanciacion.Show(this.btnRefinanciacion, new Point(1, 25));
		}

		private void mnuGenerarCXC_Click(object sender, System.EventArgs e)
		{
			if (!this.ValidaCobro()) return;

			if ((int)this.GridFacturas.ActiveRow.Cells["ER"].Value == 0)
			{
				MessageBox.Show("No se ha generado los documentos de refinanciación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.GridFacturas.ActiveRow.Cells["ER"].Value == 2)
			{
				MessageBox.Show("Ya se generarón los documentos de refinanciación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			decimal dGC = 0.00m;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
				dGC = dGC + Convert.ToDecimal(dr.Cells["Gastos"].Value);

			decimal dValorNC = Convert.ToDecimal(this.txtSaldoVencido.Value) - dGC;
			Cursor = Cursors.WaitCursor;

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

					#region Numeracion
					oCmdActualiza.CommandText = "Exec NumeracionLocales 27, 1, 0";
					string sNumeroND = oCmdActualiza.ExecuteScalar().ToString();

					oCmdActualiza.CommandText = "Exec NumeracionLocales 5, 1, 0";
					string sNumeroNC = oCmdActualiza.ExecuteScalar().ToString();					
					#endregion Numeracion
			
					string sSQL = string.Format("Exec CobranzasCreaDocsRefinanciacion {0}, '{1}', '{2}', {3}", 
						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, sNumeroND, sNumeroNC, dValorNC);					
					oCmdActualiza.CommandText = sSQL;
					oCmdActualiza.ExecuteNonQuery();

					oCmdActualiza.CommandText = string.Format("Exec CobranzaActualizaEstadoRefinanciacion {0}, 0, 0, 2, '{1}', ''", 
						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, DateTime.Today.ToString("yyyyMMdd"));
					oCmdActualiza.ExecuteNonQuery();

					oCmdActualiza.CommandText = string.Format("Select idCompra From Compra Where idTipoFactura = 27 And idOrigen = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					int idCompraND = (int)oCmdActualiza.ExecuteScalar();

					oCmdActualiza.CommandText = string.Format("Select idCompra From Compra Where idTipoFactura = 5 And idOrigen = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					int idCompraNC = (int)oCmdActualiza.ExecuteScalar(); 

					oTransaction.Commit();

					MessageBox.Show("Documentos de Refinanciación Generados Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Impresión
					string stFiltro = "{Compra.idCompra} = " + idCompraND.ToString();
								
					Reporte ReporteND = new Reporte("NotaDebitoVenta.Rpt", stFiltro);
					ReporteND.ShowDialog();															
								
					stFiltro = "{Compra.idCompra} = " + idCompraNC.ToString();
								
					Reporte ReporteNC = new Reporte("DevolucionVenta.Rpt", stFiltro);
					ReporteNC.ShowDialog();
					#endregion Impresión

					Cursor = Cursors.Default;
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

					Cursor = Cursors.Default;
				}
				finally
				{
					oConexion.Close();
				}
			}			
		}

		private void GridFacturas_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if ((int)e.Row.Cells["ER"].Value == 1)
			{
				e.Row.CellAppearance.BackColor = Color.YellowGreen;
				e.Row.CellAppearance.BackColorDisabled = Color.YellowGreen;
			}
			if ((int)e.Row.Cells["ER"].Value == 2)
			{
				e.Row.CellAppearance.BackColor = Color.Yellow;
				e.Row.CellAppearance.BackColorDisabled = Color.Yellow;
			}
		}

		private void btnCondonar_Click(object sender, System.EventArgs e)
		{
			if (!this.ValidaCobro()) return;

			if (!(bool)this.GridFacturas.ActiveRow.Cells["Asignado"].Value)
			{
				MessageBox.Show("Condonación se aplica a operaciones asignadas a Abogados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			using (SqlConnection oConCobranzas = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConCobranzas.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConCobranzas.CreateCommand();
				oCmdActualiza.Connection = oConCobranzas;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{							
					oTransaction = oConCobranzas.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Variables
					decimal Cobro = Convert.ToDecimal(this.txtValorACondonar.Value);
					decimal Gasto = 0.00m;
					decimal IntMora = 0.00m;
					decimal IntFin = 0.00m;
					decimal Capital = 0.00m;
					decimal Abono = 0.00m;
					decimal Saldo = Convert.ToDecimal(this.txtValorACondonar.Value);
					decimal NSaldo = 0.00m;					
					decimal AbGastos = 0.00m;
					decimal AbIntMora = 0.00m;
					decimal AbInteres = 0.00m;
					decimal AbCapital = 0.00m;
					int Estado = 0;
					string sConcepto = "";					
					#endregion Variables
														
					#region Genera Nota De Credito
					oCmdActualiza.CommandText = "Exec NumeracionLocales 5, 57, 0";
					string sNumero = oCmdActualiza.ExecuteScalar().ToString();

					string sSQLNC = string.Format("Exec CobranzasCreaNCCondonacion {0}, '{1}', {2}, '{3}'", 
						(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, sNumero, Math.Round(Convert.ToDecimal(this.txtValorACondonar.Value), 2), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
					oCmdActualiza.CommandText = sSQLNC;
					int idCompra = (int)oCmdActualiza.ExecuteScalar();
					#endregion Genera Nota De Credito

					#region Calculo de Cobranza
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.gridPagos.Rows.All)
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
							Estado = 0;

							Gasto = Convert.ToDecimal(dr.Cells["Gastos"].Value);
							IntMora = Convert.ToDecimal(dr.Cells["Mora"].Value);
							IntFin = Convert.ToDecimal(dr.Cells["Interes"].Value);
							Capital = Convert.ToDecimal(dr.Cells["Capital"].Value);									
							Abono = 0.00m;

							#region Gestión de cobranza
							if (Math.Round(Gasto, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Gasto, 2))
							{								
								Abono = Abono + Gasto;
								AbGastos = AbGastos + Gasto;
								Saldo = Saldo - Gasto;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Gasto, 2) > 0.00m && Math.Round(Gasto, 2) >= Math.Round(Saldo, 2)) 
							{
								Gasto = Gasto - Saldo;
								Abono = Abono + Saldo;
								AbGastos = AbGastos + Saldo;
								Saldo = 0.00m;
							}
							#endregion Gestión de cobranza

							#region Interés por mora
							if (Math.Round(IntMora, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntMora, 2))
							{
								Abono = Abono + IntMora;
								AbIntMora = AbIntMora + IntMora;
								Saldo = Saldo - IntMora;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntMora, 2) > 0.00m && Math.Round(IntMora, 2) >= Math.Round(Saldo, 2))
							{
								IntMora = IntMora - Saldo;
								Abono = Abono + Saldo;
								AbIntMora = AbIntMora + Saldo;
								Saldo = 0.00m;
							}
							#endregion Interés por mora
						
							#region Interés por financiamiento
							if (Math.Round(IntFin, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(IntFin, 2))
							{
								Abono = Abono + IntFin;
								AbInteres = AbInteres + IntFin;
								Saldo = Saldo - IntFin;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(IntFin, 2) > 0.00m && Math.Round(IntFin, 2) >= Math.Round(Saldo, 2))
							{
								IntFin = IntFin - Saldo;
								Abono = Abono + Saldo;
								AbInteres = AbInteres + Saldo;
								Saldo = 0.00m;
							}
							#endregion Interés por financiamiento

							#region Capital
							if (Math.Round(Capital, 2) > 0.00m && Math.Round(Saldo, 2) > 0.00m && Math.Round(Saldo, 2) >= Math.Round(Capital, 2))
							{
								Abono = Abono + Capital;
								AbCapital = AbCapital + Capital;
								Saldo = Saldo - Capital;
							}
							else if (Math.Round(Saldo, 2) > 0.00m && Math.Round(Capital, 2) > 0.00m && Math.Round(Capital, 2) >= Math.Round(Saldo, 2))
							{
								Capital = Capital - Saldo;
								Abono = Abono + Saldo;
								AbCapital = AbCapital + Saldo;
								Saldo = 0.00m;								
							}
							#endregion Capital

							#region Actualiza tabla Cre_TablaDeAmortización
							if (Abono > 0.00m)
							{										
								decimal dSaldoPrb = 0.00m;
								dSaldoPrb = Convert.ToDecimal(dr.Cells["Saldo"].Value);
								NSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value) - Abono;

								string sSQLAC = string.Format("Exec ActualizaCobranza {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}, {12}, {13}", 
									(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), idCompra, AbInteres, AbCapital, AbGastos, AbIntMora);
								oCmdActualiza.CommandText = sSQLAC;
								oCmdActualiza.ExecuteNonQuery();
							
								sConcepto = "";
								if (Math.Round(NSaldo, 2) == 0.00m) 
								{
									Estado = 2;
									sConcepto = "CANCELACIÓN CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
								}
								else if (Math.Round(NSaldo, 2) > 0.00m) 
								{
									Estado = 3;
									sConcepto = "ABONO CUOTA CREDIPOINT NÚMERO " + dr.Cells["NumeroCuota"].Value.ToString();
								}
							
								string sSQLTipoPago = string.Format("Select Pago From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
								oCmdActualiza.CommandText = sSQLTipoPago;
								int iPago = (int)oCmdActualiza.ExecuteScalar();

								
								#region  Descuento COVID 

								string sSQLDescGestion = string.Format("Select ISNULL (DescGestion ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
								oCmdActualiza.CommandText = sSQLDescGestion;
								decimal dDescGestion = (decimal)oCmdActualiza.ExecuteScalar();

								string sSQLDesMora = string.Format("Select ISNULL (DesMora ,0 ) From Cre_TablaDeAmortizacion Where idCre_TablaDeAmortizacion = {0}", (int)dr.Cells["idCre_TablaDeAmortizacion"].Value);
								oCmdActualiza.CommandText = sSQLDesMora;
								decimal  dDesMora = (decimal)oCmdActualiza.ExecuteScalar();

								#endregion  Descuento COVID 


								decimal dDebitado = AbIntMora + AbInteres + AbCapital;
								string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}",
									idCompra, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, 
									AbGastos, AbIntMora, AbInteres, AbCapital, Abono, Math.Round(NSaldo, 2), Estado, sConcepto, 
									Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), iPago, dDebitado, this.GridFacturas.ActiveRow.Cells["NumeroDeOperacion"].Value.ToString(), 0, dDescGestion, dDesMora);
								oCmdActualiza.CommandText = sSQLPCC;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Actualiza tabla Cre_TablaDeAmortización
						}
					}	
					#endregion Calculo de Cobranza

					#region Asiento
					string sSQLAsiento = string.Format("Exec AsientoCruceAnticipoNCCrediPoint {0}, 0, {1}, 0, {2}, '{3}'", 
						idCompra, 1, false, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")); 					
					oCmdActualiza.CommandText = sSQLAsiento;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Asiento

					#region Actualiza Saldo Estado En Documentos
					string sSQLSaldos = string.Format("Exec CobranzasActualizaSaldoCruceCuentas {0}, {1}, {2}", 
						2, idCompra, (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value);
					oCmdActualiza.CommandText = sSQLSaldos;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Actualiza Saldo Estado En Documentos 

					oTransaction.Commit();	

					oConCobranzas.Close();

					Cobranza();	

					this.Cursor = Cursors.Default;
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

					oConCobranzas.Close();					
				}
				finally
				{
					oConCobranzas.Close();
				}
			}			
		}

		private void txtIdentificacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnSNoDif_Click(object sender, System.EventArgs e)
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(1) From dbo.Temporal_Operaciones_Dif_Final Where Estado = 3 And idCompra = {0}", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value)) == 0)
			{
				MessageBox.Show("La operación no esta en diferimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("Esta seguro del solicitar el NO DIFERIMIENTO de las cuotas", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, 
							string.Format("Update Temporal_Operaciones_Dif_Final Set EstadoPoint = 2, FechaSolicitud = GETDATE(), UsuarioSolicitud = SUSER_SNAME(), NotasSolicitud = '{1}' Where idCompra = {0}", 
							(int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value, Nota.txtNotas.Text.ToString()));

						MessageBox.Show("Proceso en espera de autorizacion de cobranzas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
		}

		private void btnAux_Click(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE Cre_TablaDeAmortizacion SET DsctoIntMora = 1, DesMora = 0, DsctMorManual = 1, DsctGasManual = 1 WHERE idCompra = {0} AND VENDIDA = 0", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE Cre_TablaDeAmortizacion SET Interes = 0 WHERE idCompra = {0} AND VENDIDA = 0 and Estado Not In (2, 3, 4)", (int)this.GridFacturas.ActiveRow.Cells["idCompra"].Value));
		}

		private void txtCobro_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void GridFacturas_Click(object sender, System.EventArgs e)
		{
			this.Cobranza();
		}			
	}
}

