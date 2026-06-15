using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.IO;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCruzaDepocitosPendientes.
	/// </summary>
	public class frmCruzaDepocitosPendientes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblEstadoDepositos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorriente;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;

		int iBodega = 0;
		string sUrlNube = "";
		bool bActualiza = false;
		bool bEdicion = false;
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

		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaVencida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSiguienteCuota;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProximaVencer;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoAFavor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasVencidos;
		private Infragistics.Win.UltraWinGrid.UltraGrid gridPagos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoVencido;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProntoPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAutorizacion;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBancos;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnImportar;
	
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCruzaDepocitosPendientes()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCruzaDepocitosPendientes));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDepositosPendientes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("numeroDeposito");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Num_Fac");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fec_Fac");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlNube");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencia");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("N");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnLiquidar", 0);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnImg", 1);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnular", 2);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Abono", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Abono", 10, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDepositosPendientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("numeroDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num_Fac");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fec_Fac");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlNube");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Agencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("N");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vence");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuota");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InteresTotal");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABInteres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalTotal");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABCapital");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gastos");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DsctoGastos");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DsctoIntMora");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retraso");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCancelacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Secuencial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalDiferimiento");
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferimiento");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_TablaDeAmortizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vence");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InteresTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABInteres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DsctoGastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DsctoIntMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCancelacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Secuencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalDiferimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			this.btnVer = new System.Windows.Forms.Button();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblEstadoDepositos = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label1 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExcel = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCuotaVencida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSiguienteCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProximaVencer = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoAFavor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasVencidos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.gridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtSaldoTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtProntoPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAutorizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.cmbBancos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnImportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaVencida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSiguienteCuota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProximaVencer)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAFavor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).BeginInit();
			this.SuspendLayout();
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(432, 119);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(90, 26);
			this.btnVer.TabIndex = 714;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(10, 48);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 18);
			this.lblBodega.TabIndex = 712;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(230, 12);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 18);
			this.label6.TabIndex = 711;
			this.label6.Text = "Hasta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstadoDepositos
			// 
			this.lblEstadoDepositos.AutoSize = true;
			this.lblEstadoDepositos.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.lblEstadoDepositos.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstadoDepositos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoDepositos.Location = new System.Drawing.Point(922, 8);
			this.lblEstadoDepositos.Name = "lblEstadoDepositos";
			this.lblEstadoDepositos.Size = new System.Drawing.Size(0, 28);
			this.lblEstadoDepositos.TabIndex = 713;
			this.lblEstadoDepositos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(77, 46);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(326, 25);
			this.cmbBodega.TabIndex = 710;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBodega_KeyPress);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(278, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(125, 24);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 709;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
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
			// cmbCuentaCorriente
			// 
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataSource = this.ultraDataSource10;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.Caption = "";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 250;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCuentaCorriente.DisplayMember = "Descripcion";
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(840, 200);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(326, 102);
			this.cmbCuentaCorriente.TabIndex = 1038;
			this.cmbCuentaCorriente.ValueMember = "idCuenta";
			this.cmbCuentaCorriente.Visible = false;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 12;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 18;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn9.Header.Caption = "Banco";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn9.Width = 145;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn10.Width = 76;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Width = 84;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Header.Caption = "Cliente";
			ultraGridColumn12.Header.VisiblePosition = 14;
			ultraGridColumn12.Width = 194;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 80;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn14.CellAppearance = appearance5;
			ultraGridColumn14.Header.Caption = "N° Deposito";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 76;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance6;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Width = 120;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance7;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Width = 67;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance8;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Width = 84;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellAppearance = appearance9;
			ultraGridColumn18.Header.Caption = "Factura N.:";
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Width = 100;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance10;
			ultraGridColumn19.Header.Caption = "Fecha Factura";
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 100;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 24;
			ultraGridColumn20.Width = 210;
			ultraGridColumn21.Header.Caption = "Url Nube";
			ultraGridColumn21.Header.VisiblePosition = 23;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridColumn22.CellAppearance = appearance11;
			ultraGridColumn22.Header.VisiblePosition = 11;
			ultraGridColumn22.Width = 84;
			ultraGridColumn23.Header.VisiblePosition = 21;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.Caption = "Fecha Pago";
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn24.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownCalendar;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn25.Width = 27;
			ultraGridColumn26.Header.Caption = "N °";
			ultraGridColumn26.Header.VisiblePosition = 5;
			ultraGridColumn26.Width = 35;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn27.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn27.Header.Caption = "Fecha Ingreso";
			ultraGridColumn27.Header.VisiblePosition = 22;
			ultraGridColumn27.Width = 120;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance12.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance12.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn28.CellAppearance = appearance12;
			appearance13.Image = ((object)(resources.GetObject("appearance13.Image")));
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance13.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn28.CellButtonAppearance = appearance13;
			ultraGridColumn28.Header.Caption = "...";
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 46;
			ultraGridColumn29.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance14.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn29.CellAppearance = appearance14;
			appearance15.Image = ((object)(resources.GetObject("appearance15.Image")));
			appearance15.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance15.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn29.CellButtonAppearance = appearance15;
			ultraGridColumn29.Header.Caption = "...";
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn29.Width = 35;
			ultraGridColumn30.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance16.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance16.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn30.CellAppearance = appearance16;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			appearance17.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance17.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn30.CellButtonAppearance = appearance17;
			ultraGridColumn30.Header.Caption = "...";
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn30.Width = 27;
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
																										 ultraGridColumn30});
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance18;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance19.ForeColor = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.CaptionAppearance = appearance19;
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance20;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 8F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance22;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(10, 175);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1526, 379);
			this.ultraGrid1.TabIndex = 1037;
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(System.DateTime);
			ultraDataColumn22.DataType = typeof(bool);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24});
			// 
			// optEstados
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance25;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance26;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Pendientes";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Verificado";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Anulados";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Cruzado";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.Items.Add(valueListItem4);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(432, 7);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(386, 28);
			this.optEstados.TabIndex = 1039;
			this.optEstados.Text = "Pendientes";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(10, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 18);
			this.label1.TabIndex = 1040;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance27;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(77, 9);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(131, 24);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1041;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(538, 118);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(90, 27);
			this.btnExcel.TabIndex = 1042;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label2.Location = new System.Drawing.Point(10, 122);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(47, 20);
			this.label2.TabIndex = 1044;
			this.label2.Text = "Buscar";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBusqueda
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance28;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(77, 120);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(326, 24);
			this.txtBusqueda.TabIndex = 1043;
			// 
			// txtCuotaVencida
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaVencida.Appearance = appearance29;
			this.txtCuotaVencida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaVencida.Enabled = false;
			this.txtCuotaVencida.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtCuotaVencida.FormatString = "";
			this.txtCuotaVencida.Location = new System.Drawing.Point(1504, 8);
			this.txtCuotaVencida.MinValue = 0;
			this.txtCuotaVencida.Name = "txtCuotaVencida";
			this.txtCuotaVencida.PromptChar = ' ';
			this.txtCuotaVencida.Size = new System.Drawing.Size(16, 26);
			this.txtCuotaVencida.TabIndex = 1046;
			this.txtCuotaVencida.Visible = false;
			// 
			// txtSiguienteCuota
			// 
			appearance30.ForeColor = System.Drawing.Color.Black;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSiguienteCuota.Appearance = appearance30;
			this.txtSiguienteCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSiguienteCuota.Enabled = false;
			this.txtSiguienteCuota.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSiguienteCuota.FormatString = "#,##0.00";
			this.txtSiguienteCuota.Location = new System.Drawing.Point(1480, 8);
			this.txtSiguienteCuota.MinValue = 0;
			this.txtSiguienteCuota.Name = "txtSiguienteCuota";
			this.txtSiguienteCuota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSiguienteCuota.PromptChar = ' ';
			this.txtSiguienteCuota.Size = new System.Drawing.Size(16, 26);
			this.txtSiguienteCuota.TabIndex = 1047;
			this.txtSiguienteCuota.Visible = false;
			// 
			// txtProximaVencer
			// 
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProximaVencer.Appearance = appearance31;
			this.txtProximaVencer.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProximaVencer.Enabled = false;
			this.txtProximaVencer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtProximaVencer.FormatString = "";
			this.txtProximaVencer.Location = new System.Drawing.Point(1456, 8);
			this.txtProximaVencer.MinValue = 0;
			this.txtProximaVencer.Name = "txtProximaVencer";
			this.txtProximaVencer.PromptChar = ' ';
			this.txtProximaVencer.Size = new System.Drawing.Size(16, 26);
			this.txtProximaVencer.TabIndex = 1048;
			this.txtProximaVencer.Visible = false;
			// 
			// txtSaldoAFavor
			// 
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoAFavor.Appearance = appearance32;
			this.txtSaldoAFavor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoAFavor.Enabled = false;
			this.txtSaldoAFavor.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoAFavor.FormatString = "#,##0.00";
			this.txtSaldoAFavor.Location = new System.Drawing.Point(1432, 8);
			this.txtSaldoAFavor.MinValue = -2147483647;
			this.txtSaldoAFavor.Name = "txtSaldoAFavor";
			this.txtSaldoAFavor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoAFavor.PromptChar = ' ';
			this.txtSaldoAFavor.Size = new System.Drawing.Size(16, 26);
			this.txtSaldoAFavor.TabIndex = 1049;
			this.txtSaldoAFavor.Visible = false;
			// 
			// txtDiasVencidos
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasVencidos.Appearance = appearance33;
			this.txtDiasVencidos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasVencidos.Enabled = false;
			this.txtDiasVencidos.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDiasVencidos.FormatString = "";
			this.txtDiasVencidos.Location = new System.Drawing.Point(1408, 8);
			this.txtDiasVencidos.MinValue = -2147483647;
			this.txtDiasVencidos.Name = "txtDiasVencidos";
			this.txtDiasVencidos.PromptChar = ' ';
			this.txtDiasVencidos.Size = new System.Drawing.Size(16, 26);
			this.txtDiasVencidos.TabIndex = 1050;
			this.txtDiasVencidos.Visible = false;
			// 
			// gridPagos
			// 
			this.gridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gridPagos.CausesValidation = false;
			this.gridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.gridPagos.DataSource = this.ultraDataSource2;
			appearance34.BackColor = System.Drawing.Color.White;
			this.gridPagos.DisplayLayout.Appearance = appearance34;
			this.gridPagos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 105;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 76;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 53;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance35;
			ultraGridColumn34.Header.Caption = "Cuota";
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn34.Width = 98;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 4;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 107;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn36.Header.VisiblePosition = 5;
			ultraGridColumn36.Width = 103;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance36;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "Valor";
			ultraGridColumn37.Header.VisiblePosition = 6;
			ultraGridColumn37.Width = 97;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance37;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Total Interes";
			ultraGridColumn38.Header.VisiblePosition = 7;
			ultraGridColumn38.Width = 85;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance38;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "Interés";
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Width = 63;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 12;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance39;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "Total Capital";
			ultraGridColumn41.Header.VisiblePosition = 8;
			ultraGridColumn41.Width = 85;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance40;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 11;
			ultraGridColumn42.Width = 63;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 17;
			ultraGridColumn43.Hidden = true;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance41;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "Gestión";
			ultraGridColumn44.Header.VisiblePosition = 13;
			ultraGridColumn44.MaskInput = "";
			ultraGridColumn44.PromptChar = ' ';
			ultraGridColumn44.Width = 69;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance42;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "% Dscto";
			ultraGridColumn45.Header.VisiblePosition = 14;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.PromptChar = ' ';
			ultraGridColumn45.Width = 66;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance43;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 15;
			ultraGridColumn46.MaskInput = "";
			ultraGridColumn46.PromptChar = ' ';
			ultraGridColumn46.Width = 72;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance44;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "% Dscto";
			ultraGridColumn47.Header.VisiblePosition = 18;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.PromptChar = ' ';
			ultraGridColumn47.Width = 72;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance45;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 19;
			ultraGridColumn48.Width = 79;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn49.CellAppearance = appearance46;
			ultraGridColumn49.Header.VisiblePosition = 22;
			ultraGridColumn49.Width = 81;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn50.Header.VisiblePosition = 21;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 57;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance47;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 20;
			ultraGridColumn51.Width = 85;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.Caption = "Pago";
			ultraGridColumn52.Header.VisiblePosition = 23;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 66;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn53.Header.Caption = "Documento N.";
			ultraGridColumn53.Header.VisiblePosition = 24;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 87;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn54.Header.Caption = "Forma de pago";
			ultraGridColumn54.Header.VisiblePosition = 26;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 107;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn55.Header.VisiblePosition = 27;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 108;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn56.Header.Caption = "Local";
			ultraGridColumn56.Header.VisiblePosition = 25;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 183;
			ultraGridColumn57.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn57.Header.VisiblePosition = 28;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 207;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance48;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "T Diferimiento";
			ultraGridColumn58.Header.VisiblePosition = 9;
			ultraGridColumn58.Width = 69;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance49;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 16;
			ultraGridColumn59.Width = 75;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn60.Header.VisiblePosition = 29;
			ultraGridColumn60.Width = 114;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60});
			ultraGridBand4.Header.Caption = "";
			ultraGridBand4.HeaderVisible = true;
			appearance50.ForeColor = System.Drawing.Color.Black;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand4.Override.ActiveRowAppearance = appearance50;
			ultraGridBand4.SummaryFooterCaption = "";
			this.gridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance51.BackColor = System.Drawing.Color.Transparent;
			this.gridPagos.DisplayLayout.Override.CardAreaAppearance = appearance51;
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.gridPagos.DisplayLayout.Override.CellAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance53.FontData.BoldAsString = "True";
			appearance53.FontData.Name = "Arial";
			appearance53.FontData.SizeInPoints = 8F;
			appearance53.ForeColor = System.Drawing.Color.White;
			appearance53.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.gridPagos.DisplayLayout.Override.HeaderAppearance = appearance53;
			appearance54.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance54.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.RowAlternateAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance56;
			this.gridPagos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.gridPagos.Location = new System.Drawing.Point(230, 277);
			this.gridPagos.Name = "gridPagos";
			this.gridPagos.Size = new System.Drawing.Size(1259, 194);
			this.gridPagos.TabIndex = 1051;
			this.gridPagos.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(System.DateTime);
			ultraDataColumn30.DataType = typeof(System.DateTime);
			ultraDataColumn31.DataType = typeof(System.Double);
			ultraDataColumn31.DefaultValue = 0;
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn33.DataType = typeof(System.Double);
			ultraDataColumn33.DefaultValue = 0;
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn34.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn35.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn36.DataType = typeof(System.Double);
			ultraDataColumn36.DefaultValue = 0;
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(System.Double);
			ultraDataColumn38.DefaultValue = 0;
			ultraDataColumn39.DataType = typeof(System.Double);
			ultraDataColumn39.DefaultValue = 0;
			ultraDataColumn40.DataType = typeof(System.Double);
			ultraDataColumn40.DefaultValue = 0;
			ultraDataColumn41.DataType = typeof(System.Double);
			ultraDataColumn41.DefaultValue = 0;
			ultraDataColumn42.DataType = typeof(System.Double);
			ultraDataColumn42.DefaultValue = 0;
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(System.Double);
			ultraDataColumn45.DefaultValue = 0;
			ultraDataColumn46.DataType = typeof(System.DateTime);
			ultraDataColumn46.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn52.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn53.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54});
			// 
			// txtSaldoTotal
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoTotal.Appearance = appearance57;
			this.txtSaldoTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoTotal.Enabled = false;
			this.txtSaldoTotal.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoTotal.FormatString = "#,##0.00";
			this.txtSaldoTotal.Location = new System.Drawing.Point(1384, 8);
			this.txtSaldoTotal.MinValue = 0;
			this.txtSaldoTotal.Name = "txtSaldoTotal";
			this.txtSaldoTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoTotal.PromptChar = ' ';
			this.txtSaldoTotal.Size = new System.Drawing.Size(16, 26);
			this.txtSaldoTotal.TabIndex = 1052;
			this.txtSaldoTotal.Visible = false;
			// 
			// txtSaldoVencido
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoVencido.Appearance = appearance58;
			this.txtSaldoVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoVencido.Enabled = false;
			this.txtSaldoVencido.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtSaldoVencido.FormatString = "#,##0.00";
			this.txtSaldoVencido.Location = new System.Drawing.Point(1360, 8);
			this.txtSaldoVencido.MinValue = -2147483647;
			this.txtSaldoVencido.Name = "txtSaldoVencido";
			this.txtSaldoVencido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoVencido.PromptChar = ' ';
			this.txtSaldoVencido.Size = new System.Drawing.Size(16, 26);
			this.txtSaldoVencido.TabIndex = 1053;
			this.txtSaldoVencido.Visible = false;
			// 
			// txtProntoPago
			// 
			appearance59.ForeColor = System.Drawing.Color.Black;
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProntoPago.Appearance = appearance59;
			this.txtProntoPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProntoPago.Enabled = false;
			this.txtProntoPago.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtProntoPago.FormatString = "#,##0.00";
			this.txtProntoPago.Location = new System.Drawing.Point(1336, 8);
			this.txtProntoPago.MinValue = -2147483647;
			this.txtProntoPago.Name = "txtProntoPago";
			this.txtProntoPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProntoPago.PromptChar = ' ';
			this.txtProntoPago.Size = new System.Drawing.Size(16, 26);
			this.txtProntoPago.TabIndex = 1054;
			this.txtProntoPago.Visible = false;
			// 
			// txtIdAutorizacion
			// 
			this.txtIdAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAutorizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAutorizacion.Enabled = false;
			this.txtIdAutorizacion.Location = new System.Drawing.Point(1307, 9);
			this.txtIdAutorizacion.Name = "txtIdAutorizacion";
			this.txtIdAutorizacion.PromptChar = ' ';
			this.txtIdAutorizacion.Size = new System.Drawing.Size(18, 24);
			this.txtIdAutorizacion.TabIndex = 1055;
			this.txtIdAutorizacion.Visible = false;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(10, 87);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(43, 18);
			this.label26.TabIndex = 1056;
			this.label26.Text = "Banco";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBancos
			// 
			this.cmbBancos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBancos.DataSource = this.ultraDataSource10;
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 1;
			ultraGridColumn62.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn61,
																										 ultraGridColumn62});
			this.cmbBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBancos.DisplayMember = "Descripcion";
			this.cmbBancos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBancos.Location = new System.Drawing.Point(77, 83);
			this.cmbBancos.Name = "cmbBancos";
			this.cmbBancos.Size = new System.Drawing.Size(326, 26);
			this.cmbBancos.TabIndex = 1057;
			this.cmbBancos.ValueMember = "idCuenta";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 157);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1536, 9);
			this.groupBox1.TabIndex = 1058;
			this.groupBox1.TabStop = false;
			// 
			// btnImportar
			// 
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(643, 118);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(192, 27);
			this.btnImportar.TabIndex = 1059;
			this.btnImportar.Text = "Importar Estado De Cuenta";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// frmCruzaDepocitosPendientes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(1536, 560);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbBancos);
			this.Controls.Add(this.txtIdAutorizacion);
			this.Controls.Add(this.txtProntoPago);
			this.Controls.Add(this.txtSaldoVencido);
			this.Controls.Add(this.txtSaldoTotal);
			this.Controls.Add(this.gridPagos);
			this.Controls.Add(this.txtDiasVencidos);
			this.Controls.Add(this.txtSaldoAFavor);
			this.Controls.Add(this.txtProximaVencer);
			this.Controls.Add(this.txtSiguienteCuota);
			this.Controls.Add(this.txtCuotaVencida);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.lblEstadoDepositos);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Name = "frmCruzaDepocitosPendientes";
			this.Text = "Cruza Depositos Pendientes";
			this.Load += new System.EventHandler(this.frmCruzaDepocitosPendientes_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaVencida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSiguienteCuota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProximaVencer)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAFavor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasVencidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAutorizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBancos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCruzaDepocitosPendientes_Load(object sender, System.EventArgs e)
		{
			#region Seguridad
			miAcceso = new Acceso(cdsSeteoF, "0449");

			if (!Funcion.Permiso("1866", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Cruce Depósitos Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			#endregion Seguridad
		
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodega > 0)
			{
				this.cmbBodega.Value = iBodega; 
				this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCuentasDepositos "));
			}
			else 
			{
				this.cmbBodega.Enabled = true;
				//this.cmbBodega.Value = 57;	
				this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCuentasDepositos "));
			}
			this.cmbBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaCuentasDepositos "));

			this.dtFecha.Enabled = true;	
			this.dtDesde.Enabled = true;	
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtDesde, cdsSeteoF, false);
		}

		private void ConsultaFacturas(int iIdCompra)
		{
			//string sSQL = string.Format("Exec ConsultaFacturasDepositosPendientes {0}", iIdCompra);
			//FuncionesProcedimientos.dsGeneral(sSQL, this.GridFacturas);


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


		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) e.Cancel = true;
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 45, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iBodega = 0;
			int iBanco = 0;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Inicial ", true, "DEPOSITOS", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha Final ", true, "DEPOSITOS", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtFecha.Value), this.dtDesde, "La Fecha Inicial no puede ser mayor a la Fecha Final", false)) return;

			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;
			if (this.cmbBancos.ActiveRow != null) iBanco = (int)this.cmbBancos.Value;

			string sSQL = string.Format("Exec ListaDepositosPorCruzar '{0}', '{1}', {2},{3},'{4}',{5}", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
				iBodega, (int)this.optEstados.Value, this.txtBusqueda.Text.ToString(), iBanco);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblEstadoDepositos.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void cmbBodega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ConsultaTablaDeAmortizacion(int idCompra,DateTime dtFechaPago)
		{
			string sSQL = string.Format("Exec Cobranzas {0}, '{1}'", 
				idCompra, Convert.ToDateTime(dtFechaPago).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.gridPagos);			
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region btnLiquidar
			int sel=0;

			if (e.Cell.Column.ToString() == "btnLiquidar")
			{
				if (!miAcceso.BReliquidar) 
				{
					MessageBox.Show("No tiene acceso a Liquidar ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if (e.Cell.Row.Cells["IdDepositosPendientes"].Value != System.DBNull.Value)
				{
					iFormaPago=7;
					if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de liquidar el pago ?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{	
						#region ListaDepositosPorLiquidar
						if ((bool)e.Cell.Row.Cells["Sel"].Value)
						{
							sel=1;
							iFormaPago=24;
							using (frmListaDepositosPorLiquidar LDP = new frmListaDepositosPorLiquidar((int)e.Cell.Row.Cells["idCuenta"].Value))
							{
								if (DialogResult.OK == LDP.ShowDialog())
								{
									idDepositosPorliquida = (int)LDP.ultraGrid1.ActiveRow.Cells["idDepositosPorLiquidar"].Value;

								}
								else 
								{
									return;
								}
							}
						}
						#endregion ListaDepositosPorLiquidar
					
						#region Valida
					

						string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec ValidaEstadoDepositosPendientes {0}", (int)e.Cell.Row.Cells["IdDepositosPendientes"].Value));
						if (sMensaje.Length > 0)
						{
							MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
							return;
						}
						#endregion Valida

						#region Datos

						if (!this.ValidaSaldosAFavor()) return;
			
						if (!ValidaSaldosRefinanciacion()) return;
						ConsultaFacturas((int)e.Cell.Row.Cells["IdCompra"].Value);

						Cobranza(Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value));

						ConsultaTablaDeAmortizacion((int)e.Cell.Row.Cells["IdCompra"].Value,Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value));

						bool bCobroManual = false;
 
						bCobroManual = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaCobroManual {0}, {1}", 
							(int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value,(decimal)this.ultraGrid1.ActiveRow.Cells["Abono"].Value), true);

						if (!Validacion.vbCarteraEnTransito(iIdEntidadFinanciera,NumeroLote,Confirmado)) bCobroManual = true;

						#endregion Datos

						#region Cobro Manual
      			
						if (bCobroManual)
						{
							#region Guardar  
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
									#region Proceso
									int idAnticipo = 0;
									decimal dAbono = Convert.ToDecimal(e.Cell.Row.Cells["Abono"].Value);
									string sNumero = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 24, 57, 0");
															
									string sSQL = string.Format("Exec GuardaAnticiposDepositosPendientes 0, 0, '{0}', 57, '{1}',{2},{3} , 0, 'COBRO DE DEPOSITO POR LIQUIDAR',1, {4}, 12, '', {5}, 0",Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"),sNumero,(int)e.Cell.Row.Cells["IdCliente"].Value,iFormaPago,dAbono,(int)e.Cell.Row.Cells["IdCompra"].Value);
									idAnticipo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Insert Into DetalleFormaPago (idCompra, Origen, idFormaPago, Valor,idCuenta)	Values ({0}, 2,{1}, {2},{3})",
										idAnticipo,iFormaPago,dAbono,(int)e.Cell.Row.Cells["idCuenta"].Value));
							
									Funcion.EjecutaSQL(cdsSeteoF, 
										string.Format("Update DepositosPendientes Set IdAnticipo = {0}, Estado = 2 ,Sel={1}Where IdDepositosPendientes = {2}", 
										idAnticipo,sel,(int)e.Cell.Row.Cells["IdDepositosPendientes"].Value));
									#endregion Proceso


									#region Notifiacion Pago
									string sSQLNotificaPago = string.Format("Exec GuardaNotifiacionCuotaPagos {0}, {1}, {2},{3}, {4}",(int)e.Cell.Row.Cells["IdCompra"].Value,idAnticipo,0,0,1);
									oCmdActualiza.CommandText = sSQLNotificaPago;	
									oCmdActualiza.ExecuteNonQuery();														
									#endregion Notifiacion Pago
 
									Facturacion.ImpresionRecibos(idAnticipo, 12, cdsSeteoF);

									this.btnVer_Click(sender, e);
	
						
								}
								catch(Exception ex)
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
							#endregion Guardar

							return;
						}
						#endregion Cobro Manual

						#region Registro de Cobro
						int idCompraPago = 0;
						bool bNormalAtrasado = false;	
      
						if (!bCobroManual)
						{
							#region Conexion
							this.Cursor = Cursors.WaitCursor;

							using (SqlConnection oConCobranzas = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
							{
								oConCobranzas.Open();

								SqlTransaction oTransaction = null;
								SqlCommand oCmdActualiza = oConCobranzas.CreateCommand();
								oCmdActualiza.Connection = oConCobranzas;
								oCmdActualiza.CommandType = CommandType.Text;
								oCmdActualiza.CommandTimeout = 120;
								#endregion Conexion
						
								try
								{							
									oTransaction = oConCobranzas.BeginTransaction();
									oCmdActualiza.Transaction = oTransaction;

									#region proceso
									#region Variables
									decimal Cobro = (decimal)this.ultraGrid1.ActiveRow.Cells["Abono"].Value;
									decimal Diferimiento = 0.00m;
									decimal ABDiferimiento = 0.00m;
									decimal DescDiferimiento = 0.00m;
									decimal Gasto = 0.00m;
									decimal IntMora = 0.00m;
									decimal IntFin = 0.00m;
									decimal Capital = 0.00m;
									decimal Abono = 0.00m;
									decimal Saldo = (decimal)this.ultraGrid1.ActiveRow.Cells["Abono"].Value;
									decimal NSaldo = 0.00m;							
									decimal AbGastos = 0.00m;
									decimal AbIntMora = 0.00m;
									decimal AbInteres = 0.00m;
									decimal AbCapital = 0.00m;							
									int Estado = 0;
									string sConcepto = "";
									string sVacio = "";
									int idCompraOrigen = (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value;							
									#endregion Variables
														
									#region Genera Abono Factura 24

									string sNumero = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeracionLocales 24, 57, 0");

									if (Cobro < Convert.ToDecimal(this.txtSaldoVencido.Value)) bNormalAtrasado = true;

									string sSQLPago = string.Format("Exec GuardaPagoCobranzaDepositosPendientes {0}, {1}, {2}, 0, '{3}', '{4}',{5},{6}", 
										idCompraOrigen,57, Math.Round(Convert.ToDecimal(Cobro), 2),Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"), sNumero, bNormalAtrasado,iFormaPago);
									oCmdActualiza.CommandText = sSQLPago;								
									idCompraPago = (int)oCmdActualiza.ExecuteScalar(); 
									#endregion Genera Abono Factura 24

									#region Guarda Cobro de la factura
					
									//int idDetFormaPago = 7;
									int idBanco = (int)this.ultraGrid1.ActiveRow.Cells["idBanco"].Value;	
									int idTarjeta = 0;
									int idCuenta = (int)this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
									int idPlan = 0;
									int idPlazo = 0;
									decimal dComTarjeta = 0.00m;
									DateTime dtFecCaducidad = DateTime.Today;

									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, 0", 
										0, idCompraPago, 3,iFormaPago, Cobro,
										idBanco, idTarjeta,
										sVacio,(string)this.ultraGrid1.ActiveRow.Cells["numeroDeposito"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"),								
										(string)this.ultraGrid1.ActiveRow.Cells["numeroDeposito"].Value.ToString(), 
										sVacio,idCuenta,sVacio, sVacio,idPlan, idPlazo, dComTarjeta);
									oCmdActualiza.CommandText = sSQLCobro;
									oCmdActualiza.ExecuteNonQuery(); 										
									
		
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
												if (Convert.ToDateTime(dr.Cells["Vence"].Value) >= Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value) && (Capital + IntFin + IntMora + Gasto) <= Saldo)
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
													(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, Gasto, IntMora, IntFin, Capital, Abono, Math.Round(NSaldo, 2), (int)dr.Cells["Retraso"].Value,Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"), idCompraPago, AbInteres, AbCapital, AbGastos, AbIntMora, ABDiferimiento, DescDiferimiento);
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
												decimal dDesMora = 0.00m;

												#endregion  Descuento COVID 

												decimal dDebitado = AbIntMora + AbInteres + AbCapital + ABDiferimiento;
										
												decimal dRescate = 0.00m;
										
												string sSQLRescate = string.Format("Exec Cre_ValidaSaldoCuotaRescatada {0}, {1}, '{2}'", 
													(int)dr.Cells["idCre_TablaDeAmortizacion"].Value, dDebitado,Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"));
												oCmdActualiza.CommandText = sSQLRescate;
												dRescate = (decimal)oCmdActualiza.ExecuteScalar();

												string sSQLPCC = string.Format("Exec GuardaPagosCuotasCredito {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, '{9}', '{10}', {11}, {12}, 1, '{13}', {14}, {15}, {16}, {17}, {18}, {19}",
													idCompraPago, (int)dr.Cells["idCre_TablaDeAmortizacion"].Value, AbGastos, AbIntMora, AbInteres, AbCapital, Abono, //6
													Math.Round(NSaldo, 2), Estado, sConcepto,Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"), iPago, dDebitado, NumeroDeOperacion, //14
													0, dDescGestion, dDesMora, ABDiferimiento, DescDiferimiento, dRescate);
												oCmdActualiza.CommandText = sSQLPCC;
												oCmdActualiza.ExecuteNonQuery();

												#region Actualiza Pago en gestiones
												decimal dValorPagadoGestiones = AbCapital + AbInteres;
												string sSQLAPG = string.Format("Exec Cre_ActualizaValorPagoGestiones {0}, '{1}', '{2}', {3}, {4}",
													idCompraOrigen,Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(dr.Cells["Vence"].Value).ToString("yyyyMMdd"), dValorPagadoGestiones, Abono);
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

									string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 24, 9 Exec ActualizaSaldoFacturas {0}, 1, 9", idCompraPago, (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
									oCmdActualiza.CommandText = sSQLSaldo;
									oCmdActualiza.ExecuteNonQuery();

									string sSQLCarteraExterna = string.Format("Exec GeneraMovimientosCarteraExterna {0}, {1}", idCompraPago, (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
									oCmdActualiza.CommandText = sSQLCarteraExterna;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Asiento - Saldos
							
									#region Actualiza Compromisos de Pago - Estado de Cuenta 
									string sSQLActCompromisoPago = string.Format("Exec CPActualizaCompromisoPago {0}, '{1}', {2}", 
										(int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value,Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"), Convert.ToDecimal((decimal)this.ultraGrid1.ActiveRow.Cells["Abono"].Value));
									oCmdActualiza.CommandText = sSQLActCompromisoPago;
									oCmdActualiza.ExecuteNonQuery();								

									string sSQLAEC = string.Format("Exec Cre_ActualizaEstadoDeCuentaDesdePago {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
									oCmdActualiza.CommandText = sSQLAEC;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Actualiza Compromisos de Pago - Estado de Cuenta

									#region Notifiacion Pago
									string sSQLNotificaPago = string.Format("Exec GuardaNotifiacionCuotaPagos {0}, {1}, {2},{3}, {4}",(int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value,idCompraPago,0,0,1);
									oCmdActualiza.CommandText = sSQLNotificaPago;
									oCmdActualiza.ExecuteNonQuery();
									#endregion Notifiacion Pago

									#region Libera celulares	
									oCmdActualiza.CommandText = string.Format("Select COUNT(1) From dbo.Cre_TablaDeAmortizacion Where Estado Not In (2, 4) and idCompra = {0}", 
										idCompraOrigen);
									int iValEstadoCel = (int) oCmdActualiza.ExecuteScalar();

									if (EstadoCelular==1)
									{															
										string sSQLBLO = string.Format("Exec LiberacionDeCelulares {0}",idCompraOrigen);
										oCmdActualiza.CommandText = sSQLBLO;
										oCmdActualiza.ExecuteNonQuery();									
									}
									#endregion Libera celulares

									#region Actualiza estado Depositos Pendientes

									Funcion.EjecutaSQL(cdsSeteoF, 
										string.Format("Update DepositosPendientes Set IdAnticipo = {0},idDepositosPorLiquidar={1}, Estado = 4 ,Sel={2} Where IdDepositosPendientes ={3}",idCompraPago,idDepositosPorliquida,sel,(int)e.Cell.Row.Cells["IdDepositosPendientes"].Value));
	
									#endregion Actualiza estado Depositos Pendientes

									if ((bool)e.Cell.Row.Cells["Sel"].Value)
									{
										#region Actualiza estado Depositos Por Liquidar

										Funcion.EjecutaSQL(cdsSeteoF, 
											string.Format("Update DepositosPorLiquidar Set idCompra = {0}, Estado = 2 Where idDepositosPorLiquidar = {1}", 
											idCompraPago, idDepositosPorliquida));

										#endregion Actualiza estado Depositos Por Liquidar		
									}

									#endregion proceso
							
									oTransaction.Commit();	

									oConCobranzas.Close();

									#region Impresion
									using (frmNumeracionRecibo miMsje = new frmNumeracionRecibo (Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(Secuencial, '') From Anticipos Where idAnticipo = {0}", idCompraPago))))
									{
										if (DialogResult.OK == miMsje.ShowDialog())
											Facturacion.ImpresionRecibos(idCompraPago, 2, cdsSeteoF);
									}
									#endregion Impresion

									this.btnVer_Click(sender, e);	

									this.Cursor = Cursors.Default;
								}
								catch (Exception ex)
								{
									#region Exceptiones
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
									#endregion Exceptiones
								}
								finally
								{
									oConCobranzas.Close();
								}
							}
						}
					
					
						#endregion Registro de Cobro
					}
					else {
					return;
					}
				}
			}
			#endregion btnLiquidar

			#region img
			if (e.Cell.Column.ToString() == "btnImg")
			{
				if (!miAcceso.BImagen) 
				{
					MessageBox.Show("No tiene acceso a Imagen ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				sUrlNube = (string) e.Cell.Row.Cells["UrlNube"].Value;
				if (sUrlNube.Length > 10 )
				{
					Process.Start(sUrlNube);

					System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
				}
			}
			#endregion img

			#region Anular
			if (e.Cell.Column.ToString() == "btnAnular")
			{	
				if (!miAcceso.Anular) 
				{
					MessageBox.Show("No tiene acceso Anular ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				#region Abre observacion
				bEdicion=true;
				using (frmObservacionDepositosPendientes Observaciones = new frmObservacionDepositosPendientes(1, e.Cell.Row.Cells["Detalle"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["Detalle"].Value = Convert.ToString(e.Cell.Row.Cells["Detalle"].Value)+"-" +Observaciones.txtObservaciones.Text.ToString().Trim();
						#region Guardar      
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

								string sSQLAutEs = string.Format("Update DepositosPendientes set Estado = 3,Detalle='{1}'  where IdDepositosPendientes = {0}",
									(int) e.Cell.Row.Cells["IdDepositosPendientes"].Value,(string)e.Cell.Row.Cells["Detalle"].Value );
								oCmdActualiza.CommandText = sSQLAutEs;
								oCmdActualiza.ExecuteNonQuery();					
						
		
								oTransaction.Commit();
								Cursor = Cursors.Default;	

								MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.btnVer_Click (sender, e);
				
							}
							catch(Exception ex)
							{
								#region execpciones
								try
								{ 
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
								catch (Exception ex2)
								{
									MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
								}
								#endregion execpciones
							}
							finally
							{
								oConexion.Close();
							}
						}
				
						#endregion Guardar					
					}
				}
				#endregion Abre observacion
			}
			#endregion Anular
		}

		private void Cobranza(DateTime dtFechaPago)
		{	
				try
				{
											
					#region Valores Estado de Cuenta
					string sSQL = string.Format("Exec ValoresCobranza {0}, '{1}'", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value,Convert.ToDateTime(dtFechaPago).ToString("yyyyMMdd"));
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

					string sSQLUFP = string.Format("Exec CobranzasUltimaFechaPago {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value);
					dtUltimaFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, sSQLUFP);

					ConsultaTablaDeAmortizacion((int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value,Convert.ToDateTime(dtFechaPago));	
					
					#region Valor re liquidacion
					decimal dGC = 0.00m;
					decimal dIM = 0.00m;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drPagos in this.gridPagos.Rows.All)
					{
						dGC = dGC + Convert.ToDecimal(drPagos.Cells["Gastos"].Value);
						dIM = dIM + Convert.ToDecimal(drPagos.Cells["Mora"].Value);
					}

					sSQL = string.Format("Exec ConsultaValorReliquidacion {0}, '{1}', {2}, {3}, 1", 
						(int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value,Convert.ToDateTime(dtFechaPago).ToString("yyyyMMdd"), dGC, dIM);
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


		private bool ValidaSaldosAFavor()
		{
			bool bValidaSaldoPendiente = true;

			if (iIdEntidadFinanciera == 4 && !(bool)Confirmado && iDiasVencidos <= 0) bValidaSaldoPendiente = false;
			if (iIdEntidadFinanciera == 15 && !(bool)Confirmado && iDiasVencidos <= 0) bValidaSaldoPendiente = false;
			if (iIdEntidadFinanciera == 16 && !(bool)Confirmado && iDiasVencidos <= 0) bValidaSaldoPendiente = false;
			return true;
		}

		private bool ValidaSaldosRefinanciacion()
		{
			int iCuotasPendientesR = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CobranzaValidaRefinancion {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value));

			if (iCuotasPendientesR > 0)
			{
				MessageBox.Show("No puede generar Cobro a documentos de refinanciacion si el credito inicial no ha sido cancelado en su totalidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}

			return true;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			#region idCuenta

			if (e.Cell.Column.ToString() == "idCuenta")
			{	
				bActualiza = true;
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar la cuenta?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE DepositosPendientes SET Idbanco = {0} WHERE IdDepositosPendientes = {1} ", 
						(int)	this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value, (int)	this.ultraGrid1.ActiveRow.Cells["IdDepositosPendientes"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
					this.btnVer_Click(sender, e);		
				}
				else
				{						
					e.Cell.Row.Cells["idCuenta"].Value = (int)e.Cell.OriginalValue;
				}					
				bActualiza = false;
			}
			#endregion idCuenta

			#region FechaPago
			if (e.Cell.Column.ToString() == "FechaPago")
			{	
				bActualiza = true;
				DateTime dtHoy = FuncionesProcedimientos.dtRetornaFechaServidor(cdsSeteoF);	
				DateTime dtFechaMes = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT dbo.RetornaFechaInicioMes('{0}', - 1)",Convert.ToDateTime(dtHoy).ToString("yyyyMMdd")));	
				DateTime gdFechagrid = Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value);

				if (gdFechagrid < dtFechaMes )
				{
					MessageBox.Show(string.Format("La Fecha de Cobro no puede ser menor a '{0}'", dtFechaMes.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return;
				}
				if (Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value) > Convert.ToDateTime(dtHoy)  )
				{
					MessageBox.Show(string.Format("La Fecha de Cobro no puede ser mayor a '{0}'", dtHoy.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return;
				}

				string sSQLUFP = string.Format("Exec CobranzasUltimaFechaPago {0}", (int)e.Cell.Row.Cells["IdCompra"].Value);
				dtUltimaFechaPago = Funcion.dtEscalarSQL(cdsSeteoF, sSQLUFP);

				if (Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value) < dtUltimaFechaPago)
				{
					MessageBox.Show(string.Format("La Fecha de Cobro no puede ser menor a '{0}'", dtUltimaFechaPago.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}
			
				
				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Modificar la Fecha?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE DepositosPendientes SET FechaPago = '{0}' WHERE IdDepositosPendientes = {1} ",Convert.ToDateTime(e.Cell.Row.Cells["FechaPago"].Value).ToString("yyyyMMdd"), (int)this.ultraGrid1.ActiveRow.Cells["IdDepositosPendientes"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);	

					bActualiza = false;
					this.btnVer_Click(sender, e);	
				
				}
				else
				{						
					//e.Cell.Row.Cells["FechaPago"].Value = (DateTime)e.Cell.OriginalValue;
					e.Cell.Row.Cells["FechaPago"].Value = Convert.ToDateTime(e.Cell.OriginalValue);
					return;
				}					
				bActualiza = false;
			}
			#endregion FechaPago

			#region Numero Deposito
			if (e.Cell.Column.ToString() == "numeroDeposito")
			{	
				bActualiza = true;
				if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNumeroDepositoPendiente {0}, '{1}',2",(int)this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value,this.ultraGrid1.ActiveRow.Cells["numeroDeposito"].Value.ToString()), true))
				{
					MessageBox.Show("Numero de deposito registrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cell.Row.Cells["numeroDeposito"].Value = e.Cell.OriginalValue.ToString();
					return;					
				}

				if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de el numero de deposito?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{						
					string sSQL = string.Format("UPDATE DepositosPendientes SET numeroDeposito = '{0}' WHERE IdDepositosPendientes = {1} ",e.Cell.Row.Cells["numeroDeposito"].Value.ToString(), (int)this.ultraGrid1.ActiveRow.Cells["IdDepositosPendientes"].Value );
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);	

					bActualiza = false;
					this.btnVer_Click(sender, e);	
				
				}
				else
				{						
					e.Cell.Row.Cells["numeroDeposito"].Value = e.Cell.OriginalValue.ToString();
				}					
				bActualiza = false;
			}
			#endregion Numero Deposito
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.optEstados.Value == 1)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnLiquidar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAnular"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCuenta"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["FechaPago"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.btnVer_Click(sender, e);	
			}
			if ((int)this.optEstados.Value == 2) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnLiquidar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAnular"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCuenta"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["FechaPago"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.btnVer_Click(sender, e);	
			}
			if ((int)this.optEstados.Value == 3) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnLiquidar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAnular"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCuenta"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["FechaPago"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.btnVer_Click(sender, e);	
			}
			if ((int)this.optEstados.Value == 4) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnLiquidar"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["btnAnular"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCuenta"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["FechaPago"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.btnVer_Click(sender, e);	
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
//			if (e.Cell.Row.Cells["FechaPago"].Value != System.DBNull.Value)
//			{
//				if (bActualiza) return;
//				bActualiza = true;
//				if (e.Cell.Column.ToString() == "FechaPago")
//				{		
//					FechaBefore = (DateTime) e.Cell.Row.Cells["FechaPago"].Value ;
//					bActualiza = false;
//				}
//			}
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			int iContador = 0;
			
			try
			{				
				#region Sube Archivo
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
					
					#region Variables
					string sFecha = "";					
					string sTipo	= "";
					string sDocumento = "";
					decimal	dMonto = 0;					
					#endregion Variables

					#region Sube informacion
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{	
						if (row["Fecha"] != System.DBNull.Value) sFecha = row["Fecha"].ToString();
						if (row["Tipo"] != System.DBNull.Value) sTipo = row["Tipo"].ToString();
						if (row["Documento"] != System.DBNull.Value) sDocumento = row["Documento"].ToString();
						if (row["Monto"] != System.DBNull.Value) dMonto = decimal.Parse(row["Monto"].ToString());

						string sSQLAct = string.Format("Exec Cont_EstadoDeCuenta_Inserta {0}, '{1}', '{2}', '{3}', {4}", 
							12, DateTime.Parse(sFecha).ToString("yyyyMMdd"), sTipo, sDocumento, dMonto);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
						iContador++;			
					}
					#endregion Sube informacion

					MessageBox.Show(string.Format("Proceso finalizado correctamente\n\n{0} registros importados correctamente", iContador), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
					this.Cursor = Cursors.Default;
				}			
				#endregion Sube Archivo
			}
			catch (Exception ex)
			{								
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Demo", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;													
			}
		}
	}
}
