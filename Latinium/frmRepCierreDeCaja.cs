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
	/// Descripción breve de frmRepCierreDeCaja.
	/// </summary>
	public class frmRepCierreDeCaja : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago_1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorriente;
		private System.Windows.Forms.Label lblEstadoDepositos;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbVendedores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		int idCaja = 0;
		int iEstadoCaja = 0;
		int iEstado = 0;
		private System.Windows.Forms.ContextMenu mnuExportar;
		private System.Windows.Forms.MenuItem mnuAccionesFactura;
		private System.Windows.Forms.MenuItem mnuAccionesAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCTC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnC_Entrega;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRevisado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCCheque;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnCEfectivo;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmRepCierreDeCaja()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepCierreDeCaja));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("R");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Revisado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RucCedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Voucher");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Registro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago_1", -1, "cmbFormaPago_1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCaja");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTarjetas");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVerTarjeta");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnRetencionVenta");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cartera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 16, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 17, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 19, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Voucher", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Voucher", 20, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 21, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 24, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 22, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 23, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 31, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 31, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 18, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 25, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 26, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 27, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 30, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 30, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 29, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 29, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 28, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 28, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("R");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Revisado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RucCedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Voucher");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Registro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnTarjetas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnVerTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnRetencionVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cartera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 6, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbFormaPago_1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstadoDepositos = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.mnuExportar = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesFactura = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAsiento = new System.Windows.Forms.MenuItem();
			this.spnCTC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnC_Entrega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkRevisado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.spnCCheque = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.spnCEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago_1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCTC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnC_Entrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCEfectivo)).BeginInit();
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
			// cmbFormaPago_1
			// 
			this.cmbFormaPago_1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago_1.DataSource = this.ultraDataSource12;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbFormaPago_1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFormaPago_1.DisplayMember = "Descripcion";
			this.cmbFormaPago_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago_1.Location = new System.Drawing.Point(501, 216);
			this.cmbFormaPago_1.Name = "cmbFormaPago_1";
			this.cmbFormaPago_1.Size = new System.Drawing.Size(272, 88);
			this.cmbFormaPago_1.TabIndex = 761;
			this.cmbFormaPago_1.ValueMember = "idFormaPago_1";
			this.cmbFormaPago_1.Visible = false;
			// 
			// ultraDataSource12
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2});
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(272, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 756;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(1008, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(88, 23);
			this.btnSalir.TabIndex = 752;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(192, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 750;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(-4, 40);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1112, 8);
			this.groupBox2.TabIndex = 749;
			this.groupBox2.TabStop = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 736;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource9;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.Caption = "";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(480, 376);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(224, 86);
			this.cmbFormaPago.TabIndex = 746;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// cmbCuentaCorriente
			// 
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataSource = this.ultraDataSource10;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 250;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuentaCorriente.DisplayMember = "Descripcion";
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(88, 376);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(272, 88);
			this.cmbCuentaCorriente.TabIndex = 745;
			this.cmbCuentaCorriente.ValueMember = "idCuenta";
			this.cmbCuentaCorriente.Visible = false;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn5,
																																	ultraDataColumn6});
			// 
			// lblEstadoDepositos
			// 
			this.lblEstadoDepositos.AutoSize = true;
			this.lblEstadoDepositos.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.lblEstadoDepositos.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstadoDepositos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoDepositos.Location = new System.Drawing.Point(416, 10);
			this.lblEstadoDepositos.Name = "lblEstadoDepositos";
			this.lblEstadoDepositos.Size = new System.Drawing.Size(0, 19);
			this.lblEstadoDepositos.TabIndex = 744;
			this.lblEstadoDepositos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(864, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 738;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 733;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			ultraGridColumn7.CellAppearance = appearance3;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			ultraGridColumn7.CellButtonAppearance = appearance4;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 17;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "I";
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.RowLayoutColumnInfo.OriginX = 0;
			ultraGridColumn8.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn8.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn8.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn8.Width = 17;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.Caption = "T";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 17;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.RowLayoutColumnInfo.OriginX = 2;
			ultraGridColumn10.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn10.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn10.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn10.Width = 90;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.RowLayoutColumnInfo.OriginX = 4;
			ultraGridColumn12.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn12.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn12.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn12.Width = 50;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 36;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 33;
			ultraGridColumn17.Hidden = true;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellAppearance = appearance6;
			ultraGridColumn18.Header.Caption = "N. Documento";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.RowLayoutColumnInfo.OriginX = 6;
			ultraGridColumn18.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn18.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn18.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn18.Width = 85;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "RUC - Cedula";
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.RowLayoutColumnInfo.OriginX = 8;
			ultraGridColumn20.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn20.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn20.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn20.Width = 90;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.RowLayoutColumnInfo.OriginX = 10;
			ultraGridColumn21.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn21.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn21.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn21.Width = 223;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.RowLayoutColumnInfo.OriginX = 12;
			ultraGridColumn22.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn22.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn22.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn22.Width = 75;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance7;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 18;
			ultraGridColumn23.RowLayoutColumnInfo.OriginX = 14;
			ultraGridColumn23.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn23.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn23.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn23.Width = 75;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance8;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.RowLayoutColumnInfo.OriginX = 16;
			ultraGridColumn24.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn24.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn24.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn24.Width = 75;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance9;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "Transf";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.RowLayoutColumnInfo.OriginX = 24;
			ultraGridColumn25.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn25.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn25.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn25.Width = 75;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance10;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "T/Cred";
			ultraGridColumn26.Header.VisiblePosition = 20;
			ultraGridColumn26.RowLayoutColumnInfo.OriginX = 18;
			ultraGridColumn26.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn26.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn26.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn26.Width = 75;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance11;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.RowLayoutColumnInfo.OriginX = 22;
			ultraGridColumn27.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn27.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn27.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn27.Width = 75;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance12;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "N De Cred";
			ultraGridColumn28.Header.VisiblePosition = 22;
			ultraGridColumn28.RowLayoutColumnInfo.OriginX = 22;
			ultraGridColumn28.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn28.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn28.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn28.Width = 75;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance13;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn29.RowLayoutColumnInfo.OriginX = 28;
			ultraGridColumn29.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn29.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn29.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn29.Width = 75;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance14;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "C/Point";
			ultraGridColumn30.Header.VisiblePosition = 27;
			ultraGridColumn30.RowLayoutColumnInfo.OriginX = 30;
			ultraGridColumn30.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn30.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn30.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn30.Width = 75;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance15;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Crédito";
			ultraGridColumn31.Header.VisiblePosition = 26;
			ultraGridColumn31.RowLayoutColumnInfo.OriginX = 32;
			ultraGridColumn31.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn31.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn31.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn31.Width = 75;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance16;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "Retención";
			ultraGridColumn32.Header.VisiblePosition = 28;
			ultraGridColumn32.RowLayoutColumnInfo.OriginX = 34;
			ultraGridColumn32.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn32.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn32.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn32.Width = 75;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance17;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "T/Deb";
			ultraGridColumn33.Header.VisiblePosition = 29;
			ultraGridColumn33.RowLayoutColumnInfo.OriginX = 20;
			ultraGridColumn33.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn33.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn33.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn33.Width = 75;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance18;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 30;
			ultraGridColumn34.RowLayoutColumnInfo.OriginX = 26;
			ultraGridColumn34.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn34.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn34.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn34.Width = 75;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance19;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 31;
			ultraGridColumn35.Width = 75;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance20;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "C/Entrega";
			ultraGridColumn36.Header.VisiblePosition = 32;
			ultraGridColumn36.Width = 85;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance21;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 34;
			ultraGridColumn37.RowLayoutColumnInfo.OriginX = 36;
			ultraGridColumn37.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn37.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn37.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn37.Width = 75;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance22;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 35;
			ultraGridColumn38.RowLayoutColumnInfo.OriginX = 38;
			ultraGridColumn38.RowLayoutColumnInfo.OriginY = 0;
			ultraGridColumn38.RowLayoutColumnInfo.SpanX = 2;
			ultraGridColumn38.RowLayoutColumnInfo.SpanY = 2;
			ultraGridColumn38.Width = 85;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn39.Format = "dd/MM/yyy HH:mm";
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.Width = 120;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn41.Header.Caption = "Vendedor";
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn41.Width = 200;
			ultraGridColumn42.Header.VisiblePosition = 37;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn43.Header.Caption = "Forma De Pago";
			ultraGridColumn43.Header.VisiblePosition = 38;
			ultraGridColumn43.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn43.Width = 201;
			ultraGridColumn44.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance23.Image = ((object)(resources.GetObject("appearance23.Image")));
			appearance23.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance23.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn44.CellButtonAppearance = appearance23;
			ultraGridColumn44.Header.Caption = "...";
			ultraGridColumn44.Header.VisiblePosition = 3;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn44.Width = 25;
			ultraGridColumn45.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance24.Image = ((object)(resources.GetObject("appearance24.Image")));
			appearance24.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance24.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn45.CellButtonAppearance = appearance24;
			ultraGridColumn45.Header.Caption = "...";
			ultraGridColumn45.Header.VisiblePosition = 4;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn45.Width = 25;
			ultraGridColumn46.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance25.Image = ((object)(resources.GetObject("appearance25.Image")));
			appearance25.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellButtonAppearance = appearance25;
			ultraGridColumn46.Header.Caption = "...";
			ultraGridColumn46.Header.VisiblePosition = 21;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn46.Width = 30;
			ultraGridColumn47.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance26.Image = ((object)(resources.GetObject("appearance26.Image")));
			appearance26.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn47.CellButtonAppearance = appearance26;
			ultraGridColumn47.Header.Caption = "R";
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn47.Width = 30;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn48.Header.VisiblePosition = 42;
			ultraGridColumn48.Width = 160;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn49.Header.VisiblePosition = 6;
			ultraGridColumn49.Width = 115;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49});
			appearance27.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand4.Header.Appearance = appearance27;
			ultraGridBand4.Header.Caption = "MOVIMIENTOS DE CAJA";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
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
			summarySettings4.Hidden = true;
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
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance39;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance40;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance41;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance42;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance43;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings16});
			ultraGridBand4.SummaryFooterCaption = "Total Movimientos de Caja";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance44;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance45.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance45;
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance46;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance47.FontData.BoldAsString = "True";
			appearance47.FontData.Name = "Arial";
			appearance47.FontData.SizeInPoints = 8F;
			appearance47.ForeColor = System.Drawing.Color.White;
			appearance47.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance47;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 7.9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1088, 296);
			this.ultraGrid1.TabIndex = 735;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn9.DataType = typeof(bool);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.DateTime);
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
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.DateTime);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49});
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource8;
			appearance50.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance50;
			this.uGridDepositos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 78;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn51.Header.VisiblePosition = 1;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 54;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn52.Header.Caption = "Forma de pago";
			ultraGridColumn52.Header.VisiblePosition = 2;
			ultraGridColumn52.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn52.Width = 254;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn53.Header.Caption = "Banco";
			ultraGridColumn53.Header.VisiblePosition = 3;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn53.Width = 352;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn54.Header.Caption = "Número";
			ultraGridColumn54.Header.VisiblePosition = 4;
			ultraGridColumn54.Width = 172;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn55.Header.Caption = "Fecha";
			ultraGridColumn55.Header.VisiblePosition = 5;
			ultraGridColumn55.Width = 126;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance51;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 6;
			ultraGridColumn56.PromptChar = ' ';
			ultraGridColumn56.Width = 146;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn57.Header.VisiblePosition = 7;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 56;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57});
			ultraGridBand5.Header.Caption = "Depositos";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance52;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings17});
			ultraGridBand5.SummaryFooterCaption = "Total Depositos";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDepositos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance53.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance53;
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDepositos.DisplayLayout.Override.CellAppearance = appearance54;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance55.FontData.BoldAsString = "True";
			appearance55.FontData.Name = "Arial";
			appearance55.FontData.SizeInPoints = 8F;
			appearance55.ForeColor = System.Drawing.Color.White;
			appearance55.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance57;
			this.uGridDepositos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(8, 360);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(1088, 152);
			this.uGridDepositos.TabIndex = 747;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(System.DateTime);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn56.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn57.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn50,
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57});
			// 
			// cmbVendedores
			// 
			this.cmbVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedores.DataSource = this.ultraDataSource11;
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 1;
			ultraGridColumn59.Width = 200;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn58,
																										 ultraGridColumn59});
			this.cmbVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbVendedores.DisplayMember = "Nombre";
			this.cmbVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbVendedores.Location = new System.Drawing.Point(1020, 200);
			this.cmbVendedores.Name = "cmbVendedores";
			this.cmbVendedores.Size = new System.Drawing.Size(216, 104);
			this.cmbVendedores.TabIndex = 758;
			this.cmbVendedores.ValueMember = "idPersonal";
			this.cmbVendedores.Visible = false;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn58.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn58,
																																	ultraDataColumn59});
			// 
			// mnuExportar
			// 
			this.mnuExportar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesFactura,
																																								this.mnuAccionesAsiento});
			// 
			// mnuAccionesFactura
			// 
			this.mnuAccionesFactura.Index = 0;
			this.mnuAccionesFactura.Text = "Movimientos de Caja";
			this.mnuAccionesFactura.Click += new System.EventHandler(this.mnuAccionesFactura_Click);
			// 
			// mnuAccionesAsiento
			// 
			this.mnuAccionesAsiento.Index = 1;
			this.mnuAccionesAsiento.Text = "Depositos";
			this.mnuAccionesAsiento.Click += new System.EventHandler(this.mnuAccionesAsiento_Click);
			// 
			// spnCTC
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCTC.Appearance = appearance58;
			this.spnCTC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCTC.Enabled = false;
			this.spnCTC.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnCTC.FormatString = "#,##0.00";
			this.spnCTC.Location = new System.Drawing.Point(656, 8);
			this.spnCTC.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnCTC.Name = "spnCTC";
			this.spnCTC.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCTC.PromptChar = ' ';
			this.spnCTC.Size = new System.Drawing.Size(60, 21);
			this.spnCTC.TabIndex = 767;
			this.spnCTC.Visible = false;
			// 
			// spnC_Entrega
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnC_Entrega.Appearance = appearance59;
			this.spnC_Entrega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnC_Entrega.Enabled = false;
			this.spnC_Entrega.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnC_Entrega.FormatString = "#,##0.00";
			this.spnC_Entrega.Location = new System.Drawing.Point(792, 8);
			this.spnC_Entrega.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnC_Entrega.Name = "spnC_Entrega";
			this.spnC_Entrega.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnC_Entrega.PromptChar = ' ';
			this.spnC_Entrega.Size = new System.Drawing.Size(60, 21);
			this.spnC_Entrega.TabIndex = 766;
			this.spnC_Entrega.Visible = false;
			// 
			// chkRevisado
			// 
			this.chkRevisado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkRevisado.Appearance = appearance60;
			this.chkRevisado.BackColor = System.Drawing.Color.Transparent;
			this.chkRevisado.Location = new System.Drawing.Point(656, 8);
			this.chkRevisado.Name = "chkRevisado";
			this.chkRevisado.Size = new System.Drawing.Size(72, 22);
			this.chkRevisado.TabIndex = 765;
			this.chkRevisado.Text = "Revisado";
			this.chkRevisado.Visible = false;
			// 
			// spnCCheque
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCCheque.Appearance = appearance61;
			this.spnCCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCCheque.Enabled = false;
			this.spnCCheque.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnCCheque.FormatString = "#,##0.00";
			this.spnCCheque.Location = new System.Drawing.Point(800, 8);
			this.spnCCheque.Name = "spnCCheque";
			this.spnCCheque.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCCheque.PromptChar = ' ';
			this.spnCCheque.Size = new System.Drawing.Size(60, 21);
			this.spnCCheque.TabIndex = 764;
			this.spnCCheque.Visible = false;
			// 
			// spnCEfectivo
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.spnCEfectivo.Appearance = appearance62;
			this.spnCEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnCEfectivo.Enabled = false;
			this.spnCEfectivo.Font = new System.Drawing.Font("Tahoma", 7.75F);
			this.spnCEfectivo.FormatString = "#,##0.00";
			this.spnCEfectivo.Location = new System.Drawing.Point(864, 8);
			this.spnCEfectivo.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.spnCEfectivo.Name = "spnCEfectivo";
			this.spnCEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnCEfectivo.PromptChar = ' ';
			this.spnCEfectivo.Size = new System.Drawing.Size(60, 21);
			this.spnCEfectivo.TabIndex = 763;
			this.spnCEfectivo.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.label1.ForeColor = System.Drawing.Color.Firebrick;
			this.label1.Location = new System.Drawing.Point(688, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(0, 23);
			this.label1.TabIndex = 762;
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmRepCierreDeCaja
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1104, 518);
			this.Controls.Add(this.spnCTC);
			this.Controls.Add(this.spnC_Entrega);
			this.Controls.Add(this.chkRevisado);
			this.Controls.Add(this.spnCCheque);
			this.Controls.Add(this.spnCEfectivo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFormaPago_1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Controls.Add(this.lblEstadoDepositos);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.uGridDepositos);
			this.Controls.Add(this.cmbVendedores);
			this.MaximizeBox = false;
			this.Name = "frmRepCierreDeCaja";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reporte Cierre Caja";
			this.Load += new System.EventHandler(this.frmRepCierreDeCaja_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago_1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCTC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnC_Entrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnCEfectivo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		string strHostName = string.Empty;

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmRepCierreDeCaja_Load(object sender, System.EventArgs e)
		{

			miAcceso = new Acceso(cdsSeteoF, "044405");

			if (!Funcion.Permiso("1704", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al Reporte Cierre de Caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}
			if (miAcceso.BExportarExcel) this.btnExportar.Enabled = true;

			Funcion.Resolucion(this, this.Height, this.Width);
		
			strHostName = System.Environment.MachineName;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridDepositos);
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			FuncionesProcedimientos.FechasDocumentos(this.dtFecha);
			this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Select C.idCuenta, c.Descripcion From BodegaCuentasBancarias B INNER JOIN Cuenta C ON C.idCuenta = B.idCuenta Group by c.idCuenta, c.Descripcion"));
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2, 3, 4, 16, 19) Order By idFormaPago");
			this.cmbFormaPago_1.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("SELECT idFormaPago AS idFormaPago_1, FormaPago AS  Descripcion FROM CompraFormaPago"));
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			idCaja = 0;
			iEstado = 0;
			iEstadoCaja = 0;

			this.lblEstado.Text = "";
			this.lblEstadoDepositos.Text = "";
			
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

			this.Cursor = Cursors.WaitCursor;

			string sSQL = string.Format("Exec RepCierreDeCaja '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);

			this.lblEstadoDepositos.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
			
			#region Encabezado y Detalle
			string sSQLMaestro = string.Format("Exec RepConsultaCaja '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			SqlDataReader drM = Funcion.rEscalarSQL(cdsSeteoF, sSQLMaestro, true);
			drM.Read();
			if (drM.HasRows)
			{
				idCaja = (int)drM.GetValue(0);				
			}
			drM.Close();

			this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RepConsultaDetalleCaja '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			#endregion Encabezado y Detalle

			this.Cursor = Cursors.Default;
		}

		private void mnuAccionesFactura_Click(object sender, System.EventArgs e)
		{			
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void mnuAccionesAsiento_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDepositos);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			this.mnuExportar.Show(this.btnExportar, new Point(1, 21));
		}
	}
}
