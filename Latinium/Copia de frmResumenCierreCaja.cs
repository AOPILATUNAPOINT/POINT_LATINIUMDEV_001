using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmResumenCierreCaja.
	/// </summary>
	public class frmResumenCierreCaja : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCaja;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorriente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTurno;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotVEnta;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuadre;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTotal;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMonedas;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB1;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB5;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB10;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB20;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB50;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB100;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB100;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGasto;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVoucher;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEfectivo;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Label lblContador;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Label lblEstadoDepositos;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnDeposito;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDeposito;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.GroupBox groupBox8;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmResumenCierreCaja()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmResumenCierreCaja));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Habitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Consumo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Prenda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CompraYA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habitacion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Consumo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TC");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TD");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prenda");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CompraYA");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCaja", 0);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Habitacion", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Habitacion", 14, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TC", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TC", 17, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TD", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TD", 18, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Prenda", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Prenda", 20, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 23, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Consumo", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Consumo", 15, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 19, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CompraYA", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CompraYA", 21, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento", 22, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 16, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CompraYA", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CompraYA", 21, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 6, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			this.label19 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridCaja = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optTurno = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label24 = new System.Windows.Forms.Label();
			this.txtDescuadre = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.txtVTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtMonedas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.txtVTB1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantB1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtVTB5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantB5 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtVTB10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantB10 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtVTB20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantB20 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtVTB50 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCantB50 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtVTB100 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtCantB100 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtGasto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtVoucher = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtTotVEnta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.groupBox8 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnDeposito = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.lblEstadoDepositos = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurno)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuadre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMonedas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB50)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB50)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB100)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB100)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGasto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotVEnta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(264, 43);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(36, 17);
			this.label19.TabIndex = 280;
			this.label19.Text = "Cajero";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 43);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(26, 17);
			this.lblBodega.TabIndex = 279;
			this.lblBodega.Text = "Caja";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCajero
			// 
			appearance1.FontData.Name = "Tahoma";
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance1;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(320, 40);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(220, 22);
			this.cmbCajero.TabIndex = 281;
			this.cmbCajero.ValueMember = "idPersonal";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 278;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance2.FontData.Name = "Tahoma";
			this.cmbBodega.Appearance = appearance2;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 30;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 30;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 180;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 30;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 276;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 277;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn14.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24});
			// 
			// uGridCaja
			// 
			this.uGridCaja.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCaja.DataSource = this.ultraDataSource6;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridCaja.DisplayLayout.Appearance = appearance4;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "I";
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 20;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Width = 100;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn15.Header.VisiblePosition = 8;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 9;
			ultraGridColumn17.Width = 70;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 11;
			ultraGridColumn18.Width = 75;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance5;
			ultraGridColumn19.Header.Caption = "N.";
			ultraGridColumn19.Header.VisiblePosition = 12;
			ultraGridColumn19.Width = 40;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn20.CellAppearance = appearance6;
			ultraGridColumn20.Header.VisiblePosition = 13;
			ultraGridColumn20.Width = 40;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn21.Header.VisiblePosition = 24;
			ultraGridColumn21.Width = 100;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance7;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance8;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 15;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance9;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 17;
			ultraGridColumn24.Width = 80;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance10;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "T/Credito";
			ultraGridColumn25.Header.VisiblePosition = 18;
			ultraGridColumn25.Width = 80;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance11;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "T/Debito";
			ultraGridColumn26.Header.VisiblePosition = 19;
			ultraGridColumn26.Width = 80;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance12;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 20;
			ultraGridColumn27.Width = 80;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance13;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 21;
			ultraGridColumn28.Width = 80;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance14;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Compra Ya";
			ultraGridColumn29.Header.VisiblePosition = 22;
			ultraGridColumn29.Width = 80;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance15;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 23;
			ultraGridColumn30.Width = 80;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance16.BackColor2 = System.Drawing.Color.LightSteelBlue;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance16;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 16;
			ultraGridColumn31.Width = 80;
			ultraGridColumn32.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.Image = ((object)(resources.GetObject("appearance17.Image")));
			appearance17.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance17.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn32.CellButtonAppearance = appearance17;
			ultraGridColumn32.Header.Caption = "...";
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn32.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn32});
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance18;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance19;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance20;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance21;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance22;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance23;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance24;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance25;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.Hidden = true;
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance26;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance27;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance28;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
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
																																															summarySettings11});
			ultraGridBand3.SummaryFooterCaption = "Total Caja";
			this.uGridCaja.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridCaja.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCaja.DisplayLayout.Override.ActiveRowAppearance = appearance29;
			this.uGridCaja.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCaja.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCaja.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance30.BackColor = System.Drawing.Color.Transparent;
			this.uGridCaja.DisplayLayout.Override.CardAreaAppearance = appearance30;
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCaja.DisplayLayout.Override.CellAppearance = appearance31;
			this.uGridCaja.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridCaja.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.FontData.BoldAsString = "True";
			appearance32.FontData.Name = "Arial";
			appearance32.FontData.SizeInPoints = 8F;
			appearance32.ForeColor = System.Drawing.Color.White;
			appearance32.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCaja.DisplayLayout.Override.HeaderAppearance = appearance32;
			this.uGridCaja.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCaja.DisplayLayout.Override.RowSelectorAppearance = appearance33;
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance34.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCaja.DisplayLayout.Override.SelectedRowAppearance = appearance34;
			this.uGridCaja.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridCaja.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridCaja.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.uGridCaja.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCaja.Location = new System.Drawing.Point(8, 80);
			this.uGridCaja.Name = "uGridCaja";
			this.uGridCaja.Size = new System.Drawing.Size(900, 232);
			this.uGridCaja.TabIndex = 282;
			this.uGridCaja.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCaja_CellChange);
			this.uGridCaja.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCaja_KeyDown);
			this.uGridCaja.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCaja_ClickCellButton);
			this.uGridCaja.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridCaja_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(System.DateTime);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource4;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.Caption = "";
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 200;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(8, 344);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(224, 86);
			this.cmbFormaPago.TabIndex = 285;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource3;
			appearance35.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance35;
			this.uGridDepositos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 78;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 54;
			ultraGridColumn37.Header.Caption = "Forma de Pago";
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn37.Width = 178;
			ultraGridColumn38.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn38.Header.VisiblePosition = 3;
			ultraGridColumn38.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn38.Width = 235;
			ultraGridColumn39.Header.Caption = "Número de Deposito";
			ultraGridColumn39.Header.VisiblePosition = 4;
			ultraGridColumn39.Width = 188;
			ultraGridColumn40.Header.Caption = "Fecha";
			ultraGridColumn40.Header.VisiblePosition = 5;
			ultraGridColumn40.Width = 100;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance36;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 6;
			ultraGridColumn41.PromptChar = ' ';
			ultraGridColumn41.Width = 94;
			ultraGridColumn42.Header.VisiblePosition = 7;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 56;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42});
			ultraGridBand5.Header.Caption = "Depositos";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance37;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings12});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance38.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance38;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 8F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.uGridDepositos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(8, 312);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(833, 135);
			this.uGridDepositos.TabIndex = 283;
			this.uGridDepositos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDepositos_AfterRowInsert);
			this.uGridDepositos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDepositos_KeyPress);
			this.uGridDepositos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDepositos_BeforeRowsDeleted);
			this.uGridDepositos.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.uGridDepositos_BeforeRowInsert);
			this.uGridDepositos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDepositos_KeyDown);
			this.uGridDepositos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDepositos_InitializeLayout);
			// 
			// cmbCuentaCorriente
			// 
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataSource = this.ultraDataSource5;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.Caption = "";
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 250;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn43,
																										 ultraGridColumn44});
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCuentaCorriente.DisplayMember = "Descripcion";
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(528, 344);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(272, 88);
			this.cmbCuentaCorriente.TabIndex = 284;
			this.cmbCuentaCorriente.ValueMember = "idCuenta";
			this.cmbCuentaCorriente.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn35,
																																 ultraDataColumn36});
			// 
			// optTurno
			// 
			appearance42.FontData.Name = "Tahoma";
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTurno.Appearance = appearance42;
			this.optTurno.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTurno.CheckedIndex = 1;
			this.optTurno.ItemAppearance = appearance43;
			this.optTurno.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Turno 1";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Turno 2";
			this.optTurno.Items.Add(valueListItem1);
			this.optTurno.Items.Add(valueListItem2);
			this.optTurno.Items.Add(valueListItem3);
			this.optTurno.ItemSpacingVertical = 10;
			this.optTurno.Location = new System.Drawing.Point(192, 6);
			this.optTurno.Name = "optTurno";
			this.optTurno.Size = new System.Drawing.Size(216, 24);
			this.optTurno.TabIndex = 704;
			this.optTurno.Text = "Turno 1";
			this.optTurno.ValueChanged += new System.EventHandler(this.optTurno_ValueChanged);
			// 
			// btnImprimir
			// 
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(1008, 488);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 24);
			this.btnImprimir.TabIndex = 734;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1104, 488);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 733;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(552, 39);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(88, 24);
			this.btnBuscar.TabIndex = 732;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.txtDeposito);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.groupBox4);
			this.groupBox1.Controls.Add(this.groupBox3);
			this.groupBox1.Controls.Add(this.label24);
			this.groupBox1.Controls.Add(this.txtDescuadre);
			this.groupBox1.Controls.Add(this.label21);
			this.groupBox1.Controls.Add(this.txtVTotal);
			this.groupBox1.Controls.Add(this.label20);
			this.groupBox1.Controls.Add(this.txtMonedas);
			this.groupBox1.Controls.Add(this.label18);
			this.groupBox1.Controls.Add(this.txtVTB1);
			this.groupBox1.Controls.Add(this.txtCantB1);
			this.groupBox1.Controls.Add(this.label17);
			this.groupBox1.Controls.Add(this.txtVTB5);
			this.groupBox1.Controls.Add(this.txtCantB5);
			this.groupBox1.Controls.Add(this.label16);
			this.groupBox1.Controls.Add(this.txtVTB10);
			this.groupBox1.Controls.Add(this.txtCantB10);
			this.groupBox1.Controls.Add(this.label15);
			this.groupBox1.Controls.Add(this.txtVTB20);
			this.groupBox1.Controls.Add(this.txtCantB20);
			this.groupBox1.Controls.Add(this.label14);
			this.groupBox1.Controls.Add(this.txtVTB50);
			this.groupBox1.Controls.Add(this.txtCantB50);
			this.groupBox1.Controls.Add(this.label13);
			this.groupBox1.Controls.Add(this.txtVTB100);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtCantB100);
			this.groupBox1.Controls.Add(this.label11);
			this.groupBox1.Controls.Add(this.txtGasto);
			this.groupBox1.Controls.Add(this.label10);
			this.groupBox1.Controls.Add(this.groupBox2);
			this.groupBox1.Controls.Add(this.txtEfectivo);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.txtVoucher);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.txtTotVEnta);
			this.groupBox1.Controls.Add(this.label22);
			this.groupBox1.Controls.Add(this.groupBox7);
			this.groupBox1.Controls.Add(this.groupBox8);
			this.groupBox1.Location = new System.Drawing.Point(920, 0);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(272, 472);
			this.groupBox1.TabIndex = 735;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Resumén";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// txtDeposito
			// 
			appearance44.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			appearance44.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtDeposito.Appearance = appearance44;
			this.txtDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDeposito.Enabled = false;
			this.txtDeposito.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDeposito.FormatString = "#,##0.00";
			this.txtDeposito.Location = new System.Drawing.Point(176, 296);
			this.txtDeposito.Name = "txtDeposito";
			this.txtDeposito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDeposito.PromptChar = ' ';
			this.txtDeposito.Size = new System.Drawing.Size(80, 22);
			this.txtDeposito.TabIndex = 751;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.label1.Location = new System.Drawing.Point(32, 299);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(74, 17);
			this.label1.TabIndex = 750;
			this.label1.Text = "(-) DEPOSITO";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox4.Location = new System.Drawing.Point(32, 344);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(232, 8);
			this.groupBox4.TabIndex = 749;
			this.groupBox4.TabStop = false;
			// 
			// groupBox3
			// 
			this.groupBox3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(32, 280);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(232, 8);
			this.groupBox3.TabIndex = 748;
			this.groupBox3.TabStop = false;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.ForeColor = System.Drawing.Color.Firebrick;
			this.label24.Location = new System.Drawing.Point(48, 64);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(52, 17);
			this.label24.TabIndex = 303;
			this.label24.Text = "BILLETES";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescuadre
			// 
			appearance45.ForeColor = System.Drawing.Color.Firebrick;
			appearance45.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtDescuadre.Appearance = appearance45;
			this.txtDescuadre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuadre.Enabled = false;
			this.txtDescuadre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescuadre.FormatString = "#,##0.00";
			this.txtDescuadre.Location = new System.Drawing.Point(176, 360);
			this.txtDescuadre.Name = "txtDescuadre";
			this.txtDescuadre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuadre.PromptChar = ' ';
			this.txtDescuadre.Size = new System.Drawing.Size(80, 22);
			this.txtDescuadre.TabIndex = 298;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.ForeColor = System.Drawing.Color.Firebrick;
			this.label21.Location = new System.Drawing.Point(32, 363);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(87, 17);
			this.label21.TabIndex = 297;
			this.label21.Text = "(=) DESCUADRE";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTotal
			// 
			appearance46.ForeColor = System.Drawing.Color.Firebrick;
			appearance46.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtVTotal.Appearance = appearance46;
			this.txtVTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTotal.Enabled = false;
			this.txtVTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTotal.FormatString = "#,##0.00";
			this.txtVTotal.Location = new System.Drawing.Point(176, 256);
			this.txtVTotal.Name = "txtVTotal";
			this.txtVTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTotal.PromptChar = ' ';
			this.txtVTotal.Size = new System.Drawing.Size(80, 22);
			this.txtVTotal.TabIndex = 296;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.ForeColor = System.Drawing.Color.Firebrick;
			this.label20.Location = new System.Drawing.Point(48, 259);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(92, 17);
			this.label20.TabIndex = 294;
			this.label20.Text = "TOTAL EFECTIVO";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMonedas
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtMonedas.Appearance = appearance47;
			this.txtMonedas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMonedas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtMonedas.FormatString = "#,##0.00";
			this.txtMonedas.Location = new System.Drawing.Point(176, 232);
			this.txtMonedas.Name = "txtMonedas";
			this.txtMonedas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMonedas.PromptChar = ' ';
			this.txtMonedas.Size = new System.Drawing.Size(80, 22);
			this.txtMonedas.TabIndex = 293;
			this.txtMonedas.Click += new System.EventHandler(this.txtMonedas_Click);
			this.txtMonedas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMonedas_KeyPress);
			this.txtMonedas.Validated += new System.EventHandler(this.txtMonedas_Validated);
			this.txtMonedas.Enter += new System.EventHandler(this.txtMonedas_Enter);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.ForeColor = System.Drawing.Color.Black;
			this.label18.Location = new System.Drawing.Point(48, 235);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(49, 17);
			this.label18.TabIndex = 291;
			this.label18.Text = "Monedas";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTB1
			// 
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB1.Appearance = appearance48;
			this.txtVTB1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB1.Enabled = false;
			this.txtVTB1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTB1.FormatString = "#,##0.00";
			this.txtVTB1.Location = new System.Drawing.Point(176, 208);
			this.txtVTB1.Name = "txtVTB1";
			this.txtVTB1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB1.PromptChar = ' ';
			this.txtVTB1.Size = new System.Drawing.Size(80, 22);
			this.txtVTB1.TabIndex = 290;
			// 
			// txtCantB1
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB1.Appearance = appearance49;
			this.txtCantB1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantB1.FormatString = "##0";
			this.txtCantB1.Location = new System.Drawing.Point(104, 208);
			this.txtCantB1.Name = "txtCantB1";
			this.txtCantB1.PromptChar = ' ';
			this.txtCantB1.Size = new System.Drawing.Size(64, 22);
			this.txtCantB1.TabIndex = 289;
			this.txtCantB1.Click += new System.EventHandler(this.txtCantB1_Click);
			this.txtCantB1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB1_KeyPress);
			this.txtCantB1.Validated += new System.EventHandler(this.txtCantB1_Validated);
			this.txtCantB1.Enter += new System.EventHandler(this.txtCantB1_Enter);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(48, 211);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(20, 17);
			this.label17.TabIndex = 288;
			this.label17.Text = "B.1";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTB5
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB5.Appearance = appearance50;
			this.txtVTB5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB5.Enabled = false;
			this.txtVTB5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTB5.FormatString = "#,##0.00";
			this.txtVTB5.Location = new System.Drawing.Point(176, 184);
			this.txtVTB5.Name = "txtVTB5";
			this.txtVTB5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB5.PromptChar = ' ';
			this.txtVTB5.Size = new System.Drawing.Size(80, 22);
			this.txtVTB5.TabIndex = 287;
			// 
			// txtCantB5
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB5.Appearance = appearance51;
			this.txtCantB5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantB5.FormatString = "##0";
			this.txtCantB5.Location = new System.Drawing.Point(104, 184);
			this.txtCantB5.Name = "txtCantB5";
			this.txtCantB5.PromptChar = ' ';
			this.txtCantB5.Size = new System.Drawing.Size(64, 22);
			this.txtCantB5.TabIndex = 286;
			this.txtCantB5.Click += new System.EventHandler(this.txtCantB5_Click);
			this.txtCantB5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB5_KeyPress);
			this.txtCantB5.Validated += new System.EventHandler(this.txtCantB5_Validated);
			this.txtCantB5.Enter += new System.EventHandler(this.txtCantB5_Enter);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(48, 187);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(20, 17);
			this.label16.TabIndex = 285;
			this.label16.Text = "B.5";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTB10
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB10.Appearance = appearance52;
			this.txtVTB10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB10.Enabled = false;
			this.txtVTB10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTB10.FormatString = "#,##0.00";
			this.txtVTB10.Location = new System.Drawing.Point(176, 160);
			this.txtVTB10.Name = "txtVTB10";
			this.txtVTB10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB10.PromptChar = ' ';
			this.txtVTB10.Size = new System.Drawing.Size(80, 22);
			this.txtVTB10.TabIndex = 284;
			// 
			// txtCantB10
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB10.Appearance = appearance53;
			this.txtCantB10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantB10.FormatString = "##0";
			this.txtCantB10.Location = new System.Drawing.Point(104, 160);
			this.txtCantB10.Name = "txtCantB10";
			this.txtCantB10.PromptChar = ' ';
			this.txtCantB10.Size = new System.Drawing.Size(64, 22);
			this.txtCantB10.TabIndex = 283;
			this.txtCantB10.Click += new System.EventHandler(this.txtCantB10_Click);
			this.txtCantB10.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB10_KeyPress);
			this.txtCantB10.Validated += new System.EventHandler(this.txtCantB10_Validated);
			this.txtCantB10.Enter += new System.EventHandler(this.txtCantB10_Enter);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(48, 163);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(27, 17);
			this.label15.TabIndex = 282;
			this.label15.Text = "B.10";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTB20
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB20.Appearance = appearance54;
			this.txtVTB20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB20.Enabled = false;
			this.txtVTB20.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTB20.FormatString = "#,##0.00";
			this.txtVTB20.Location = new System.Drawing.Point(176, 136);
			this.txtVTB20.Name = "txtVTB20";
			this.txtVTB20.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB20.PromptChar = ' ';
			this.txtVTB20.Size = new System.Drawing.Size(80, 22);
			this.txtVTB20.TabIndex = 281;
			// 
			// txtCantB20
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB20.Appearance = appearance55;
			this.txtCantB20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB20.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantB20.FormatString = "##0";
			this.txtCantB20.Location = new System.Drawing.Point(104, 136);
			this.txtCantB20.Name = "txtCantB20";
			this.txtCantB20.PromptChar = ' ';
			this.txtCantB20.Size = new System.Drawing.Size(64, 22);
			this.txtCantB20.TabIndex = 280;
			this.txtCantB20.Click += new System.EventHandler(this.txtCantB20_Click);
			this.txtCantB20.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB20_KeyPress);
			this.txtCantB20.Validated += new System.EventHandler(this.txtCantB20_Validated);
			this.txtCantB20.Enter += new System.EventHandler(this.txtCantB20_Enter);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(48, 139);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(27, 17);
			this.label14.TabIndex = 279;
			this.label14.Text = "B.20";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTB50
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB50.Appearance = appearance56;
			this.txtVTB50.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB50.Enabled = false;
			this.txtVTB50.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTB50.FormatString = "#,##0.00";
			this.txtVTB50.Location = new System.Drawing.Point(176, 112);
			this.txtVTB50.Name = "txtVTB50";
			this.txtVTB50.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB50.PromptChar = ' ';
			this.txtVTB50.Size = new System.Drawing.Size(80, 22);
			this.txtVTB50.TabIndex = 278;
			// 
			// txtCantB50
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB50.Appearance = appearance57;
			this.txtCantB50.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB50.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantB50.FormatString = "##0";
			this.txtCantB50.Location = new System.Drawing.Point(104, 112);
			this.txtCantB50.Name = "txtCantB50";
			this.txtCantB50.PromptChar = ' ';
			this.txtCantB50.Size = new System.Drawing.Size(64, 22);
			this.txtCantB50.TabIndex = 277;
			this.txtCantB50.Click += new System.EventHandler(this.txtCantB50_Click);
			this.txtCantB50.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB50_KeyPress);
			this.txtCantB50.Validated += new System.EventHandler(this.txtCantB50_Validated);
			this.txtCantB50.Enter += new System.EventHandler(this.txtCantB50_Enter);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(48, 115);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(27, 17);
			this.label13.TabIndex = 276;
			this.label13.Text = "B.50";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTB100
			// 
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB100.Appearance = appearance58;
			this.txtVTB100.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB100.Enabled = false;
			this.txtVTB100.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTB100.FormatString = "#,##0.00";
			this.txtVTB100.Location = new System.Drawing.Point(176, 88);
			this.txtVTB100.Name = "txtVTB100";
			this.txtVTB100.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB100.PromptChar = ' ';
			this.txtVTB100.Size = new System.Drawing.Size(80, 22);
			this.txtVTB100.TabIndex = 275;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.ForeColor = System.Drawing.Color.Firebrick;
			this.label12.Location = new System.Drawing.Point(176, 64);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(38, 17);
			this.label12.TabIndex = 274;
			this.label12.Text = "VALOR";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCantB100
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB100.Appearance = appearance59;
			this.txtCantB100.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB100.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCantB100.FormatString = "##0";
			this.txtCantB100.Location = new System.Drawing.Point(104, 88);
			this.txtCantB100.Name = "txtCantB100";
			this.txtCantB100.PromptChar = ' ';
			this.txtCantB100.Size = new System.Drawing.Size(64, 22);
			this.txtCantB100.TabIndex = 273;
			this.txtCantB100.Click += new System.EventHandler(this.txtCantB100_Click);
			this.txtCantB100.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB100_KeyPress);
			this.txtCantB100.Validated += new System.EventHandler(this.txtCantB100_Validated);
			this.txtCantB100.Enter += new System.EventHandler(this.txtCantB100_Enter);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(48, 91);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(33, 17);
			this.label11.TabIndex = 272;
			this.label11.Text = "B.100";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtGasto
			// 
			appearance60.ForeColor = System.Drawing.Color.Firebrick;
			appearance60.ForeColorDisabled = System.Drawing.Color.Firebrick;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtGasto.Appearance = appearance60;
			this.txtGasto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGasto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtGasto.FormatString = "#,##0.00";
			this.txtGasto.Location = new System.Drawing.Point(176, 320);
			this.txtGasto.Name = "txtGasto";
			this.txtGasto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGasto.PromptChar = ' ';
			this.txtGasto.Size = new System.Drawing.Size(80, 22);
			this.txtGasto.TabIndex = 271;
			this.txtGasto.Click += new System.EventHandler(this.txtGasto_Click);
			this.txtGasto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGasto_KeyPress);
			this.txtGasto.Validated += new System.EventHandler(this.txtGasto_Validated);
			this.txtGasto.Enter += new System.EventHandler(this.txtGasto_Enter);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.ForeColor = System.Drawing.Color.Firebrick;
			this.label10.Location = new System.Drawing.Point(32, 323);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(56, 17);
			this.label10.TabIndex = 270;
			this.label10.Text = "(-) GASTO";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox2
			// 
			this.groupBox2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(16, 48);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(248, 8);
			this.groupBox2.TabIndex = 747;
			this.groupBox2.TabStop = false;
			// 
			// txtEfectivo
			// 
			appearance61.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			appearance61.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.txtEfectivo.Appearance = appearance61;
			this.txtEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEfectivo.Enabled = false;
			this.txtEfectivo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEfectivo.FormatString = "#,##0.00";
			this.txtEfectivo.Location = new System.Drawing.Point(176, 24);
			this.txtEfectivo.Name = "txtEfectivo";
			this.txtEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEfectivo.PromptChar = ' ';
			this.txtEfectivo.Size = new System.Drawing.Size(80, 22);
			this.txtEfectivo.TabIndex = 267;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.label8.Location = new System.Drawing.Point(16, 27);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(76, 17);
			this.label8.TabIndex = 266;
			this.label8.Text = "(+) EFECTIVO";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVoucher
			// 
			appearance62.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			appearance62.ForeColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.txtVoucher.Appearance = appearance62;
			this.txtVoucher.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVoucher.Enabled = false;
			this.txtVoucher.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVoucher.FormatString = "#,##0.00";
			this.txtVoucher.Location = new System.Drawing.Point(176, 400);
			this.txtVoucher.Name = "txtVoucher";
			this.txtVoucher.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVoucher.PromptChar = ' ';
			this.txtVoucher.Size = new System.Drawing.Size(80, 22);
			this.txtVoucher.TabIndex = 269;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(192)));
			this.label9.Location = new System.Drawing.Point(16, 403);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(75, 17);
			this.label9.TabIndex = 268;
			this.label9.Text = "(+) VOUCHER";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotVEnta
			// 
			appearance63.ForeColor = System.Drawing.Color.Firebrick;
			appearance63.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotVEnta.Appearance = appearance63;
			this.txtTotVEnta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotVEnta.Enabled = false;
			this.txtTotVEnta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotVEnta.FormatString = "#,##0.00";
			this.txtTotVEnta.Location = new System.Drawing.Point(176, 440);
			this.txtTotVEnta.Name = "txtTotVEnta";
			this.txtTotVEnta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotVEnta.PromptChar = ' ';
			this.txtTotVEnta.Size = new System.Drawing.Size(80, 22);
			this.txtTotVEnta.TabIndex = 300;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.ForeColor = System.Drawing.Color.Firebrick;
			this.label22.Location = new System.Drawing.Point(16, 443);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(59, 17);
			this.label22.TabIndex = 299;
			this.label22.Text = "(=) VENTA";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox7
			// 
			this.groupBox7.Location = new System.Drawing.Point(16, 424);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(248, 8);
			this.groupBox7.TabIndex = 753;
			this.groupBox7.TabStop = false;
			// 
			// groupBox8
			// 
			this.groupBox8.Location = new System.Drawing.Point(32, 384);
			this.groupBox8.Name = "groupBox8";
			this.groupBox8.Size = new System.Drawing.Size(232, 8);
			this.groupBox8.TabIndex = 753;
			this.groupBox8.TabStop = false;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 478);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 17);
			this.label4.TabIndex = 740;
			this.label4.Text = "Notas";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance64;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(64, 464);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(580, 44);
			this.txtObservaciones.TabIndex = 739;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(432, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 745;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnDeposito
			// 
			this.btnDeposito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDeposito.BackColor = System.Drawing.SystemColors.Control;
			this.btnDeposito.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnDeposito.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnDeposito.Image = ((System.Drawing.Image)(resources.GetObject("btnDeposito.Image")));
			this.btnDeposito.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDeposito.Location = new System.Drawing.Point(672, 488);
			this.btnDeposito.Name = "btnDeposito";
			this.btnDeposito.Size = new System.Drawing.Size(88, 24);
			this.btnDeposito.TabIndex = 746;
			this.btnDeposito.Text = "Depositar";
			this.btnDeposito.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDeposito.Click += new System.EventHandler(this.btnDeposito_Click);
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
			// groupBox5
			// 
			this.groupBox5.Location = new System.Drawing.Point(0, 64);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(912, 8);
			this.groupBox5.TabIndex = 747;
			this.groupBox5.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(912, 488);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(88, 24);
			this.btnGuardar.TabIndex = 749;
			this.btnGuardar.Text = "Cerrar Caja";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(768, 488);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(136, 24);
			this.btnActualizar.TabIndex = 748;
			this.btnActualizar.Text = "Actualizar Numeración";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// lblEstadoDepositos
			// 
			this.lblEstadoDepositos.AutoSize = true;
			this.lblEstadoDepositos.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.lblEstadoDepositos.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold);
			this.lblEstadoDepositos.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstadoDepositos.Location = new System.Drawing.Point(648, 41);
			this.lblEstadoDepositos.Name = "lblEstadoDepositos";
			this.lblEstadoDepositos.Size = new System.Drawing.Size(0, 21);
			this.lblEstadoDepositos.TabIndex = 751;
			this.lblEstadoDepositos.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(648, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 21);
			this.lblEstado.TabIndex = 750;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox6
			// 
			this.groupBox6.Location = new System.Drawing.Point(664, 472);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(528, 8);
			this.groupBox6.TabIndex = 752;
			this.groupBox6.TabStop = false;
			this.groupBox6.Enter += new System.EventHandler(this.groupBox6_Enter);
			// 
			// frmResumenCierreCaja
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1194, 518);
			this.Controls.Add(this.groupBox6);
			this.Controls.Add(this.lblEstadoDepositos);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.uGridDepositos);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.btnDeposito);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.optTurno);
			this.Controls.Add(this.cmbCajero);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridCaja);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmResumenCierreCaja";
			this.Text = "Resumen Diario De Caja";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmResumenCierreCaja_KeyDown);
			this.Load += new System.EventHandler(this.frmResumenCierreCaja_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTurno)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuadre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMonedas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB50)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB50)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTB100)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCantB100)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGasto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotVEnta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idCaja = 0;
		int idDesCaja = 0;
		int iEstado = 0;
		int iEstadoCaja = 0;
		decimal dValorDeposito = 0;
		private Acceso miAcceso;

		private void Controles(bool bEstado)
		{
			this.txtCantB100.Enabled = bEstado;
			this.txtCantB50.Enabled = bEstado;
			this.txtCantB20.Enabled = bEstado;
			this.txtCantB10.Enabled = bEstado;
			this.txtCantB5.Enabled = bEstado;
			this.txtCantB1.Enabled = bEstado;
			this.txtMonedas.Enabled = bEstado;
			this.txtGasto.Enabled = bEstado;
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmResumenCierreCaja_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1110");

			if (!Funcion.Permiso("833", cdsSeteoF))
			{					
				MessageBox.Show("No tiene acceso a cuadre de caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
						
			if (miAcceso.B15DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-15);
			else if (miAcceso.B30DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-30);
			else if (miAcceso.B60DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-60);
			
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;			
			
			this.dtFecha.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("08/12/2017");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 45", MenuLatinium.IdUsuario));

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);	

			if (iBodega > 0) this.cmbBodega.Value = iBodega;
			else 
			{
				this.cmbBodega.Value = 1;
				this.cmbBodega.Enabled = true;
			}

			FuncionesProcedimientos.EstadoGrids(false, this.uGridDepositos);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridCaja);
			Controles(false);

			this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CuantaCorrienteLotte");
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2) Order By idFormaPago");
		}

		private void Consulta()
		{
			#region Inicia Valores
			idCaja = 0;
			iEstado = 0;
			iEstadoCaja = 1;

			this.lblEstado.Text = "";
			this.lblEstadoDepositos.Text = "";
			this.txtObservaciones.Text = "";
			this.txtObservaciones.Enabled = false;

			this.btnDeposito.Enabled = false;
			this.btnActualizar.Enabled = false;
			this.btnGuardar.Enabled = false;
			
			this.uGridCaja.DisplayLayout.Bands[0].Columns["Numero"].CellActivation = Activation.ActivateOnly;
			this.uGridCaja.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.ActivateOnly;
			
			this.uGridCaja.DisplayLayout.Bands[0].Columns["Impreso"].CellActivation = Activation.Disabled;
			this.uGridCaja.DisplayLayout.Bands[0].Columns["btnCaja"].CellActivation = Activation.Disabled;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridCaja);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridDepositos);
			Controles(false);

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", false, "", cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Caja")) return;
		
			int idCajero = 0;
			if (this.cmbCajero.ActiveRow != null) idCajero = (int)this.cmbCajero.Value;
			#endregion Inicia Valores

			#region Documentos de Caja
			string sSQL = string.Format("Exec ResumenCierreCaja '{0}', {1}, {2}, {3}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, idCajero, (int)this.optTurno.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridCaja);
			
			this.lblContador.Text = this.uGridCaja.Rows.Count + " Movimientos";
			#endregion Documentos de Caja

			#region Encabezado
			if ((int)this.optTurno.Value != 0)
			{
				string sSQLMaestro = string.Format("Exec ConsultaCaja '{0}', {1}, {2}", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, (int)this.optTurno.Value);
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
			}
			#endregion Encabezado
			
			#region Depositos de Caja
			sSQL = string.Format("Exec DepositosDeCaja '{0}', {1}, {2}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodega.Value, (int)this.optTurno.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDepositos);
			#endregion Depositos de Caja

			#region Resumen de caja
			string sSQLDMC = string.Format("Exec ResumenCierreDeCaja '{0}', {1}, {2}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),(int)this.cmbBodega.Value, (int)this.optTurno.Value);
			SqlDataReader drMovCaja = Funcion.rEscalarSQL(cdsSeteoF, sSQLDMC, true);
			drMovCaja.Read();

			if (drMovCaja.HasRows)
			{
				this.txtCantB100.Value = drMovCaja.GetInt32(0);
				this.txtCantB50.Value = drMovCaja.GetInt32(1);
				this.txtCantB20.Value = drMovCaja.GetInt32(2);
				this.txtCantB10.Value = drMovCaja.GetInt32(3);
				this.txtCantB5.Value = drMovCaja.GetInt32(4);
				this.txtCantB1.Value = drMovCaja.GetInt32(5);
				this.txtMonedas.Value = drMovCaja.GetDecimal(6);
				this.txtGasto.Value = drMovCaja.GetDecimal(7);				
			}
			drMovCaja.Close();
			#endregion Resumen de caja

			Total();

			int idProyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value), true);		

			if ((int)this.optTurno.Value != 0)
			{
				string sSQLActualiza = string.Format("Exec GuardaCajaLotte {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}", 
					idCaja, (int)this.cmbBodega.Value, idProyecto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					"", Convert.ToDecimal(this.txtEfectivo.Value), 0, dValorDeposito, 1, false, 
					this.txtObservaciones.Text.ToString(), (int)this.optTurno.Value, iEstadoCaja);			
				idCaja = Funcion.iEscalarSQL(cdsSeteoF, sSQLActualiza);	

				this.lblEstadoDepositos.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeCaja {0}", idCaja));
			}

			#region Controles
			if ((int)this.optTurno.Value > 0)
			{
				if (iEstadoCaja == 1) 
				{
					this.lblEstado.Text = "CAJA ABIERTA";
					this.btnGuardar.Text = "Cerrar Caja";
					this.btnDeposito.Enabled = true;
					this.txtObservaciones.Enabled = true;

					Controles(true);

					if (miAcceso.BActNum) 
					{
						this.btnActualizar.Enabled = true;
						this.uGridCaja.DisplayLayout.Bands[0].Columns["Numero"].CellActivation = Activation.AllowEdit;
						//this.uGridCaja.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.AllowEdit;
					}

					if (miAcceso.BModificarImpreso) this.uGridCaja.DisplayLayout.Bands[0].Columns["Impreso"].CellActivation = Activation.AllowEdit;
					if (miAcceso.BModifcarFormaDePago) this.uGridCaja.DisplayLayout.Bands[0].Columns["btnCaja"].CellActivation = Activation.AllowEdit;

					if (miAcceso.BCerrarCaja) this.btnGuardar.Enabled = true;
				
					if (this.uGridCaja.Rows.Count > 0) 
					{
						FuncionesProcedimientos.EstadoGrids(true, this.uGridCaja);
						FuncionesProcedimientos.EstadoGrids(true, this.uGridDepositos);
					}
				}
				if (iEstadoCaja == 2) 
				{
					this.lblEstado.Text = "CAJA CERRADA";
					this.btnGuardar.Text = "Abrir Caja";
					if (miAcceso.BAbrirCaja) this.btnGuardar.Enabled = true;
				}
			}
			if (this.uGridCaja.Rows.Count > 0) this.btnImprimir.Enabled = true;			
			#endregion Controles
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{	
			Consulta();
		}

		private void Total()
		{					
			#region Inicia Valores
			this.txtEfectivo.Value = 0.00m;
			this.txtVoucher.Value = 0.00m;
			this.txtVTotal.Value = 0.00m;
			this.txtTotVEnta.Value = 0.00m;
			this.txtDescuadre.Value = 0.00m;
			
			this.txtVTB100.Value = 0.00m;
			this.txtVTB50.Value = 0.00m;
			this.txtVTB20.Value = 0.00m;
			this.txtVTB10.Value = 0.00m;
			this.txtVTB5.Value = 0.00m;
			this.txtVTB1.Value = 0.00m;

			this.txtDeposito.Value = 0.00m;
			#endregion Inicia Valores

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridCaja.Rows.All)
			{
				if (dr.Cells["Efectivo"].Value != DBNull.Value)
					this.txtEfectivo.Value = (decimal)this.txtEfectivo.Value + Math.Round(Convert.ToDecimal(dr.Cells["Efectivo"].Value), 2);
				if (dr.Cells["TC"].Value != DBNull.Value)
					this.txtVoucher.Value = (decimal)this.txtVoucher.Value + Math.Round(Convert.ToDecimal(dr.Cells["TC"].Value), 2) + Math.Round(Convert.ToDecimal(dr.Cells["TD"].Value), 2);
				if (dr.Cells["Total"].Value != DBNull.Value)
					this.txtTotVEnta.Value = (decimal)this.txtTotVEnta.Value + Math.Round(Convert.ToDecimal(dr.Cells["Total"].Value), 2);				
			}
			
			#region Total Billetes
			this.txtVTB100.Value = Convert.ToDecimal(this.txtCantB100.Value) * 100;
			this.txtVTB50.Value = Convert.ToDecimal(this.txtCantB50.Value) * 50;
			this.txtVTB20.Value = Convert.ToDecimal(this.txtCantB20.Value) * 20;
			this.txtVTB10.Value = Convert.ToDecimal(this.txtCantB10.Value) * 10;
			this.txtVTB5.Value = Convert.ToDecimal(this.txtCantB5.Value) * 5;
			this.txtVTB1.Value = Convert.ToDecimal(this.txtCantB1.Value) * 1;
			#endregion Total Billetes

			this.txtVTotal.Value = Convert.ToDecimal(this.txtVTB100.Value) +
				Convert.ToDecimal(this.txtVTB50.Value) +
				Convert.ToDecimal(this.txtVTB20.Value) +
				Convert.ToDecimal(this.txtVTB10.Value) +
				Convert.ToDecimal(this.txtVTB5.Value) +
				Convert.ToDecimal(this.txtVTB1.Value) +
				Convert.ToDecimal(this.txtMonedas.Value);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
				this.txtDeposito.Value = Convert.ToDecimal(this.txtDeposito.Value) + Convert.ToDecimal(dr.Cells["Valor"].Value);

			this.txtDescuadre.Value = Convert.ToDecimal(this.txtEfectivo.Value) - Convert.ToDecimal(this.txtDeposito.Value) - Convert.ToDecimal(this.txtGasto.Value);
			#endregion Desplazamiento de grilla			
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridCaja);
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaPersonalFacturas(this.cmbCajero, (DateTime)this.dtFecha.Value, 4, cdsSeteoF);	

				if (this.cmbBodega.ActiveRow != null)
					Consulta();
			}
		}

		private void txtCantB100_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCantB50.Focus();
		}

		private void txtCantB50_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCantB20.Focus();
		}

		private void txtCantB20_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCantB10.Focus();
		}

		private void txtCantB10_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCantB5.Focus();
		}

		private void txtCantB5_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCantB1.Focus();
		}

		private void txtCantB1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtMonedas.Focus();
		}

		private void txtMonedas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtGasto.Focus();
		}

		private void txtGasto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnDeposito.Focus();
		}

		private void txtCantB100_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCantB50_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCantB20_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCantB10_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCantB5_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCantB1_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtMonedas_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtGasto_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void txtCheque_Validated(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void GeneraAsiento()
		{		
			#region Crea Asiento
			int idDetCaja = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{
				if(dr.Cells["idDetalleCaja"].Value == System.DBNull.Value) idDetCaja = 0;
				else idDetCaja = (int)dr.Cells["idDetalleCaja"].Value;

				string sSQL = string.Format("Exec CreaAsiCierreCaja {0}, {1}", idDetCaja, idCaja);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
			}			

			string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCaja {0}", idCaja);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLDiferencias, true);
			#endregion Crea Asiento
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void uGridDepositos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if (DialogResult.Yes ==	MessageBox.Show("Desea Eliminar Registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From AsientoDetalle Where idAsiento = {0} Delete From Asiento Where idAsiento = {0}", (int)e.Rows[i].Cells["idAsiento"].Value));
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetalleCaja Where idDetalleCaja = {0}", (int)e.Rows[i].Cells["idDetalleCaja"].Value));
					e.DisplayPromptMsg = false;
				}
				else e.Cancel = true;
			}
		}

		private void uGridDepositos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDepositos);
		}

		private void txtCantB100_Click(object sender, System.EventArgs e)
		{
			this.txtCantB100.SelectAll();
		}

		private void txtCantB50_Click(object sender, System.EventArgs e)
		{
			this.txtCantB50.SelectAll();
		}

		private void txtCantB20_Click(object sender, System.EventArgs e)
		{
			this.txtCantB20.SelectAll();
		}

		private void txtCantB10_Click(object sender, System.EventArgs e)
		{
			this.txtCantB10.SelectAll();
		}

		private void txtCantB5_Click(object sender, System.EventArgs e)
		{
			this.txtCantB5.SelectAll();
		}

		private void txtCantB1_Click(object sender, System.EventArgs e)
		{
			this.txtCantB1.SelectAll();
		}

		private void txtMonedas_Click(object sender, System.EventArgs e)
		{
			this.txtMonedas.SelectAll();
		}

		private void txtGasto_Click(object sender, System.EventArgs e)
		{
			this.txtGasto.SelectAll();
		}
		
		private void txtCantB100_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB100.SelectAll();
		}

		private void txtCantB50_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB50.SelectAll();
		}

		private void txtCantB20_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB20.SelectAll();
		}

		private void txtCantB10_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB10.SelectAll();
		}

		private void txtCantB5_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB5.SelectAll();
		}

		private void txtCantB1_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB1.SelectAll();
		}

		private void txtGasto_Enter(object sender, System.EventArgs e)
		{
			this.txtGasto.SelectAll();
		}
		
		private void txtMonedas_Enter(object sender, System.EventArgs e)
		{
			this.txtMonedas.SelectAll();
		}

		private void frmResumenCierreCaja_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridDepositos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void uGridDepositos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO") Funcion.CamposNumericos(sender, e);
			if (this.uGridDepositos.ActiveCell.Column.Key.ToString().ToUpper() == "NUMCUENTA") Funcion.CamposNumericos(sender, e);
		}

		private void uGridDepositos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idDetalleCaja"].Value = 0;
			e.Row.Cells["Valor"].Value = 0.00m;
		}

		private void btnDeposito_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", true, "Bancos", cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una caja")) return;

			if ((int)this.optTurno.Value == 0)
			{
				MessageBox.Show("Debe seleccionar TURNO 1 o 2 para realizar el Depósito?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (this.uGridCaja.Rows.Count == 0)
			{
				MessageBox.Show("No Existen datos de caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (this.uGridDepositos.Rows.Count == 0)
			{
				MessageBox.Show("No Existen documentos para realizar el deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}			
			#endregion Validacion
	
			#region Proyecto
			string sSQLProyecto = string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value);
			int idProyecto = Funcion.iEscalarSQL(cdsSeteoF, sSQLProyecto, true);
			#endregion Proyecto

			DateTime FechaCaja = (DateTime)this.dtFecha.Value;
						
			string sSQLMaestro = string.Format("Exec GuardaCajaLotte {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}", 
				idCaja, (int)this.cmbBodega.Value, idProyecto, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
				"", Convert.ToDecimal(this.txtEfectivo.Value), 0, dValorDeposito, 1, false, 
				this.txtObservaciones.Text.ToString(), (int)this.optTurno.Value, iEstadoCaja);			
			idCaja = Funcion.iEscalarSQL(cdsSeteoF, sSQLMaestro);			
			
			#region Detalle
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{				
				string sSQLDetalle = string.Format("Exec GuardaDetalleCaja {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
					(int)dr.Cells["idDetalleCaja"].Value, idCaja, (int)dr.Cells["idCuenta"].Value, (int)dr.Cells["idFormaPago"].Value, dr.Cells["Numero"].Value.ToString(),  
					Convert.ToDateTime(dr.Cells["FechaDeposito"].Value).ToString("yyyyMMdd"), Convert.ToDecimal(dr.Cells["Valor"].Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle, true);
			}

			this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCaja {0}", idCaja));
			this.lblEstadoDepositos.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec MensajeCaja {0}", idCaja));			
			#endregion Detalle
			
			#region registro movimiento de caja
			DateTime FeDesc = DateTime.Today;

			string sSQLDescdr = string.Format("Exec RegDescuadre {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}", 
				idCaja, 0, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtObservaciones.Text.ToString(),
				Convert.ToDecimal(this.txtDescuadre.Value),
				Convert.ToInt32(this.txtCantB100.Value), Convert.ToInt32(this.txtCantB50.Value), 
				Convert.ToInt32(this.txtCantB20.Value), Convert.ToInt32(this.txtCantB10.Value),
				Convert.ToInt32(this.txtCantB5.Value), Convert.ToInt32(this.txtCantB1.Value),
				Convert.ToDecimal(this.txtMonedas.Value), Convert.ToDecimal(this.txtGasto.Value),
				0.00m, idDesCaja, 0.00m,
				//Convert.ToDecimal(this.txtCheque.Value), idDesCaja, Convert.ToDecimal(this.txtDepositoReserva.Value),
				0.00m,
				//Convert.ToDecimal(this.txtValorPrenda.Value), 
				Convert.ToDecimal(this.txtDeposito.Value), Convert.ToDecimal(this.txtEfectivo.Value));
			Funcion.EjecutaSQL(cdsSeteoF, sSQLDescdr, true);
			#endregion registro movimiento de caja

			GeneraAsiento();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtFecha.Value;
						
			int idCajera = 0;
			if (this.cmbCajero.ActiveRow != null) idCajera = 0;

			#region Impresion
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Fecha";
			discreteValDesde.Value = dtFecha;
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);

			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idBodg = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@Bodega";
			idBodg.Value = (int)this.cmbBodega.Value;
			pfRet.CurrentValues.Add (idBodg);
			paramFields.Add (pfRet);
				
			ParameterField idCjero = new ParameterField ();
			ParameterDiscreteValue idCjr = new ParameterDiscreteValue ();
			idCjero.ParameterFieldName = "@idCajero";
			idCjr.Value = idCajera;
			idCjero.CurrentValues.Add (idCjr);
			paramFields.Add (idCjero);

			ParameterField nTurn = new ParameterField ();
			ParameterDiscreteValue nTurnCaj = new ParameterDiscreteValue ();
			nTurn.ParameterFieldName = "@Turno";
			nTurnCaj.Value = (int)this.optTurno.Value;
			nTurn.CurrentValues.Add (nTurnCaj);
			paramFields.Add (nTurn);
				
			Reporte miRepor = new Reporte("ResumenDiarioDeCajas.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Retenciones");
			miRepor.crVista.ParameterFieldInfo = paramFields;
			#endregion Impresion
				
			miRepor.Show();
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			string sMensaje = "";
			string sNumero = "";

			if (this.uGridCaja.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros de caja para actualizar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
				return;
			}
															
			#region Valida Numeracion Existente
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCaja.Rows.All)
			{
				if (dr.Cells["Origen"].Value.ToString() == "FAC")
				{
					char Pad = '0';
					sNumero = dr.Cells["Factura"].Value.ToString().PadLeft(9, Pad);
					
					string sSQL = string.Format("Exec ValidaNumeracionExistente {0}, '{1}', {2}", (int)dr.Cells["idCompra"].Value, sNumero, (int)this.cmbBodega.Value);
					
					if (Funcion.sEscalarSQL(cdsSeteoF, sSQL, true).ToString().Length > 0)
					{
						MessageBox.Show(Funcion.sEscalarSQL(cdsSeteoF, sSQL, true).ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
						this.uGridCaja.ActiveRow = dr;
						this.uGridCaja.ActiveRow.Selected = true;						
						return;
					}
				}
				else
				{
					char Pad = '0';
					sNumero = dr.Cells["Factura"].Value.ToString().PadLeft(7, Pad);
					
					string sSQL = string.Format("Select COUNT(*) From Anticipos Where Secuencial = '{1}' And Bodega = {2} And idAnticipo <> {0}", (int)dr.Cells["idCompra"].Value, sNumero, (int)this.cmbBodega.Value);
					
					if (Funcion.iEscalarSQL(cdsSeteoF, sSQL, true) > 0)
					{
						MessageBox.Show(string.Format("El Numero de Recibo '{0}' ya existe", sNumero), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);		
						this.uGridCaja.ActiveRow = dr;
						this.uGridCaja.ActiveRow.Selected = true;						
						return;
					}
				}
			}
			#endregion Valida Numeracion Existente
				
			#region Actualizacion
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCaja.Rows.All)
			{
				int idTipoFactura = 0;

				if (dr.Cells["Origen"].Value.ToString() == "FAC") idTipoFactura = 1;
				else if (dr.Cells["Origen"].Value.ToString() == "CP" || dr.Cells["Origen"].Value.ToString() == "ANT" || dr.Cells["Origen"].Value.ToString() == "AB") idTipoFactura = 24;

				string sSQL = string.Format("Exec ActualizaNumeracionCuadrecaja {0}, {1}, '{2}', '{3}'", 
					(int)dr.Cells["idCompra"].Value, idTipoFactura, dr.Cells["Factura"].Value.ToString(), Convert.ToDateTime(dr.Cells["Fecha"].Value).ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
			#endregion Actualizacion

			this.btnBuscar_Click(sender, e);
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			int iEst = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select EstadoCaja From Caja Where idCaja = {0}", idCaja));

			#region Validaciones
			if (iEst != iEstadoCaja)
			{
				MessageBox.Show("El registro de esta caja ha sido modificada por otro usuario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.btnBuscar_Click(sender, e);
				return;
			}

			if (iEstadoCaja == 1)
			{
				if ((decimal)this.txtEfectivo.Value > 0.00m)
				{
					int iCont = 0;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
						if ((int)dr.Cells["idFormaPago"].Value == 1) iCont++;

					if (iCont == 0)
					{						
						MessageBox.Show("No se ha registrado el deposito en efectivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
						return;
					}
				}		

				if (iEstado == 1)
				{
					MessageBox.Show("No se ha registrado el deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}

				if (iEstado == 4)
				{
					MessageBox.Show(this.lblEstadoDepositos.Text.ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}
			}
			#endregion Validaciones

			string sMensaje = "Cerrar";
			if (iEstadoCaja == 2) sMensaje = "Abrir";

			if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Esta seguro de '{0}' esta caja?", sMensaje), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				int iNuevoEstado = 1;
				if (iEstadoCaja == 1) iNuevoEstado = 2;

				string sSQL = string.Format("Exec RevisadoCaja {0}, '{1}', {2}, {3}", true, this.txtObservaciones.Text.ToString(), idCaja, iNuevoEstado);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				if (iEstadoCaja == 1) GeneraAsiento();

				MessageBox.Show("Proceso terminado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnBuscar_Click(sender, e);		
			}
		}

		private void groupBox6_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void uGridCaja_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Impreso")
				{
					this.uGridCaja.UpdateData();
					
					string sSQL = string.Format("Exec ActualizaImpresoFacturasAnticipos {0}, {1}, '{2}'", 
						(bool)e.Cell.Row.Cells["Impreso"].Value, (int)e.Cell.Row.Cells["idCompra"].Value, 1);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				}
			}
		}

		private void uGridCaja_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
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
				
				#region Formas de pago
				using (VentaFPagoLot Cobro = new VentaFPagoLot((int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value, Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value), 
								 e.Cell.Row.Cells["Numero"].Value.ToString(), false, 0, 0, false, 0, false, 0, false, 0, 0, 0, (int)e.Cell.Row.Cells["idCliente"].Value, 0, 0, 0, 0))
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
									(int)e.Cell.Row.Cells["idCompra"].Value, 1, 
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

								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CreaAsientoDeVenta {0}", (int)e.Cell.Row.Cells["idCompra"].Value));
							}
							#endregion Guarda							
						}								
						#endregion Guarda Cobro de la factura

						this.btnBuscar_Click(sender, e);
					}					
				}
				#endregion Formas de pago
			}
		}

		private void uGridCaja_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridCaja_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridCaja);
		}

		private void uGridDepositos_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
		
		}

		private void optTurno_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
				Consulta();
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}
