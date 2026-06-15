using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRepoCajaLotte.
	/// </summary>
	public class frmRepoCajaLotte : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label6;
		public DevExpress.XtraEditors.SimpleButton btVer;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public DevExpress.XtraEditors.SimpleButton btExportar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteo;
		private System.Windows.Forms.ImageList imageList1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFeFinaliza;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCajero;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		public DevExpress.XtraEditors.SimpleButton btnGuardar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdOperador;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorriente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugCaja;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.RadioButton rb1;
		private System.Windows.Forms.RadioButton rb2;
		private System.Windows.Forms.RadioButton rb3;
		private System.Windows.Forms.Label lblEdoDpto;
		public DevExpress.XtraEditors.SimpleButton simpleButton1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVoucher;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGasto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB100;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB100;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB50;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB20;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB10;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTB1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCantB1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMonedas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuadre;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotVEnta;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCheque;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDepositoReserva;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorPrenda;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private System.ComponentModel.IContainer components;

		public frmRepoCajaLotte()
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRepoCajaLotte));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Turno");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroHabitacion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HABITACION");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("COSUMO");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EFE");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TC");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TD");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DEPOSITO");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PRENDA");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CompraYA");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descuento");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NomLocal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Operadora");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "HABITACION", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "HABITACION", 5, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "COSUMO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "COSUMO", 6, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "EFE", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "EFE", 7, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TC", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TC", 8, true);
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TD", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TD", 9, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PRENDA", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PRENDA", 11, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 14, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Count, null, "NumeroHabitacion", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NumeroHabitacion", 4, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DEPOSITO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DEPOSITO", 10, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CompraYA", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CompraYA", 12, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Descuento", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Descuento", 13, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("HABITACION");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("COSUMO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EFE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DEPOSITO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PRENDA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CompraYA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NomLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Operadora");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuentaCorriente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanB100");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanB50");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanB20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanB10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanB5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanB1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CanMonedas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorGasto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDescuadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOperadorDescuadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetMovCaja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValDepReserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValPrenda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TurnoCaja");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanB100");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanB50");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanB20");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanB10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanB5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanB1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CanMonedas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorGasto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDescuadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOperadorDescuadre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetMovCaja");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValDepReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValPrenda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TurnoCaja");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroHabitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Turno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Habitacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Consumo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TD");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Prenda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CompraYA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			this.lblBodega = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btVer = new DevExpress.XtraEditors.SimpleButton();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btExportar = new DevExpress.XtraEditors.SimpleButton();
			this.ugCaja = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFeFinaliza = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCajero = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.cmdOperador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnGuardar = new DevExpress.XtraEditors.SimpleButton();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCuentaCorriente = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label7 = new System.Windows.Forms.Label();
			this.rb1 = new System.Windows.Forms.RadioButton();
			this.rb2 = new System.Windows.Forms.RadioButton();
			this.rb3 = new System.Windows.Forms.RadioButton();
			this.lblEdoDpto = new System.Windows.Forms.Label();
			this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtValorPrenda = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.txtDepositoReserva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.txtCheque = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtTotVEnta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label22 = new System.Windows.Forms.Label();
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
			this.txtVoucher = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFinaliza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdOperador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtValorPrenda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepositoReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotVEnta)).BeginInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtVoucher)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 8);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(41, 16);
			this.lblBodega.TabIndex = 249;
			this.lblBodega.Text = "LOCAL";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(288, 16);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(40, 24);
			this.label6.TabIndex = 248;
			this.label6.Text = "Fecha BUSCA";
			// 
			// btVer
			// 
			this.btVer.Appearance.Options.UseTextOptions = true;
			this.btVer.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btVer.Image = ((System.Drawing.Image)(resources.GetObject("btVer.Image")));
			this.btVer.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btVer.ImageIndex = 4;
			this.btVer.Location = new System.Drawing.Point(592, 12);
			this.btVer.Name = "btVer";
			this.btVer.Size = new System.Drawing.Size(48, 32);
			this.btVer.TabIndex = 250;
			this.btVer.Click += new System.EventHandler(this.btVer_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 30;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 180;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(8, 24);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(112, 21);
			this.cmbBodega.TabIndex = 244;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(328, 18);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 245;
			this.dtFecha.Value = new System.DateTime(2013, 7, 1, 0, 0, 0, 0);
			// 
			// btExportar
			// 
			this.btExportar.Appearance.Options.UseTextOptions = true;
			this.btExportar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btExportar.Image = ((System.Drawing.Image)(resources.GetObject("btExportar.Image")));
			this.btExportar.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.btExportar.ImageIndex = 2;
			this.btExportar.Location = new System.Drawing.Point(640, 12);
			this.btExportar.Name = "btExportar";
			this.btExportar.Size = new System.Drawing.Size(48, 32);
			this.btExportar.TabIndex = 247;
			this.btExportar.Click += new System.EventHandler(this.btExportar_Click);
			// 
			// ugCaja
			// 
			this.ugCaja.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.ugCaja.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugCaja.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ugCaja.DisplayLayout.Appearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance3;
			ultraGridColumn7.Header.Caption = "Trno";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 30;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Header.Caption = "#";
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 30;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Width = 56;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Width = 56;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Width = 56;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 56;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 56;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance10;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 56;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance11;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 56;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance12;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "CpaYA";
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Width = 56;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance13;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Dscto";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Width = 56;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance14;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Width = 65;
			ultraGridColumn20.Header.VisiblePosition = 15;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 87;
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 17;
			ultraGridColumn22.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn22});
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance15;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance16;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance17;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance18;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance19;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance20;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance21;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings8.Appearance = appearance22;
			summarySettings8.DisplayFormat = "{0: #,##0}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance23;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance24;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance25;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings11});
			ultraGridBand2.SummaryFooterCaption = "Total CAJA";
			this.ugCaja.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ugCaja.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugCaja.DisplayLayout.Override.ActiveRowAppearance = appearance26;
			this.ugCaja.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugCaja.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugCaja.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.ugCaja.DisplayLayout.Override.CardAreaAppearance = appearance27;
			this.ugCaja.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.ugCaja.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.None;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugCaja.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.ugCaja.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugCaja.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugCaja.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.ugCaja.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugCaja.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ugCaja.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ugCaja.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugCaja.Location = new System.Drawing.Point(8, 48);
			this.ugCaja.Name = "ugCaja";
			this.ugCaja.Size = new System.Drawing.Size(896, 312);
			this.ugCaja.TabIndex = 251;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn22});
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
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(592, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 24);
			this.label1.TabIndex = 253;
			this.label1.Text = "Fecha FINAL";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.label1.Visible = false;
			// 
			// dtFeFinaliza
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFeFinaliza.Appearance = appearance31;
			this.dtFeFinaliza.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFeFinaliza.DateButtons.Add(dateButton2);
			this.dtFeFinaliza.Format = "dd/MM/yyyy";
			this.dtFeFinaliza.Location = new System.Drawing.Point(632, 16);
			this.dtFeFinaliza.Name = "dtFeFinaliza";
			this.dtFeFinaliza.NonAutoSizeHeight = 23;
			this.dtFeFinaliza.Size = new System.Drawing.Size(104, 21);
			this.dtFeFinaliza.SpinButtonsVisible = true;
			this.dtFeFinaliza.TabIndex = 252;
			this.dtFeFinaliza.Value = new System.DateTime(2013, 7, 1, 0, 0, 0, 0);
			this.dtFeFinaliza.Visible = false;
			// 
			// cmbCajero
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCajero.Appearance = appearance32;
			this.cmbCajero.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCajero.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 335;
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.cmbCajero.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCajero.DisplayMember = "Nombre";
			this.cmbCajero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCajero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCajero.Location = new System.Drawing.Point(128, 24);
			this.cmbCajero.Name = "cmbCajero";
			this.cmbCajero.Size = new System.Drawing.Size(152, 21);
			this.cmbCajero.TabIndex = 255;
			this.cmbCajero.ValueMember = "idPersonal";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(128, 8);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(38, 16);
			this.label19.TabIndex = 254;
			this.label19.Text = "Cajero";
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource3;
			appearance33.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance33;
			this.uGridDepositos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 78;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 54;
			ultraGridColumn28.Header.Caption = "Forma de pago";
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn28.Width = 167;
			ultraGridColumn29.Header.Caption = "Banco";
			ultraGridColumn29.Header.VisiblePosition = 3;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn29.Width = 171;
			ultraGridColumn30.Header.Caption = "Número";
			ultraGridColumn30.Header.VisiblePosition = 4;
			ultraGridColumn30.Width = 160;
			ultraGridColumn31.Header.Caption = "Fecha";
			ultraGridColumn31.Header.VisiblePosition = 5;
			ultraGridColumn31.Width = 71;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance34;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 6;
			ultraGridColumn32.PromptChar = ' ';
			ultraGridColumn32.Width = 89;
			ultraGridColumn33.Header.VisiblePosition = 7;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 56;
			ultraGridColumn34.Header.VisiblePosition = 8;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 25;
			ultraGridColumn35.Header.VisiblePosition = 10;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 23;
			ultraGridColumn36.Header.VisiblePosition = 11;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 23;
			ultraGridColumn37.Header.VisiblePosition = 12;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 23;
			ultraGridColumn38.Header.VisiblePosition = 13;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 23;
			ultraGridColumn39.Header.VisiblePosition = 14;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 23;
			ultraGridColumn40.Header.VisiblePosition = 15;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 45;
			ultraGridColumn41.Header.VisiblePosition = 16;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 46;
			ultraGridColumn42.Header.VisiblePosition = 17;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 48;
			ultraGridColumn43.Header.VisiblePosition = 18;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 59;
			ultraGridColumn44.Header.VisiblePosition = 19;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 78;
			ultraGridColumn45.Header.VisiblePosition = 20;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 84;
			ultraGridColumn46.Header.VisiblePosition = 9;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 74;
			ultraGridColumn47.Header.VisiblePosition = 21;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 77;
			ultraGridColumn48.Header.VisiblePosition = 22;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 62;
			ultraGridColumn49.Header.VisiblePosition = 23;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 61;
			ultraGridBand4.Columns.AddRange(new object[] {
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
			ultraGridBand4.Header.Caption = "Depositos";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance35.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance35;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance36.FontData.BoldAsString = "True";
			appearance36.FontData.Name = "Arial";
			appearance36.FontData.SizeInPoints = 8F;
			appearance36.ForeColor = System.Drawing.Color.White;
			appearance36.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance38;
			this.uGridDepositos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(8, 360);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(696, 104);
			this.uGridDepositos.TabIndex = 257;
			this.uGridDepositos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDepositos_BeforeRowsDeleted);
			this.uGridDepositos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDepositos_KeyDown);
			this.uGridDepositos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDepositos_AfterCellUpdate);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(System.DateTime);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(System.Decimal);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn46.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn46});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 472);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(77, 16);
			this.label2.TabIndex = 258;
			this.label2.Text = "DESCUADRE";
			this.label2.TextAlign = System.Drawing.ContentAlignment.TopRight;
			// 
			// cmdOperador
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmdOperador.Appearance = appearance39;
			this.cmdOperador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdOperador.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 1;
			ultraGridColumn51.Width = 250;
			ultraGridColumn52.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52});
			this.cmdOperador.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmdOperador.DisplayMember = "Nombre";
			this.cmdOperador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdOperador.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdOperador.Location = new System.Drawing.Point(48, 496);
			this.cmdOperador.Name = "cmdOperador";
			this.cmdOperador.Size = new System.Drawing.Size(152, 21);
			this.cmdOperador.TabIndex = 260;
			this.cmdOperador.ValueMember = "idPersonal";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 498);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(43, 16);
			this.label3.TabIndex = 259;
			this.label3.Text = "Usuario";
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(248, 480);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 16);
			this.label4.TabIndex = 262;
			this.label4.Text = "Observaciones";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance40;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(248, 496);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(344, 40);
			this.txtObservaciones.TabIndex = 261;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.btnGuardar.Appearance.Options.UseFont = true;
			this.btnGuardar.Appearance.Options.UseTextOptions = true;
			this.btnGuardar.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.ImageIndex = 3;
			this.btnGuardar.ImageList = this.imageList1;
			this.btnGuardar.Location = new System.Drawing.Point(1000, 440);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 48);
			this.btnGuardar.TabIndex = 263;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtValor
			// 
			this.txtValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance41;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(600, 496);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(80, 21);
			this.txtValor.TabIndex = 265;
			this.txtValor.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(560, 480);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(31, 16);
			this.label5.TabIndex = 264;
			this.label5.Text = "Valor";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.label5.Visible = false;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource4;
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.Caption = "";
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Width = 200;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn53,
																										 ultraGridColumn54});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(112, 368);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(224, 86);
			this.cmbFormaPago.TabIndex = 267;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn47.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn47,
																																 ultraDataColumn48});
			// 
			// cmbCuentaCorriente
			// 
			this.cmbCuentaCorriente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorriente.DataSource = this.ultraDataSource5;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.Caption = "";
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Width = 250;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn55,
																										 ultraGridColumn56});
			this.cmbCuentaCorriente.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbCuentaCorriente.DisplayMember = "Descripcion";
			this.cmbCuentaCorriente.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorriente.Location = new System.Drawing.Point(344, 368);
			this.cmbCuentaCorriente.Name = "cmbCuentaCorriente";
			this.cmbCuentaCorriente.Size = new System.Drawing.Size(272, 88);
			this.cmbCuentaCorriente.TabIndex = 266;
			this.cmbCuentaCorriente.ValueMember = "idCuenta";
			this.cmbCuentaCorriente.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn49.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn49,
																																 ultraDataColumn50});
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(448, 8);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(46, 16);
			this.label7.TabIndex = 268;
			this.label7.Text = "TURNO";
			// 
			// rb1
			// 
			this.rb1.Checked = true;
			this.rb1.Location = new System.Drawing.Point(448, 24);
			this.rb1.Name = "rb1";
			this.rb1.Size = new System.Drawing.Size(32, 16);
			this.rb1.TabIndex = 269;
			this.rb1.TabStop = true;
			this.rb1.Text = "1";
			// 
			// rb2
			// 
			this.rb2.Location = new System.Drawing.Point(484, 24);
			this.rb2.Name = "rb2";
			this.rb2.Size = new System.Drawing.Size(32, 16);
			this.rb2.TabIndex = 270;
			this.rb2.Text = "2";
			// 
			// rb3
			// 
			this.rb3.Location = new System.Drawing.Point(520, 24);
			this.rb3.Name = "rb3";
			this.rb3.Size = new System.Drawing.Size(56, 16);
			this.rb3.TabIndex = 271;
			this.rb3.Text = "Todos";
			// 
			// lblEdoDpto
			// 
			this.lblEdoDpto.Location = new System.Drawing.Point(8, 520);
			this.lblEdoDpto.Name = "lblEdoDpto";
			this.lblEdoDpto.Size = new System.Drawing.Size(192, 23);
			this.lblEdoDpto.TabIndex = 272;
			this.lblEdoDpto.Text = "...";
			// 
			// simpleButton1
			// 
			this.simpleButton1.Appearance.Options.UseTextOptions = true;
			this.simpleButton1.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
			this.simpleButton1.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.Image")));
			this.simpleButton1.ImageAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.simpleButton1.ImageIndex = 2;
			this.simpleButton1.Location = new System.Drawing.Point(688, 12);
			this.simpleButton1.Name = "simpleButton1";
			this.simpleButton1.Size = new System.Drawing.Size(48, 32);
			this.simpleButton1.TabIndex = 273;
			this.simpleButton1.ToolTipTitle = "Imprimir INFORMACION";
			this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.groupBox1.Controls.Add(this.txtValorPrenda);
			this.groupBox1.Controls.Add(this.label26);
			this.groupBox1.Controls.Add(this.txtDepositoReserva);
			this.groupBox1.Controls.Add(this.label25);
			this.groupBox1.Controls.Add(this.label24);
			this.groupBox1.Controls.Add(this.txtCheque);
			this.groupBox1.Controls.Add(this.label23);
			this.groupBox1.Controls.Add(this.txtTotVEnta);
			this.groupBox1.Controls.Add(this.label22);
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
			this.groupBox1.Controls.Add(this.txtVoucher);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.txtEfectivo);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Location = new System.Drawing.Point(913, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(280, 384);
			this.groupBox1.TabIndex = 274;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Movimientos de CAJA";
			// 
			// txtValorPrenda
			// 
			this.txtValorPrenda.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtValorPrenda.Appearance = appearance42;
			this.txtValorPrenda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorPrenda.FormatString = "#,##0.00";
			this.txtValorPrenda.Location = new System.Drawing.Point(176, 330);
			this.txtValorPrenda.Name = "txtValorPrenda";
			this.txtValorPrenda.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorPrenda.PromptChar = ' ';
			this.txtValorPrenda.ReadOnly = true;
			this.txtValorPrenda.Size = new System.Drawing.Size(80, 21);
			this.txtValorPrenda.TabIndex = 307;
			this.txtValorPrenda.Click += new System.EventHandler(this.txtValorPrenda_Click);
			this.txtValorPrenda.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValorPrenda_KeyPress);
			this.txtValorPrenda.Validated += new System.EventHandler(this.txtValorPrenda_Validated);
			this.txtValorPrenda.Enter += new System.EventHandler(this.txtValorPrenda_Enter);
			// 
			// label26
			// 
			this.label26.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label26.AutoSize = true;
			this.label26.Location = new System.Drawing.Point(123, 330);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(41, 16);
			this.label26.TabIndex = 306;
			this.label26.Text = "Prenda";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtDepositoReserva
			// 
			this.txtDepositoReserva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtDepositoReserva.Appearance = appearance43;
			this.txtDepositoReserva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepositoReserva.FormatString = "#,##0.00";
			this.txtDepositoReserva.Location = new System.Drawing.Point(176, 308);
			this.txtDepositoReserva.Name = "txtDepositoReserva";
			this.txtDepositoReserva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDepositoReserva.PromptChar = ' ';
			this.txtDepositoReserva.ReadOnly = true;
			this.txtDepositoReserva.Size = new System.Drawing.Size(80, 21);
			this.txtDepositoReserva.TabIndex = 305;
			this.txtDepositoReserva.Click += new System.EventHandler(this.txtDepositoReserva_Click);
			this.txtDepositoReserva.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDepositoReserva_KeyPress);
			this.txtDepositoReserva.Validated += new System.EventHandler(this.txtDepositoReserva_Validated);
			this.txtDepositoReserva.Enter += new System.EventHandler(this.txtDepositoReserva_Enter);
			// 
			// label25
			// 
			this.label25.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label25.AutoSize = true;
			this.label25.Location = new System.Drawing.Point(58, 310);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(106, 16);
			this.label25.TabIndex = 304;
			this.label25.Text = "Deposito RESERVA";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label24
			// 
			this.label24.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label24.AutoSize = true;
			this.label24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.Location = new System.Drawing.Point(42, 16);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(43, 16);
			this.label24.TabIndex = 303;
			this.label24.Text = "Billetes";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtCheque
			// 
			this.txtCheque.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCheque.Appearance = appearance44;
			this.txtCheque.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCheque.FormatString = "#,##0.00";
			this.txtCheque.Location = new System.Drawing.Point(176, 286);
			this.txtCheque.Name = "txtCheque";
			this.txtCheque.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCheque.PromptChar = ' ';
			this.txtCheque.Size = new System.Drawing.Size(80, 21);
			this.txtCheque.TabIndex = 302;
			this.txtCheque.Click += new System.EventHandler(this.txtCheque_Click);
			this.txtCheque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCheque_KeyPress);
			this.txtCheque.Validated += new System.EventHandler(this.txtCheque_Validated);
			this.txtCheque.Enter += new System.EventHandler(this.txtCheque_Enter);
			// 
			// label23
			// 
			this.label23.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label23.AutoSize = true;
			this.label23.Location = new System.Drawing.Point(120, 288);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(44, 16);
			this.label23.TabIndex = 301;
			this.label23.Text = "Cheque";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtTotVEnta
			// 
			this.txtTotVEnta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance45.ForeColor = System.Drawing.Color.Brown;
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotVEnta.Appearance = appearance45;
			this.txtTotVEnta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotVEnta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotVEnta.FormatString = "#,##0.00";
			this.txtTotVEnta.Location = new System.Drawing.Point(64, 240);
			this.txtTotVEnta.Name = "txtTotVEnta";
			this.txtTotVEnta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotVEnta.PromptChar = ' ';
			this.txtTotVEnta.ReadOnly = true;
			this.txtTotVEnta.Size = new System.Drawing.Size(80, 21);
			this.txtTotVEnta.TabIndex = 300;
			// 
			// label22
			// 
			this.label22.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label22.AutoSize = true;
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.ForeColor = System.Drawing.Color.Brown;
			this.label22.Location = new System.Drawing.Point(16, 242);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(43, 16);
			this.label22.TabIndex = 299;
			this.label22.Text = "VENTA";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtDescuadre
			// 
			this.txtDescuadre.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuadre.Appearance = appearance46;
			this.txtDescuadre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuadre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescuadre.FormatString = "#,##0.00";
			this.txtDescuadre.Location = new System.Drawing.Point(176, 354);
			this.txtDescuadre.Name = "txtDescuadre";
			this.txtDescuadre.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuadre.PromptChar = ' ';
			this.txtDescuadre.ReadOnly = true;
			this.txtDescuadre.Size = new System.Drawing.Size(80, 21);
			this.txtDescuadre.TabIndex = 298;
			// 
			// label21
			// 
			this.label21.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label21.AutoSize = true;
			this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.Location = new System.Drawing.Point(88, 356);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(77, 16);
			this.label21.TabIndex = 297;
			this.label21.Text = "DESCUADRE";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTotal
			// 
			this.txtVTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTotal.Appearance = appearance47;
			this.txtVTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTotal.FormatString = "#,##0.00";
			this.txtVTotal.Location = new System.Drawing.Point(176, 192);
			this.txtVTotal.Name = "txtVTotal";
			this.txtVTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTotal.PromptChar = ' ';
			this.txtVTotal.ReadOnly = true;
			this.txtVTotal.Size = new System.Drawing.Size(80, 21);
			this.txtVTotal.TabIndex = 296;
			// 
			// label20
			// 
			this.label20.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label20.AutoSize = true;
			this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(4, 194);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(166, 16);
			this.label20.TabIndex = 294;
			this.label20.Text = "TOTAL EFE. NOMINACIONES";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtMonedas
			// 
			this.txtMonedas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtMonedas.Appearance = appearance48;
			this.txtMonedas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMonedas.FormatString = "#,##0.00";
			this.txtMonedas.Location = new System.Drawing.Point(176, 164);
			this.txtMonedas.Name = "txtMonedas";
			this.txtMonedas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMonedas.PromptChar = ' ';
			this.txtMonedas.Size = new System.Drawing.Size(80, 21);
			this.txtMonedas.TabIndex = 293;
			this.txtMonedas.Click += new System.EventHandler(this.txtMonedas_Click);
			this.txtMonedas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtMonedas_KeyPress);
			this.txtMonedas.Validated += new System.EventHandler(this.txtMonedas_Validated);
			this.txtMonedas.Enter += new System.EventHandler(this.txtMonedas_Enter);
			// 
			// label18
			// 
			this.label18.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(119, 166);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(51, 16);
			this.label18.TabIndex = 291;
			this.label18.Text = "Monedas";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTB1
			// 
			this.txtVTB1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB1.Appearance = appearance49;
			this.txtVTB1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB1.FormatString = "#,##0.00";
			this.txtVTB1.Location = new System.Drawing.Point(176, 142);
			this.txtVTB1.Name = "txtVTB1";
			this.txtVTB1.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB1.PromptChar = ' ';
			this.txtVTB1.ReadOnly = true;
			this.txtVTB1.Size = new System.Drawing.Size(80, 21);
			this.txtVTB1.TabIndex = 290;
			// 
			// txtCantB1
			// 
			this.txtCantB1.Anchor = System.Windows.Forms.AnchorStyles.None;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB1.Appearance = appearance50;
			this.txtCantB1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB1.FormatString = "##0";
			this.txtCantB1.Location = new System.Drawing.Point(96, 142);
			this.txtCantB1.Name = "txtCantB1";
			this.txtCantB1.PromptChar = ' ';
			this.txtCantB1.Size = new System.Drawing.Size(56, 21);
			this.txtCantB1.TabIndex = 289;
			this.txtCantB1.Click += new System.EventHandler(this.txtCantB1_Click);
			this.txtCantB1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB1_KeyPress);
			this.txtCantB1.Validated += new System.EventHandler(this.txtCantB1_Validated);
			this.txtCantB1.Enter += new System.EventHandler(this.txtCantB1_Enter);
			// 
			// label17
			// 
			this.label17.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(64, 144);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(21, 16);
			this.label17.TabIndex = 288;
			this.label17.Text = "B.1";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTB5
			// 
			this.txtVTB5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB5.Appearance = appearance51;
			this.txtVTB5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB5.FormatString = "#,##0.00";
			this.txtVTB5.Location = new System.Drawing.Point(176, 120);
			this.txtVTB5.Name = "txtVTB5";
			this.txtVTB5.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB5.PromptChar = ' ';
			this.txtVTB5.ReadOnly = true;
			this.txtVTB5.Size = new System.Drawing.Size(80, 21);
			this.txtVTB5.TabIndex = 287;
			// 
			// txtCantB5
			// 
			this.txtCantB5.Anchor = System.Windows.Forms.AnchorStyles.None;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB5.Appearance = appearance52;
			this.txtCantB5.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB5.FormatString = "##0";
			this.txtCantB5.Location = new System.Drawing.Point(96, 120);
			this.txtCantB5.Name = "txtCantB5";
			this.txtCantB5.PromptChar = ' ';
			this.txtCantB5.Size = new System.Drawing.Size(56, 21);
			this.txtCantB5.TabIndex = 286;
			this.txtCantB5.Click += new System.EventHandler(this.txtCantB5_Click);
			this.txtCantB5.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB5_KeyPress);
			this.txtCantB5.Validated += new System.EventHandler(this.txtCantB5_Validated);
			this.txtCantB5.Enter += new System.EventHandler(this.txtCantB5_Enter);
			// 
			// label16
			// 
			this.label16.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(64, 122);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(21, 16);
			this.label16.TabIndex = 285;
			this.label16.Text = "B.5";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTB10
			// 
			this.txtVTB10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB10.Appearance = appearance53;
			this.txtVTB10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB10.FormatString = "#,##0.00";
			this.txtVTB10.Location = new System.Drawing.Point(176, 98);
			this.txtVTB10.Name = "txtVTB10";
			this.txtVTB10.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB10.PromptChar = ' ';
			this.txtVTB10.ReadOnly = true;
			this.txtVTB10.Size = new System.Drawing.Size(80, 21);
			this.txtVTB10.TabIndex = 284;
			// 
			// txtCantB10
			// 
			this.txtCantB10.Anchor = System.Windows.Forms.AnchorStyles.None;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB10.Appearance = appearance54;
			this.txtCantB10.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB10.FormatString = "##0";
			this.txtCantB10.Location = new System.Drawing.Point(96, 98);
			this.txtCantB10.Name = "txtCantB10";
			this.txtCantB10.PromptChar = ' ';
			this.txtCantB10.Size = new System.Drawing.Size(56, 21);
			this.txtCantB10.TabIndex = 283;
			this.txtCantB10.Click += new System.EventHandler(this.txtCantB10_Click);
			this.txtCantB10.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB10_KeyPress);
			this.txtCantB10.Validated += new System.EventHandler(this.txtCantB10_Validated);
			this.txtCantB10.Enter += new System.EventHandler(this.txtCantB10_Enter);
			// 
			// label15
			// 
			this.label15.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(58, 100);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(27, 16);
			this.label15.TabIndex = 282;
			this.label15.Text = "B.10";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTB20
			// 
			this.txtVTB20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB20.Appearance = appearance55;
			this.txtVTB20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB20.FormatString = "#,##0.00";
			this.txtVTB20.Location = new System.Drawing.Point(176, 76);
			this.txtVTB20.Name = "txtVTB20";
			this.txtVTB20.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB20.PromptChar = ' ';
			this.txtVTB20.ReadOnly = true;
			this.txtVTB20.Size = new System.Drawing.Size(80, 21);
			this.txtVTB20.TabIndex = 281;
			// 
			// txtCantB20
			// 
			this.txtCantB20.Anchor = System.Windows.Forms.AnchorStyles.None;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB20.Appearance = appearance56;
			this.txtCantB20.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB20.FormatString = "##0";
			this.txtCantB20.Location = new System.Drawing.Point(96, 76);
			this.txtCantB20.Name = "txtCantB20";
			this.txtCantB20.PromptChar = ' ';
			this.txtCantB20.Size = new System.Drawing.Size(56, 21);
			this.txtCantB20.TabIndex = 280;
			this.txtCantB20.Click += new System.EventHandler(this.txtCantB20_Click);
			this.txtCantB20.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB20_KeyPress);
			this.txtCantB20.Validated += new System.EventHandler(this.txtCantB20_Validated);
			this.txtCantB20.Enter += new System.EventHandler(this.txtCantB20_Enter);
			// 
			// label14
			// 
			this.label14.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(58, 78);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(27, 16);
			this.label14.TabIndex = 279;
			this.label14.Text = "B.20";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTB50
			// 
			this.txtVTB50.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB50.Appearance = appearance57;
			this.txtVTB50.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB50.FormatString = "#,##0.00";
			this.txtVTB50.Location = new System.Drawing.Point(176, 54);
			this.txtVTB50.Name = "txtVTB50";
			this.txtVTB50.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB50.PromptChar = ' ';
			this.txtVTB50.ReadOnly = true;
			this.txtVTB50.Size = new System.Drawing.Size(80, 21);
			this.txtVTB50.TabIndex = 278;
			// 
			// txtCantB50
			// 
			this.txtCantB50.Anchor = System.Windows.Forms.AnchorStyles.None;
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB50.Appearance = appearance58;
			this.txtCantB50.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB50.FormatString = "##0";
			this.txtCantB50.Location = new System.Drawing.Point(96, 54);
			this.txtCantB50.Name = "txtCantB50";
			this.txtCantB50.PromptChar = ' ';
			this.txtCantB50.Size = new System.Drawing.Size(56, 21);
			this.txtCantB50.TabIndex = 277;
			this.txtCantB50.Click += new System.EventHandler(this.txtCantB50_Click);
			this.txtCantB50.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB50_KeyPress);
			this.txtCantB50.Validated += new System.EventHandler(this.txtCantB50_Validated);
			this.txtCantB50.Enter += new System.EventHandler(this.txtCantB50_Enter);
			// 
			// label13
			// 
			this.label13.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(58, 56);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(27, 16);
			this.label13.TabIndex = 276;
			this.label13.Text = "B.50";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVTB100
			// 
			this.txtVTB100.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTB100.Appearance = appearance59;
			this.txtVTB100.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTB100.FormatString = "#,##0.00";
			this.txtVTB100.Location = new System.Drawing.Point(176, 32);
			this.txtVTB100.Name = "txtVTB100";
			this.txtVTB100.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTB100.PromptChar = ' ';
			this.txtVTB100.ReadOnly = true;
			this.txtVTB100.Size = new System.Drawing.Size(80, 21);
			this.txtVTB100.TabIndex = 275;
			// 
			// label12
			// 
			this.label12.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label12.AutoSize = true;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(208, 16);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(44, 16);
			this.label12.TabIndex = 274;
			this.label12.Text = "VALOR";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtCantB100
			// 
			this.txtCantB100.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtCantB100.Appearance = appearance60;
			this.txtCantB100.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCantB100.FormatString = "##0";
			this.txtCantB100.Location = new System.Drawing.Point(96, 32);
			this.txtCantB100.Name = "txtCantB100";
			this.txtCantB100.PromptChar = ' ';
			this.txtCantB100.Size = new System.Drawing.Size(56, 21);
			this.txtCantB100.TabIndex = 273;
			this.txtCantB100.Click += new System.EventHandler(this.txtCantB100_Click);
			this.txtCantB100.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCantB100_KeyPress);
			this.txtCantB100.Validated += new System.EventHandler(this.txtCantB100_Validated);
			this.txtCantB100.Enter += new System.EventHandler(this.txtCantB100_Enter);
			// 
			// label11
			// 
			this.label11.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(51, 34);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(34, 16);
			this.label11.TabIndex = 272;
			this.label11.Text = "B.100";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtGasto
			// 
			this.txtGasto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Right;
			this.txtGasto.Appearance = appearance61;
			this.txtGasto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGasto.FormatString = "#,##0.00";
			this.txtGasto.Location = new System.Drawing.Point(176, 264);
			this.txtGasto.Name = "txtGasto";
			this.txtGasto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGasto.PromptChar = ' ';
			this.txtGasto.Size = new System.Drawing.Size(80, 21);
			this.txtGasto.TabIndex = 271;
			this.txtGasto.Click += new System.EventHandler(this.txtGasto_Click);
			this.txtGasto.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtGasto_KeyPress);
			this.txtGasto.Validated += new System.EventHandler(this.txtGasto_Validated);
			this.txtGasto.AfterEnterEditMode += new System.EventHandler(this.txtGasto_AfterEnterEditMode);
			this.txtGasto.Enter += new System.EventHandler(this.txtGasto_Enter);
			// 
			// label10
			// 
			this.label10.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(91, 266);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(73, 16);
			this.label10.TabIndex = 270;
			this.label10.Text = "Valor GASTO";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtVoucher
			// 
			this.txtVoucher.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVoucher.Appearance = appearance62;
			this.txtVoucher.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVoucher.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVoucher.FormatString = "#,##0.00";
			this.txtVoucher.Location = new System.Drawing.Point(192, 216);
			this.txtVoucher.Name = "txtVoucher";
			this.txtVoucher.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVoucher.PromptChar = ' ';
			this.txtVoucher.ReadOnly = true;
			this.txtVoucher.Size = new System.Drawing.Size(80, 21);
			this.txtVoucher.TabIndex = 269;
			// 
			// label9
			// 
			this.label9.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(144, 218);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(46, 16);
			this.label9.TabIndex = 268;
			this.label9.Text = "Voucher";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtEfectivo
			// 
			this.txtEfectivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEfectivo.Appearance = appearance63;
			this.txtEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEfectivo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEfectivo.FormatString = "#,##0.00";
			this.txtEfectivo.Location = new System.Drawing.Point(64, 216);
			this.txtEfectivo.Name = "txtEfectivo";
			this.txtEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtEfectivo.PromptChar = ' ';
			this.txtEfectivo.ReadOnly = true;
			this.txtEfectivo.Size = new System.Drawing.Size(80, 21);
			this.txtEfectivo.TabIndex = 267;
			// 
			// label8
			// 
			this.label8.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(16, 218);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(45, 16);
			this.label8.TabIndex = 266;
			this.label8.Text = "Efectivo";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn51.DefaultValue = 0;
			ultraDataColumn53.DataType = typeof(System.DateTime);
			ultraDataColumn53.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn55.DefaultValue = 0;
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn56.DefaultValue = 0;
			ultraDataColumn57.DataType = typeof(System.DateTime);
			ultraDataColumn57.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn58.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn59.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn60.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn61.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn62.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn63.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn64.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn65.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn66.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn67.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53,
																																 ultraDataColumn54,
																																 ultraDataColumn55,
																																 ultraDataColumn56,
																																 ultraDataColumn57,
																																 ultraDataColumn58,
																																 ultraDataColumn59,
																																 ultraDataColumn60,
																																 ultraDataColumn61,
																																 ultraDataColumn62,
																																 ultraDataColumn63,
																																 ultraDataColumn64,
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67});
			// 
			// frmRepoCajaLotte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1200, 549);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.simpleButton1);
			this.Controls.Add(this.lblEdoDpto);
			this.Controls.Add(this.rb3);
			this.Controls.Add(this.rb2);
			this.Controls.Add(this.rb1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.cmbCuentaCorriente);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.cmdOperador);
			this.Controls.Add(this.uGridDepositos);
			this.Controls.Add(this.cmbCajero);
			this.Controls.Add(this.ugCaja);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.btVer);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btExportar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFeFinaliza);
			this.Name = "frmRepoCajaLotte";
			this.Text = "Reporte de CAJA";
			this.Load += new System.EventHandler(this.frmRepoCajaLotte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFeFinaliza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCajero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdOperador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorriente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtValorPrenda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepositoReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCheque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotVEnta)).EndInit();
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
			((System.ComponentModel.ISupportInitialize)(this.txtVoucher)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int idCaja = 0;
		int idDesCaja = 0;
		string niUsuario = "";
		int NBodga = 0;
		decimal VlrDpositado = 0;

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void frmRepoCajaLotte_Load(object sender, System.EventArgs e)
		{
			//string sSQL = string.Format("Exec ListaBodegas {0}, 45", MenuLatinium.IdUsuario);
			//this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
			this.dtFecha.Value = DateTime.Today;
			this.dtFeFinaliza.Value = DateTime.Today;
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Parse("07/12/2017");
			string stExecBodega = string.Format("Select idBodega, Bodega, Codigo, Nombre From Bodega Where Activo = 1");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteo, stExecBodega);

			string stUser = string.Format("Select idPersonal, Codigo, Nombre From Personal Where FechaSalida is Null");			
			this.cmdOperador.DataSource = Funcion.dvProcedimiento(cdsSeteo, stUser);
			this.cmbCuentaCorriente.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Exec CuantaCorrienteLotte");
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteo, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 2) Order By idFormaPago");
			
			//this.dtFecha.Value = DateTime.Today;
			//this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			int codBodHabtcn = Funcion.iEscalarSQL(cdsSeteo, "SELECT Bodega FROM UsuariosBodegas WHERE idUsuario="+MenuLatinium.IdUsuario.ToString());
			int BodaBusk = Funcion.iEscalarSQL(cdsSeteo, "SELECT Bodega FROM Bodega WHERE Bodega="+codBodHabtcn.ToString());
			this.cmbBodega.Value = BodaBusk;

			niUsuario = MenuLatinium.iNivel.ToString().Trim();
			if(niUsuario.Equals("4") || niUsuario.Equals("1") || niUsuario.Equals("2") || niUsuario.Equals("3")) 
			{
				this.cmbBodega.Enabled = true;
			}
		}

		private void btVer_Click(object sender, System.EventArgs e)
 		{
			this.txtEfectivo.Value = 0.00m;
			this.txtVoucher.Value = 0.00m;
			this.txtVTotal.Value = 0.00m;
			this.txtTotVEnta.Value = 0.00m;
			this.txtDescuadre.Value = 0;
			this.txtValorPrenda.Value = 0.00m;
			this.txtDepositoReserva.Value = 0.00m;
			this.txtVTB100.Value = 0;
			this.txtVTB50.Value = 0;
			this.txtVTB20.Value = 0;
			this.txtVTB10.Value = 0;
			this.txtVTB5.Value = 0;
			this.txtVTB1.Value = 0;
			this.txtCantB1.Value = 0;
			this.txtCantB5.Value = 0;
			this.txtCantB10.Value = 0;
			this.txtCantB20.Value = 0;
			this.txtCantB50.Value = 0;
			this.txtCantB100.Value = 0;
			this.txtMonedas.Value = 0;
			this.txtGasto.Value = 0;
			this.txtCheque.Value = 0;
			VlrDpositado = 0;

			if(this.rb3.Checked) this.btnGuardar.Enabled = false;
			else this.btnGuardar.Enabled = true;

			DateTime FeInicia = (DateTime)this.dtFecha.Value;
			DateTime FeFinaliza = (DateTime)this.dtFeFinaliza.Value;
			int idCajro = 0;
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese una fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();				
			}
			else if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
			}
			else
			{
				DateTime dtFecha = (DateTime) this.dtFecha.Value;
				DateTime dtFFin = (DateTime)this.dtFeFinaliza.Value;
				if (this.cmbCajero.ActiveRow != null) idCajro = (int)this.cmbCajero.Value;
				
				int Turno = 0;
				if(this.rb1.Checked) Turno=1;
				else if(this.rb2.Checked) Turno=2;
				else Turno = 0;
			
				/*Facturas*/
				string sSQL = string.Format("Exec ReporteCaja_Lotte '{0}','{1}',{2}, {3}, {4}", 
					dtFecha.ToString("yyyyMMdd"), 
					dtFFin.ToString("yyyyMMdd"), 
					(int)this.cmbBodega.Value, idCajro, Turno);
				this.ugCaja.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
				
				/*Depositos*/
				string sSQLDep = string.Format("Exec ReporteCajaDepo {0},'{1}',{2}", 
					(int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"), Turno);
				this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQLDep);
				int idCja = 0;
				this.cmdOperador.Value = 0;
				this.txtObservaciones.Value = "";

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
				{	
					idCja = Convert.ToInt32(dr.Cells["idCaja"].Value);
					idDesCaja = Convert.ToInt32(dr.Cells["idDetMovCaja"].Value);
					VlrDpositado = VlrDpositado+Convert.ToDecimal(dr.Cells["Valor"].Value);
				}

				/*Detalle Deposito*/
				string sSQLDMC = string.Format("Exec DetCajaMovimiento {0},{1},'{2}',{3}", 
					idCja,Turno,dtFecha.ToString("yyyyMMdd"),(int)this.cmbBodega.Value);
				SqlDataReader drMovCaja = Funcion.rEscalarSQL(cdsSeteo, sSQLDMC, true);
//				drMovCaja.Read();

				int b100 = 0;
				int b50 = 0;
				int b20 = 0;
				int b10 = 0;
				int b5 = 0;
				int b1 = 0;
				decimal tmods = 0.00m;
				decimal gsto = 0.00m;
				decimal chqe = 0.00m;
				decimal resrv = 0.00m;
				decimal prnda = 0.00m;

				while (drMovCaja.Read())
				{
					b100 = b100 + drMovCaja.GetInt32(3);
					b50 = b50 + drMovCaja.GetInt32(4);
					b20 = b20 + drMovCaja.GetInt32(5);
					b10 = b10 + drMovCaja.GetInt32(6);
					b5 = b5 + drMovCaja.GetInt32(7);
					b1 = b1 + drMovCaja.GetInt32(8);
					tmods = tmods + drMovCaja.GetDecimal(9);
					gsto = gsto + drMovCaja.GetDecimal(10);
					chqe = chqe + drMovCaja.GetDecimal(11);
					this.cmdOperador.Value = Convert.ToInt32(drMovCaja.GetInt32(13));
					this.txtObservaciones.Value = drMovCaja.GetString(14);
					resrv = resrv + drMovCaja.GetDecimal(15);
					prnda = prnda + drMovCaja.GetDecimal(16);
					//drMovCaja.NextResult();
				}
				drMovCaja.Close();

				this.txtCantB100.Value = b100;
				this.txtCantB50.Value = b50;
				this.txtCantB20.Value = b20;
				this.txtCantB10.Value = b10;
				this.txtCantB5.Value = b5;
				this.txtCantB1.Value = b1;
				this.txtMonedas.Value = tmods;
				this.txtGasto.Value = gsto;
				this.txtCheque.Value = chqe;
				this.txtDepositoReserva.Value = resrv;
				this.txtValorPrenda.Value = prnda;
				
				/*Descuadre Caja
				string sSQLDescdr = string.Format("Exec ReporteCajaDescdr {0}", idCja);
				SqlDataReader drDCaj = Funcion.rEscalarSQL(cdsSeteo, sSQLDescdr, true);
				drDCaj.Read();
				if (drDCaj.HasRows)
				{					
					idDesCaja = drDCaj.GetInt32(0);
					this.cmdOperador.Value = drDCaj.GetInt32(2);
					this.txtObservaciones.Value = drDCaj.GetString(4);
					this.txtValor.Value = drDCaj.GetDecimal(5);
				}
				drDCaj.Close();
				*/
				/*Calcular Efectivo y Voucher*/
				Total();
				//this.btnGuardar.Enabled = true;
			}
		}

		private void Total()
		{
			#region Creacion de Variables					
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;

			decimal dSubTotal1 = 0.00m;
			decimal dSubTotal2 = 0.00m;
			decimal dSubTotal3 = 0.00m;
			decimal dSubTotal4 = 0.00m;
			#endregion Creacion de Variables
			
			this.txtEfectivo.Value = 0.00m;
			this.txtVoucher.Value = 0.00m;
			this.txtVTotal.Value = 0.00m;
			this.txtTotVEnta.Value = 0.00m;
			this.txtDescuadre.Value = 0.00m;
			this.txtValorPrenda.Value = 0.00m;
			this.txtDepositoReserva.Value = 0.00m;

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugCaja.Rows.All)
			{
				if (dr.Cells["EFE"].Value != DBNull.Value)
					this.txtEfectivo.Value = (decimal)this.txtEfectivo.Value + Math.Round(Convert.ToDecimal(dr.Cells["EFE"].Value), 2);
				if (dr.Cells["TC"].Value != DBNull.Value)
					this.txtVoucher.Value = (decimal)this.txtVoucher.Value + Math.Round(Convert.ToDecimal(dr.Cells["TC"].Value), 2) + Math.Round(Convert.ToDecimal(dr.Cells["TD"].Value), 2);
				if (dr.Cells["TOTAL"].Value != DBNull.Value)
					this.txtTotVEnta.Value = (decimal)this.txtTotVEnta.Value + Math.Round(Convert.ToDecimal(dr.Cells["TOTAL"].Value), 2);
				if (dr.Cells["PRENDA"].Value != DBNull.Value)
					this.txtValorPrenda.Value = (decimal)this.txtValorPrenda.Value + Math.Round(Convert.ToDecimal(dr.Cells["PRENDA"].Value), 2);
				if (dr.Cells["DEPOSITO"].Value != DBNull.Value)
					this.txtDepositoReserva.Value = (decimal)this.txtDepositoReserva.Value + Math.Round(Convert.ToDecimal(dr.Cells["DEPOSITO"].Value), 2);
			}

			this.txtVTB100.Value = Convert.ToDecimal(this.txtCantB100.Value)*100;
			this.txtVTB50.Value = Convert.ToDecimal(this.txtCantB50.Value)*50;
			this.txtVTB20.Value = Convert.ToDecimal(this.txtCantB20.Value)*20;
			this.txtVTB10.Value = Convert.ToDecimal(this.txtCantB10.Value)*10;
			this.txtVTB5.Value = Convert.ToDecimal(this.txtCantB5.Value)*5;
			this.txtVTB1.Value = Convert.ToDecimal(this.txtCantB1.Value)*1;

			this.txtVTotal.Value = Convert.ToDecimal(this.txtVTB100.Value)+
				Convert.ToDecimal(this.txtVTB50.Value)+
				Convert.ToDecimal(this.txtVTB20.Value)+
				Convert.ToDecimal(this.txtVTB10.Value)+
				Convert.ToDecimal(this.txtVTB5.Value)+
				Convert.ToDecimal(this.txtVTB1.Value)+
				Convert.ToDecimal(this.txtMonedas.Value);

			this.txtDescuadre.Value = Convert.ToDecimal(this.txtEfectivo.Value)-(VlrDpositado+Convert.ToDecimal(this.txtGasto.Value));

//			this.txtDescuadre.Value = Convert.ToDecimal(this.txtTotVEnta.Value)-
//				(Convert.ToDecimal(this.txtVTotal.Value)+
//				Convert.ToDecimal(this.txtGasto.Value)+
//				Convert.ToDecimal(this.txtCheque.Value)+
//				Convert.ToDecimal(this.txtVoucher.Value)+
//				Convert.ToDecimal(this.txtDepositoReserva.Value)+
//				Convert.ToDecimal(this.txtValorPrenda.Value));

			this.txtValor.Value = Convert.ToDecimal(this.txtDescuadre.Value);
			#endregion Desplazamiento de grilla			
		}

		private void btExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ugCaja.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.ugCaja, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			/*Obtiene la bodega*/
			int idBodg = (int)this.cmbBodega.Value;

			this.cmbCajero.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec PersonalLotte_Cajero {0}", idBodg));
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una bodega.", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();				
			}
			else if (this.ugCaja.Rows.Count == 0)
			{
				MessageBox.Show("No Existen datos de caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
//			else if (this.uGridDepositos.Rows.Count == 0)
//			{
//				MessageBox.Show("No Existen documentos para realizar el deposito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
//			}
			else if (this.rb3.Checked)
			{
				MessageBox.Show("Debe seleccionar TURNO 1 o 2 Para Generar Depósito.?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{

				//DateTime FechaCaja = DateTime.Today;			
				DateTime FechaCaja = (DateTime)this.dtFecha.Value;

				#region Proyecto
				string sSQLProyecto = string.Format("Exec BodegaCentroDeCosto {0}", (int)this.cmbBodega.Value);
				int iProyecto = Funcion.iEscalarSQL(cdsSeteo, sSQLProyecto, true);
				#endregion Proyecto

				decimal dValDeposito = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugCaja.Rows.All)
				{
					dValDeposito = dValDeposito + Convert.ToDecimal(dr.Cells["EFE"].Value); 
				}

				DateTime FechaDpo = DateTime.Today;
				string NumDepo = "";
				decimal ValEfec = 0.00m;

				int Turno = 0;
				if(this.rb1.Checked) Turno=1;
				else if(this.rb2.Checked) Turno=2;
				else Turno = 0;

				/*Obtener el IdCAJA para generar DEPOSITO*/
				string sSQLidCja = string.Format("Exec ObteneridCaja '{0}',{1},{2}",
					FechaCaja.ToString("yyyyMMdd"),Turno,(int)this.cmbBodega.Value);			
				idCaja = Funcion.iEscalarSQL(cdsSeteo, sSQLidCja);

				//if(dValDeposito!=0)
				//{
				if(this.uGridDepositos.Rows.Count>0)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDet in this.uGridDepositos.Rows.All)
					{
						if(idCaja==0)
						{
							if(drDet.Cells["idCaja"].Value == System.DBNull.Value) idCaja = 0;
							else idCaja = (int)drDet.Cells["idCaja"].Value;
						}
						if((int)drDet.Cells["idFormaPago"].Value==1)
						{
							FechaDpo = (DateTime)drDet.Cells["FechaDeposito"].Value; 
							ValEfec = ValEfec + (decimal)drDet.Cells["Valor"].Value;
						}
						string sSQLMaestro = string.Format("Exec GuardaCajaLotte {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}", 
							idCaja, (int)this.cmbBodega.Value, iProyecto, FechaCaja.ToString("yyyyMMdd"), FechaDpo.ToString("yyyyMMdd"), 
							NumDepo, ValEfec, 0, dValDeposito, 1, false, 
							this.txtObservaciones.Text.ToString(), Turno);			
						idCaja = Funcion.iEscalarSQL(cdsSeteo, sSQLMaestro);
					}
				} 
				else 
				{
					string sSQLMaestro = string.Format("Exec GuardaCajaLotte {0}, {1}, {2}, '{3}', '{4}', '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}", 
						idCaja, (int)this.cmbBodega.Value, iProyecto, FechaCaja.ToString("yyyyMMdd"), FechaDpo.ToString("yyyyMMdd"), 
						NumDepo, ValEfec, 0, dValDeposito, 1, false, 
						this.txtObservaciones.Text.ToString(), Turno);			
					idCaja = Funcion.iEscalarSQL(cdsSeteo, sSQLMaestro);
				}

				decimal ValDepositar = 0.00m;
			
				#region Detalle
				if(this.uGridDepositos.Rows.Count>0)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
					{
						#region Validacion
						if (dr.Cells["idCuenta"].Value == System.DBNull.Value)
						{
							MessageBox.Show("Seleccione la cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
						if (dr.Cells["Numero"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Escriba el número de depósito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
						if (dr.Cells["FechaDeposito"].Value == System.DBNull.Value)
						{
							MessageBox.Show("Ingrese la fecha de depósito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
						#endregion Validacion

						DateTime dtFechaDep = (DateTime)dr.Cells["FechaDeposito"].Value;
						int dcaj = 0;
						if (dr.Cells["idDetalleCaja"].Value == System.DBNull.Value) dcaj = 0;
						else dcaj = (int)dr.Cells["idDetalleCaja"].Value;
						int icta = (int)dr.Cells["idCuenta"].Value;
						int ifpa = (int)dr.Cells["idFormaPago"].Value;
						string nro = dr.Cells["Numero"].Value.ToString();
						decimal valdep = (decimal)dr.Cells["Valor"].Value;

						ValDepositar = ValDepositar+valdep;

						string sSQLDetalleEf = string.Format("Exec GuardaDetCajaLoote {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
							dcaj, idCaja, icta, ifpa, nro, 
							dtFechaDep.ToString("yyyyMMdd"), valdep);
						Funcion.EjecutaSQL(cdsSeteo, sSQLDetalleEf, true);
					}
				} 
				else 
				{
					int dcaj = 0;
					int icta = 0;
					int ifpa = 0;
					string nro = "";
					decimal valdep = 0;
					DateTime dtFechaDep = DateTime.Today;
					string sSQLDetalleEf = string.Format("Exec GuardaDetCajaLoote {0}, {1}, {2}, {3}, '{4}', '{5}', {6}", 
						dcaj, idCaja, icta, ifpa, nro, 
						dtFechaDep.ToString("yyyyMMdd"), valdep);
					Funcion.EjecutaSQL(cdsSeteo, sSQLDetalleEf, true);
				}
				
				//this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteo, string.Format("Exec ConsultaDetalleCaja {0}", idCaja));
				DateTime dtFecha = (DateTime) this.dtFecha.Value;
				string sSQLDep = string.Format("Exec ReporteCajaDepo {0},'{1}',{2}", 
					(int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"), Turno);
				this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQLDep);
				this.lblEdoDpto.Text = Funcion.sEscalarSQL(cdsSeteo, string.Format("Exec MensajeCaja {0}", idCaja));			
				#endregion Detalle

				#region registro movimiento de caja
				DateTime FeDesc = DateTime.Today;

				int idPerDes = 0;
				if (this.cmdOperador.Value == null) idPerDes = 0;
				else idPerDes = (int)this.cmdOperador.Value;

				decimal valDesc = Convert.ToDecimal(this.txtDescuadre.Value);
				int Billete100 = Convert.ToInt32(this.txtCantB100.Value);
				int Billete50 = Convert.ToInt32(this.txtCantB50.Value);
				int Billete20 = Convert.ToInt32(this.txtCantB20.Value);
				int Billete10 = Convert.ToInt32(this.txtCantB10.Value);
				int Billete5 = Convert.ToInt32(this.txtCantB5.Value);
				int Billete1 = Convert.ToInt32(this.txtCantB1.Value);
				decimal VMonedas = Convert.ToDecimal(this.txtMonedas.Value);
				decimal VGasto = Convert.ToDecimal(this.txtGasto.Value);
				decimal VCheque = Convert.ToDecimal(this.txtCheque.Value);
				decimal VDpResrv = Convert.ToDecimal(this.txtDepositoReserva.Value);
				decimal VPrnda = Convert.ToDecimal(this.txtValorPrenda.Value);
				decimal VEfectvo = Convert.ToDecimal(this.txtEfectivo.Value);

				string sSQLDescdr = string.Format("Exec RegDescuadre {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}", 
					idCaja, idPerDes, dtFecha.ToString("yyyyMMdd"), this.txtObservaciones.Text.ToString(),valDesc,
					Billete100,Billete50,Billete20,Billete10,Billete5,Billete1,VMonedas,VGasto,VCheque,idDesCaja,VDpResrv,VPrnda,ValDepositar,VEfectvo);
				Funcion.EjecutaSQL(cdsSeteo, sSQLDescdr, true);
				#endregion registro movimiento de caja

				/*Genera asiento contable*/
				if(dValDeposito!=0)	GeneraAsiento();
				this.btnGuardar.Enabled = false;
				
				//}
			}
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
				Funcion.EjecutaSQL(cdsSeteo, sSQL, true);
			}			

			string sSQLDiferencias = string.Format("Exec AsientoDiferenciasDeCaja {0}", idCaja);
			Funcion.EjecutaSQL(cdsSeteo, sSQLDiferencias, true);
			#endregion Crea Asiento
		}

		private void simpleButton1_Click(object sender, System.EventArgs e)
		{
			//Visualizar información
			#region Impresión
			if(this.rb1.Checked==true || this.rb2.Checked==true) Imprimir(0);
			else Imprimir(0);
			#endregion Impresión
		}

		private void Imprimir(int iTipo)
		{
			int idCpaFC = 0;

			#region Imprime el reporte
			//using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			DateTime FeInicia = (DateTime)this.dtFecha.Value;
			DateTime FeFinaliza = (DateTime)this.dtFeFinaliza.Value;
			int Turno = 0;
			if(this.rb1.Checked) Turno=1;
			else if(this.rb2.Checked) Turno=2;
			else Turno = 0;
			int laBodeg = (int)this.cmbBodega.Value;
			int elCajero = 0;
			if (this.cmbCajero.ActiveRow == null) elCajero = 0;
			else elCajero = (int)this.cmbCajero.Value;

			if (iTipo == 0) 
			{
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldDesde = new ParameterField ();
				ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
				paramFieldDesde.ParameterFieldName = "@FeBusk";
				discreteValDesde.Value = FeInicia;
				paramFieldDesde.CurrentValues.Add (discreteValDesde);
				paramFields.Add (paramFieldDesde);

				ParameterField paramFieldHasta = new ParameterField ();
				ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
				paramFieldHasta.ParameterFieldName = "@FeFin";
				discreteValHasta.Value = FeFinaliza;
				paramFieldHasta.CurrentValues.Add (discreteValHasta);
				paramFields.Add (paramFieldHasta);

				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idBodg = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@Bodga";
				idBodg.Value = laBodeg;
				pfRet.CurrentValues.Add (idBodg);
				paramFields.Add (pfRet);
				
				ParameterField idCjero = new ParameterField ();
				ParameterDiscreteValue idCjr = new ParameterDiscreteValue ();
				idCjero.ParameterFieldName = "@idCajro";
				idCjr.Value = elCajero;
				idCjero.CurrentValues.Add (idCjr);
				paramFields.Add (idCjero);

				ParameterField nTurn = new ParameterField ();
				ParameterDiscreteValue nTurnCaj = new ParameterDiscreteValue ();
				nTurn.ParameterFieldName = "@Trn";
				nTurnCaj.Value = Turno;
				nTurn.CurrentValues.Add (nTurnCaj);
				paramFields.Add (nTurn);
				
				Reporte miRepor = new Reporte("ReporteCajaLotte.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Retenciones");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				#endregion Impresion
				
				miRepor.Show();

			}
			if (iTipo == 1) 
			{
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldDesde = new ParameterField ();
				ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
				paramFieldDesde.ParameterFieldName = "@FeBusk";
				discreteValDesde.Value = FeInicia;
				paramFieldDesde.CurrentValues.Add (discreteValDesde);
				paramFields.Add (paramFieldDesde);

				ParameterField paramFieldHasta = new ParameterField ();
				ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
				paramFieldHasta.ParameterFieldName = "@FeFin";
				discreteValHasta.Value = FeFinaliza;
				paramFieldHasta.CurrentValues.Add (discreteValHasta);
				paramFields.Add (paramFieldHasta);

				ParameterField pfRet = new ParameterField ();
				ParameterDiscreteValue idBodg = new ParameterDiscreteValue ();
				pfRet.ParameterFieldName = "@Bodga";
				idBodg.Value = laBodeg;
				pfRet.CurrentValues.Add (idBodg);
				paramFields.Add (pfRet);
				
				ParameterField idCjero = new ParameterField ();
				ParameterDiscreteValue idCjr = new ParameterDiscreteValue ();
				idCjero.ParameterFieldName = "@idCajro";
				idCjr.Value = elCajero;
				idCjero.CurrentValues.Add (idCjr);
				paramFields.Add (idCjero);

				ParameterField nTurn = new ParameterField ();
				ParameterDiscreteValue nTurnCaj = new ParameterDiscreteValue ();
				nTurn.ParameterFieldName = "@Trn";
				nTurnCaj.Value = Turno;
				nTurn.CurrentValues.Add (nTurnCaj);
				paramFields.Add (nTurn);
				
				Reporte miRepor = new Reporte("RpCjaLot2.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Retenciones");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				#endregion Impresion
				
				miRepor.Show();
			}
			if (iTipo == 2)
			{
				DateTime dtFecha = (DateTime) this.dtFecha.Value;

				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue discreteValFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@Fecha";
				discreteValFecha.Value = dtFecha; //.ToString("yyyyMMdd");
				paramFieldFecha.CurrentValues.Add (discreteValFecha);
				paramFields.Add (paramFieldFecha);

				ParameterField paramFieldBodega = new ParameterField ();
				ParameterDiscreteValue discreteValBodega = new ParameterDiscreteValue ();
				paramFieldBodega.ParameterFieldName = "@Bodega";
				discreteValBodega.Value = (string)this.cmbBodega.Text; //int.Parse(this.cmbBodega.Value.ToString());
				paramFieldBodega.CurrentValues.Add (discreteValBodega);
				paramFields.Add (paramFieldBodega);

				ParameterField paramFieldidCaja = new ParameterField ();
				ParameterDiscreteValue discreteValidCaja = new ParameterDiscreteValue ();
				paramFieldidCaja.ParameterFieldName = "idCaja";
				discreteValidCaja.Value = idCaja.ToString();
				paramFieldidCaja.CurrentValues.Add (discreteValidCaja);
				paramFields.Add (paramFieldidCaja);

				Reporte miRep = new Reporte("CierreDeCajaLotte.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Cierre de Caja");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
			}
			#endregion Imprime el reporte
		}

		private void txtGasto_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtCheque.Focus();
		}

		private void txtGasto_AfterEnterEditMode(object sender, System.EventArgs e)
		{
		  Total();
		}

		private void txtCantB100_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		  if (e.KeyChar == 13) this.txtCantB50.Focus();
		}

		private void txtCantB100_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCantB50_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCantB20_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCantB10_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCantB5_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCantB1_Validated(object sender, System.EventArgs e)
		{
			Total();
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

		private void txtMonedas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtGasto.Focus();
		}

		private void txtMonedas_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCantB1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtMonedas.Focus();
		}

		private void txtGasto_Validated(object sender, System.EventArgs e)
		{
			Total();		
		}

		private void txtCheque_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtCheque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDepositoReserva.Focus();
		}

		private void txtCantB100_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB100.SelectAll();
		}

		private void txtCantB100_Click(object sender, System.EventArgs e)
		{
		  this.txtCantB100.SelectAll();
		}

		private void txtCantB50_Click(object sender, System.EventArgs e)
		{
			this.txtCantB50.SelectAll();
		}

		private void txtCantB50_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB50.SelectAll();
		}

		private void txtCantB20_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB20.SelectAll();
		}

		private void txtCantB20_Click(object sender, System.EventArgs e)
		{
			this.txtCantB20.SelectAll();
		}

		private void txtCantB10_Click(object sender, System.EventArgs e)
		{
			this.txtCantB10.SelectAll();
		}

		private void txtCantB10_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB10.SelectAll();
		}

		private void txtCantB5_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB5.SelectAll();
		}

		private void txtCantB5_Click(object sender, System.EventArgs e)
		{
			this.txtCantB5.SelectAll();
		}

		private void txtCantB1_Click(object sender, System.EventArgs e)
		{
			this.txtCantB1.SelectAll();
		}

		private void txtCantB1_Enter(object sender, System.EventArgs e)
		{
			this.txtCantB1.SelectAll();
		}

		private void txtMonedas_Enter(object sender, System.EventArgs e)
		{
		  this.txtMonedas.SelectAll();
		}

		private void txtMonedas_Click(object sender, System.EventArgs e)
		{
			this.txtMonedas.SelectAll();
		}

		private void txtGasto_Click(object sender, System.EventArgs e)
		{
			this.txtGasto.SelectAll();
		}

		private void txtGasto_Enter(object sender, System.EventArgs e)
		{
			this.txtGasto.SelectAll();
		}

		private void txtCheque_Enter(object sender, System.EventArgs e)
		{
			this.txtCheque.SelectAll();
		}

		private void txtCheque_Click(object sender, System.EventArgs e)
		{
			this.txtCheque.SelectAll();
		}

		private void uGridDepositos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void uGridDepositos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			for (int i = 0; i < e.Rows.Length; i++)
			{
				if (DialogResult.Yes ==	MessageBox.Show("Desea Eliminar Registro.?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{
					Funcion.EjecutaSQL(cdsSeteo, string.Format("Delete From AsientoDetalle Where idAsiento = {0} Delete From Asiento Where idAsiento = {0}", (int)e.Rows[i].Cells["idAsiento"].Value));
					Funcion.EjecutaSQL(cdsSeteo, string.Format("Delete From DetalleCaja Where idDetalleCaja = {0}", (int)e.Rows[i].Cells["idDetalleCaja"].Value));
					e.DisplayPromptMsg = false;
				}
				else e.Cancel = true;
			}
		}

		private void uGridDepositos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridDepositos.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{
						#region Enter
						case (int) Keys.Enter:
							uGridDepositos.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGridDepositos.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							uGridDepositos.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		private void txtDepositoReserva_Validated(object sender, System.EventArgs e)
		{
			Total();
		}

		private void txtDepositoReserva_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtValorPrenda.Focus();
		}

		private void txtDepositoReserva_Enter(object sender, System.EventArgs e)
		{
			this.txtDepositoReserva.SelectAll();
		}

		private void txtDepositoReserva_Click(object sender, System.EventArgs e)
		{
			this.txtDepositoReserva.SelectAll();		
		}

		private void txtValorPrenda_Click(object sender, System.EventArgs e)
		{
		  this.txtValorPrenda.SelectAll();
		}

		private void txtValorPrenda_Enter(object sender, System.EventArgs e)
		{
			this.txtValorPrenda.SelectAll();
		}

		private void txtValorPrenda_Validated(object sender, System.EventArgs e)
		{
			Total();	
		}

		private void txtValorPrenda_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtObservaciones.Focus();
		}

	}
}
