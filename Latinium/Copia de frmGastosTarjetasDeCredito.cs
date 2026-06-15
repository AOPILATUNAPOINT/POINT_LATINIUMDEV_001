using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmGastosTarjetasDeCredito.
	/// </summary>
	public class frmGastosTarjetasDeCredito : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblTarjeta;
		private System.Windows.Forms.Label lblBanco;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbTarjeta;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.Label lblSubTotal;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGastos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnProcesar;
		private System.Windows.Forms.Button btnActualizar;
		private System.Windows.Forms.Button btnCalcular;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentroDeCosto;
		private System.Windows.Forms.Label label2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalVoucher;
		private System.Windows.Forms.ImageList imageList1;
		private System.ComponentModel.IContainer components;

		public frmGastosTarjetasDeCredito()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGastosTarjetasDeCredito));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleGastosTarjetas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosTarjetas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SinFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Point");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbCentroDeCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA1");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal1");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA2");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal2");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA3");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal3");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA4");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal4");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA5");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal5");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA6");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal6");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA7");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal7");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA8");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal8");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA9");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal9");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA10");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal10");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA11");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal11");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA12");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal12");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA13");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal13");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA14");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal14");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA15");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal15");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIVA0");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalIVA");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorTotal");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIVA", 76, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIVA", 76, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalIVA0", 75, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalIVA0", 75, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 77, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 77, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 78, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 78, true);
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorTotal", 80, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorTotal", 80, true);
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorCuotas", 83, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorCuotas", 83, true);
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 84, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 84, true);
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleGastosTarjetas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosTarjetas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SinFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Point");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalIVA0");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTarjetasCorporativas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTarjetasCorporativas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblTarjeta = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.cmbTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.txtIdGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.btnCalcular = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCentroDeCosto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.txtTotalVoucher = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalVoucher)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.label3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.label3.Location = new System.Drawing.Point(1080, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 253;
			this.label3.Text = "Número";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.label1.Location = new System.Drawing.Point(648, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(33, 16);
			this.label1.TabIndex = 248;
			this.label1.Text = "Notas";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.label6.Location = new System.Drawing.Point(8, 75);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 245;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblEstado.Location = new System.Drawing.Point(392, 64);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 254;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtNumero.Location = new System.Drawing.Point(1168, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 21);
			this.txtNumero.TabIndex = 252;
			// 
			// txtNotas
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance2;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(704, 8);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(344, 88);
			this.txtNotas.TabIndex = 249;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 72);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 244;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// lblTarjeta
			// 
			this.lblTarjeta.AutoSize = true;
			this.lblTarjeta.BackColor = System.Drawing.Color.Transparent;
			this.lblTarjeta.Location = new System.Drawing.Point(320, 11);
			this.lblTarjeta.Name = "lblTarjeta";
			this.lblTarjeta.Size = new System.Drawing.Size(40, 16);
			this.lblTarjeta.TabIndex = 258;
			this.lblTarjeta.Text = "Tarjeta";
			this.lblTarjeta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Location = new System.Drawing.Point(8, 11);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(34, 16);
			this.lblBanco.TabIndex = 259;
			this.lblBanco.Text = "Banco";
			this.lblBanco.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTarjeta
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTarjeta.Appearance = appearance4;
			this.cmbTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTarjeta.DataSource = this.ultraDataSource5;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 224;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTarjeta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTarjeta.DisplayMember = "Nombre";
			this.cmbTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTarjeta.Enabled = false;
			this.cmbTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTarjeta.Location = new System.Drawing.Point(376, 8);
			this.cmbTarjeta.Name = "cmbTarjeta";
			this.cmbTarjeta.Size = new System.Drawing.Size(224, 21);
			this.cmbTarjeta.TabIndex = 256;
			this.cmbTarjeta.ValueMember = "idTarjeta";
			this.cmbTarjeta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTarjeta_KeyDown);
			this.cmbTarjeta.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTarjeta_Validating);
			this.cmbTarjeta.ValueChanged += new System.EventHandler(this.cmbTarjeta_ValueChanged);
			this.cmbTarjeta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTarjeta_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// cmbBanco
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBanco.Appearance = appearance5;
			this.cmbBanco.CausesValidation = false;
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBanco.DataSource = this.ultraDataSource4;
			ultraGridBand2.AddButtonCaption = "Compra";
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 224;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.Enabled = false;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(72, 8);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(224, 21);
			this.cmbBanco.TabIndex = 255;
			this.cmbBanco.ValueMember = "idBanco";
			this.cmbBanco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBanco_KeyDown);
			this.cmbBanco.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBanco_Validating);
			this.cmbBanco.ValueChanged += new System.EventHandler(this.cmbBanco_ValueChanged);
			this.cmbBanco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBanco_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// txtTotal
			// 
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance6;
			this.txtTotal.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtTotal.ForeColor = System.Drawing.Color.Black;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(1168, 40);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(120, 21);
			this.txtTotal.TabIndex = 251;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.lblSubTotal.ForeColor = System.Drawing.Color.Black;
			this.lblSubTotal.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblSubTotal.Location = new System.Drawing.Point(1080, 43);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(29, 16);
			this.lblSubTotal.TabIndex = 250;
			this.lblSubTotal.Text = "Total";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Appearance = appearance7;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.Caption = "S/F";
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 20;
			ultraGridColumn8.Header.Caption = "P";
			ultraGridColumn8.Header.VisiblePosition = 3;
			ultraGridColumn8.Width = 20;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 7;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 250;
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 65;
			ultraGridColumn15.Header.Caption = "Centro De Costo";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn15.Width = 160;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 50;
			ultraGridColumn17.Header.Caption = "Autorización";
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Width = 212;
			ultraGridColumn18.Header.Caption = "Número";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Width = 65;
			ultraGridColumn19.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn19.Header.VisiblePosition = 14;
			ultraGridColumn19.Width = 315;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance8;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance9;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 20;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance10;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 22;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance11;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 24;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance12;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance13;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 28;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 29;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance14;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 30;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance15;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 31;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 32;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 33;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance16;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance17;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 35;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 36;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn41.Header.VisiblePosition = 37;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance18;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 38;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance19;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 39;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 40;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 41;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance20;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 42;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance21;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.VisiblePosition = 43;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn48.Header.VisiblePosition = 44;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn49.Header.VisiblePosition = 45;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance22;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 46;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance23;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 47;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.VisiblePosition = 48;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn53.Header.VisiblePosition = 49;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance24;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 50;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance25;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 51;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn56.Header.VisiblePosition = 52;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn57.Header.VisiblePosition = 53;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance26;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 54;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance27;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 55;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.VisiblePosition = 56;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn61.Header.VisiblePosition = 57;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance28;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 58;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance29;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 59;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn64.Header.VisiblePosition = 60;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn65.Header.VisiblePosition = 61;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance30;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 62;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance31;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.VisiblePosition = 63;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn68.Header.VisiblePosition = 64;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn69.Header.VisiblePosition = 65;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance32;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 66;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance33;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 67;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn72.Header.VisiblePosition = 68;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn73.Header.VisiblePosition = 69;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance34;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 70;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance35;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 71;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn76.Header.VisiblePosition = 72;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn77.Header.VisiblePosition = 73;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance36;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 74;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance37;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 75;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance38;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "Subtotal IVA 0";
			ultraGridColumn80.Header.VisiblePosition = 77;
			ultraGridColumn80.Width = 85;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance39;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.Caption = "Subtotal IVA";
			ultraGridColumn81.Header.VisiblePosition = 76;
			ultraGridColumn81.Width = 85;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance40;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.VisiblePosition = 78;
			ultraGridColumn82.Width = 85;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance41;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.VisiblePosition = 79;
			ultraGridColumn83.Width = 85;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance42;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.VisiblePosition = 80;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance43;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.Caption = "Voucher";
			ultraGridColumn85.Header.VisiblePosition = 81;
			ultraGridColumn85.Width = 85;
			ultraGridColumn86.Header.Caption = "Dif";
			ultraGridColumn86.Header.VisiblePosition = 82;
			ultraGridColumn86.Width = 22;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance44;
			ultraGridColumn87.Header.VisiblePosition = 83;
			ultraGridColumn87.Width = 45;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance45;
			ultraGridColumn88.Format = "#,##0.00";
			ultraGridColumn88.Header.Caption = "Valor Cuotas";
			ultraGridColumn88.Header.VisiblePosition = 84;
			ultraGridColumn88.Width = 85;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance46;
			ultraGridColumn89.Format = "#,##0.00";
			ultraGridColumn89.Header.VisiblePosition = 85;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 85;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn90.Header.VisiblePosition = 86;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance47.Image = ((object)(resources.GetObject("appearance47.Image")));
			appearance47.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance47.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn91.CellButtonAppearance = appearance47;
			ultraGridColumn91.Header.Caption = "...";
			ultraGridColumn91.Header.VisiblePosition = 15;
			ultraGridColumn91.PromptChar = ' ';
			ultraGridColumn91.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn91.Width = 35;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91});
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance48;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance49;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance50;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance51;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance52;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance53;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance54;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.Hidden = true;
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance55.ForeColor = System.Drawing.Color.Black;
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance55;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance56.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance56;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.FontData.BoldAsString = "True";
			appearance57.FontData.Name = "Arial";
			appearance57.FontData.SizeInPoints = 8F;
			appearance57.ForeColor = System.Drawing.Color.White;
			appearance57.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance57;
			appearance58.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance58.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance59;
			appearance60.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance60.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance60;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F);
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1304, 240);
			this.ultraGrid1.TabIndex = 260;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(bool);
			ultraDataColumn8.DataType = typeof(bool);
			ultraDataColumn8.DefaultValue = false;
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Double);
			ultraDataColumn23.DataType = typeof(System.Double);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.Double);
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(System.Double);
			ultraDataColumn30.DefaultValue = 0;
			ultraDataColumn31.DataType = typeof(System.Double);
			ultraDataColumn31.DefaultValue = 0;
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(System.Double);
			ultraDataColumn34.DefaultValue = 0;
			ultraDataColumn35.DataType = typeof(System.Double);
			ultraDataColumn35.DefaultValue = 0;
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(System.Double);
			ultraDataColumn38.DefaultValue = 0;
			ultraDataColumn39.DataType = typeof(System.Double);
			ultraDataColumn39.DefaultValue = 0;
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(System.Double);
			ultraDataColumn42.DefaultValue = 0;
			ultraDataColumn43.DataType = typeof(System.Double);
			ultraDataColumn43.DefaultValue = 0;
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(System.Double);
			ultraDataColumn46.DefaultValue = 0;
			ultraDataColumn47.DataType = typeof(System.Double);
			ultraDataColumn47.DefaultValue = 0;
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(System.Double);
			ultraDataColumn50.DefaultValue = 0;
			ultraDataColumn51.DataType = typeof(System.Double);
			ultraDataColumn51.DefaultValue = 0;
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(System.Double);
			ultraDataColumn54.DefaultValue = 0;
			ultraDataColumn55.DataType = typeof(System.Double);
			ultraDataColumn55.DefaultValue = 0;
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(int);
			ultraDataColumn58.DataType = typeof(System.Double);
			ultraDataColumn58.DefaultValue = 0;
			ultraDataColumn59.DataType = typeof(System.Double);
			ultraDataColumn59.DefaultValue = 0;
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(int);
			ultraDataColumn62.DataType = typeof(System.Double);
			ultraDataColumn62.DefaultValue = 0;
			ultraDataColumn63.DataType = typeof(System.Double);
			ultraDataColumn63.DefaultValue = 0;
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(System.Double);
			ultraDataColumn66.DefaultValue = 0;
			ultraDataColumn67.DataType = typeof(System.Double);
			ultraDataColumn67.DefaultValue = 0;
			ultraDataColumn68.DataType = typeof(int);
			ultraDataColumn69.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(System.Double);
			ultraDataColumn70.DefaultValue = 0;
			ultraDataColumn71.DataType = typeof(System.Double);
			ultraDataColumn71.DefaultValue = 0;
			ultraDataColumn72.DataType = typeof(int);
			ultraDataColumn73.DataType = typeof(int);
			ultraDataColumn74.DataType = typeof(System.Double);
			ultraDataColumn74.DefaultValue = 0;
			ultraDataColumn75.DataType = typeof(System.Double);
			ultraDataColumn75.DefaultValue = 0;
			ultraDataColumn76.DataType = typeof(int);
			ultraDataColumn77.DataType = typeof(int);
			ultraDataColumn78.DataType = typeof(System.Double);
			ultraDataColumn78.DefaultValue = 0;
			ultraDataColumn79.DataType = typeof(System.Double);
			ultraDataColumn79.DefaultValue = 0;
			ultraDataColumn80.DataType = typeof(System.Double);
			ultraDataColumn80.DefaultValue = 0;
			ultraDataColumn81.DataType = typeof(System.Double);
			ultraDataColumn81.DefaultValue = 0;
			ultraDataColumn82.DataType = typeof(System.Double);
			ultraDataColumn82.DefaultValue = 0;
			ultraDataColumn83.DataType = typeof(System.Double);
			ultraDataColumn84.DataType = typeof(System.Double);
			ultraDataColumn84.DefaultValue = 0;
			ultraDataColumn85.DataType = typeof(System.Double);
			ultraDataColumn85.DefaultValue = 0;
			ultraDataColumn86.DataType = typeof(bool);
			ultraDataColumn87.DataType = typeof(int);
			ultraDataColumn88.DataType = typeof(System.Double);
			ultraDataColumn89.DataType = typeof(System.Double);
			ultraDataColumn89.DefaultValue = 0;
			ultraDataColumn90.DataType = typeof(int);
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
																																 ultraDataColumn49,
																																 ultraDataColumn50,
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
																																 ultraDataColumn67,
																																 ultraDataColumn68,
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71,
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74,
																																 ultraDataColumn75,
																																 ultraDataColumn76,
																																 ultraDataColumn77,
																																 ultraDataColumn78,
																																 ultraDataColumn79,
																																 ultraDataColumn80,
																																 ultraDataColumn81,
																																 ultraDataColumn82,
																																 ultraDataColumn83,
																																 ultraDataColumn84,
																																 ultraDataColumn85,
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88,
																																 ultraDataColumn89,
																																 ultraDataColumn90});
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(1128, 368);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 357;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(8, 352);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1304, 8);
			this.gbBotones.TabIndex = 356;
			this.gbBotones.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 368);
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
			this.btnConsultar.Location = new System.Drawing.Point(93, 368);
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
			this.btnEditar.Location = new System.Drawing.Point(176, 368);
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
			this.btnImprimir.Location = new System.Drawing.Point(257, 368);
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
			this.btnGuardar.Location = new System.Drawing.Point(339, 368);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 346;
			this.btnGuardar.Text = "&Guardar";
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
			this.btnCancelar.Location = new System.Drawing.Point(671, 368);
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
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(748, 368);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 349;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// txtIdGastos
			// 
			this.txtIdGastos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGastos.Enabled = false;
			this.txtIdGastos.Location = new System.Drawing.Point(1232, 368);
			this.txtIdGastos.Name = "txtIdGastos";
			this.txtIdGastos.PromptChar = ' ';
			this.txtIdGastos.Size = new System.Drawing.Size(16, 22);
			this.txtIdGastos.TabIndex = 358;
			this.txtIdGastos.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(1256, 368);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 359;
			this.txtEstado.Visible = false;
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
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnProcesar.Enabled = false;
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.ImageIndex = 69;
			this.btnProcesar.ImageList = this.imageList1;
			this.btnProcesar.Location = new System.Drawing.Point(420, 368);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(76, 23);
			this.btnProcesar.TabIndex = 360;
			this.btnProcesar.Text = "&Procesar";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnActualizar.Location = new System.Drawing.Point(1104, 368);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(112, 23);
			this.btnActualizar.TabIndex = 363;
			this.btnActualizar.Text = "&Actualizar Cuentas";
			this.btnActualizar.Visible = false;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// btnCalcular
			// 
			this.btnCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCalcular.CausesValidation = false;
			this.btnCalcular.Enabled = false;
			this.btnCalcular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCalcular.ImageIndex = 70;
			this.btnCalcular.ImageList = this.imageList1;
			this.btnCalcular.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnCalcular.Location = new System.Drawing.Point(586, 368);
			this.btnCalcular.Name = "btnCalcular";
			this.btnCalcular.Size = new System.Drawing.Size(80, 23);
			this.btnCalcular.TabIndex = 364;
			this.btnCalcular.Text = "&Calcular";
			this.btnCalcular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCalcular.Click += new System.EventHandler(this.btnCalcular_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.ImageIndex = 68;
			this.btnExportar.ImageList = this.imageList1;
			this.btnExportar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnExportar.Location = new System.Drawing.Point(501, 368);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 365;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn91.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn91,
																																 ultraDataColumn92});
			// 
			// cmbCentroDeCosto
			// 
			this.cmbCentroDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroDeCosto.DataSource = this.ultraDataSource6;
			this.cmbCentroDeCosto.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn92.Header.VisiblePosition = 0;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 177;
			ultraGridColumn93.Header.VisiblePosition = 1;
			ultraGridColumn93.Width = 413;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn92,
																										 ultraGridColumn93});
			ultraGridBand4.SummaryFooterCaption = "Grand Summaries";
			this.cmbCentroDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCentroDeCosto.DisplayMember = "Nombre";
			this.cmbCentroDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCentroDeCosto.Location = new System.Drawing.Point(32, 200);
			this.cmbCentroDeCosto.Name = "cmbCentroDeCosto";
			this.cmbCentroDeCosto.Size = new System.Drawing.Size(432, 88);
			this.cmbCentroDeCosto.TabIndex = 366;
			this.cmbCentroDeCosto.ValueMember = "idProyecto";
			this.cmbCentroDeCosto.Visible = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 368;
			this.label2.Text = "Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNombre
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbNombre.Appearance = appearance61;
			this.cmbNombre.CausesValidation = false;
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataSource = this.ultraDataSource7;
			ultraGridBand5.AddButtonCaption = "Compra";
			ultraGridColumn94.Header.VisiblePosition = 0;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 1;
			ultraGridColumn95.Width = 408;
			ultraGridColumn96.Header.VisiblePosition = 2;
			ultraGridColumn96.Width = 200;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.Enabled = false;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(72, 40);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(304, 21);
			this.cmbNombre.TabIndex = 367;
			this.cmbNombre.ValueMember = "idTarjetasCorporativas";
			this.cmbNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNombre_KeyDown);
			this.cmbNombre.Validating += new System.ComponentModel.CancelEventHandler(this.cmbNombre_Validating);
			this.cmbNombre.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNombre_InitializeLayout);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn93.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.label4.ForeColor = System.Drawing.Color.Black;
			this.label4.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.label4.Location = new System.Drawing.Point(1080, 75);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(75, 16);
			this.label4.TabIndex = 369;
			this.label4.Text = "Total Voucher";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalVoucher
			// 
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			appearance62.ForeColor = System.Drawing.Color.Black;
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalVoucher.Appearance = appearance62;
			this.txtTotalVoucher.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.txtTotalVoucher.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalVoucher.Enabled = false;
			this.txtTotalVoucher.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtTotalVoucher.ForeColor = System.Drawing.Color.Black;
			this.txtTotalVoucher.FormatString = "#,##0.00";
			this.txtTotalVoucher.Location = new System.Drawing.Point(1168, 72);
			this.txtTotalVoucher.Name = "txtTotalVoucher";
			this.txtTotalVoucher.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalVoucher.PromptChar = ' ';
			this.txtTotalVoucher.Size = new System.Drawing.Size(120, 21);
			this.txtTotalVoucher.TabIndex = 370;
			// 
			// frmGastosTarjetasDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1320, 398);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtTotalVoucher);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.cmbCentroDeCosto);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnCalcular);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdGastos);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.lblTarjeta);
			this.Controls.Add(this.lblBanco);
			this.Controls.Add(this.cmbTarjeta);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.dtFecha);
			this.Name = "frmGastosTarjetasDeCredito";
			this.Text = "Gastos de Tarjetas De Crédito";
			this.Load += new System.EventHandler(this.frmGastosTarjetasDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalVoucher)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		bool bEdicion = false;
		bool bNuevo = false;
		bool bEliminaAlValidar = false;
		CultureInfo us = new CultureInfo("en-US");
		#endregion Variables

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void frmGastosTarjetasDeCredito_Load(object sender, System.EventArgs e)
		{
			ModificaGrilla(false);

			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Where ISNULL(Pais, '') = 'Ecuador' Order By Nombre");
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.txtIdGastos.Value = 0;
			this.txtEstado.Value = 0;
			this.cmbBanco.Value = System.DBNull.Value;
			this.cmbTarjeta.Value = System.DBNull.Value;
			this.cmbNombre.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtNotas.Text = "";
			this.txtNumero.Text = "";
			this.txtTotal.Value = 0;
			this.txtTotalVoucher.Value = 0;
			this.lblEstado.Text = "";

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec GTCConsultaDetalle 0");

			this.cmbBanco.Enabled = false;
			this.cmbTarjeta.Enabled = false;
			this.cmbNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNotas.Enabled = false;

			ModificaGrilla(false);

      this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnProcesar.Enabled = false;
			this.btnExportar.Enabled = false;
			this.btnCalcular.Enabled = false;
			this.btnCancelar.Enabled = false;

			bEdicion = false;
			bNuevo = false;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbBanco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbBanco_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBanco.ActiveRow == null) return;

			string sSQL = string.Format("Exec BTPP {0}, 0, 0, 0, 1", (int)this.cmbBanco.Value);
			this.cmbTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.dtFecha.Value = DateTime.Today;
			this.txtEstado.Value = 0;

			this.cmbBanco.Enabled = true;
			this.cmbTarjeta.Enabled = true;
			this.cmbNombre.Enabled = true;
			this.dtFecha.Enabled = true;
			this.txtNotas.Enabled = true;

			ModificaGrilla(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnProcesar.Enabled = false;
			this.btnExportar.Enabled = false;
			this.btnCalcular.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.cmbBanco.Focus();
		}

		private void cmbBanco_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbBanco.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbBanco.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void cmbTarjeta_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTarjeta.ActiveRow == null)
				{
					MessageBox.Show("Seleccione una Tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbTarjeta.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbBanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbTarjeta.Focus();
		}

		private void cmbTarjeta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNombre.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNotas.Focus();
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Ruc"];
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaGastosTarjetas Busqueda = new frmBuscaGastosTarjetas())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					try
					{						
						this.txtIdGastos.Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idGastosTarjetas"].Value;
						this.cmbBanco.Value  = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idBanco"].Value;
						this.cmbTarjeta.Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idTarjeta"].Value;
						this.cmbNombre.Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idTarjetasCorporativas"].Value;
						this.dtFecha.Value = (DateTime)Busqueda.ultraGrid1.ActiveRow.Cells["Fecha"].Value;
						this.txtNumero.Text = Busqueda.ultraGrid1.ActiveRow.Cells["Numero"].Value.ToString();
						this.txtNotas.Text = Busqueda.ultraGrid1.ActiveRow.Cells["Notas"].Value.ToString();
						this.txtTotal.Value  = (double)Busqueda.ultraGrid1.ActiveRow.Cells["Total"].Value;
						this.txtTotalVoucher.Value  = (double)Busqueda.ultraGrid1.ActiveRow.Cells["TotalVoucher"].Value;
						this.txtEstado.Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["Estado"].Value;						
						
						if ((int)this.txtEstado.Value == 0) this.lblEstado.Text = "EN ESPERA";
						if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PROCESADO";

            this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec GTCConsultaDetalle {0}", (int)this.txtIdGastos.Value));

						this.btnEditar.Enabled = true;
						this.btnImprimir.Enabled = true;						
						this.btnProcesar.Enabled = true;
						this.btnExportar.Enabled = true;						
						this.btnCancelar.Enabled = true;
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}	
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.cmbBanco.Enabled = true;
			this.cmbTarjeta.Enabled = true;
			this.cmbNombre.Enabled = true;
			this.dtFecha.Enabled = true;
			this.txtNotas.Enabled = true;

			ModificaGrilla(true);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnProcesar.Enabled = false;
			this.btnExportar.Enabled = false;
			this.btnCalcular.Enabled = true;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = true;

			this.cmbBanco.Focus();			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion Encabezado
			if (this.cmbBanco.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBanco.Focus();				
				return;
			}

			if (this.cmbTarjeta.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbTarjeta.Focus();				
				return;
			}

			if (this.cmbNombre.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Usuario de la Tarjeta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbNombre.Focus();				
				return;
			}

			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();		
				return;
			}
			
			#endregion Validacion Detalle

			#region Validacion Detalle
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los documentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}	
		
			#region Elimina filas vacias
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEl in ultraGrid1.Rows.All)
			{			
				if ((double)drEl.Cells["ValorToTal"].Value == 0 && (double)drEl.Cells["Total"].Value == 0)
				{
					bEliminaAlValidar = true;
					drEl.Delete();					
				}
			}
			#endregion Elimina filas vacias

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{				
				#region Ingreso Con Factura
				if (!(bool)dr.Cells["SinFactura"].Value)
				{
					#region Valida Ingreso de Ruc
					if (dr.Cells["Ruc"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el Ruc del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de Ruc

					#region Valida identificación
					int idTipoRUC = 1;
					if (dr.Cells["Ruc"].Value.ToString().Length == 10) idTipoRUC = 2;
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["Ruc"].Value.ToString(), idTipoRUC);
					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["Ruc"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["Ruc"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor

					#region Autorización
					if (dr.Cells["Autorizacion"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Autorizacion"].Value.ToString().Length != 10)
					{
						if (dr.Cells["Autorizacion"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización es de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Autorización

					#region Serie
					if (dr.Cells["Serie"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese la Serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Serie

					#region Numero
					if (dr.Cells["Numero"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el número de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}

					char Pad = '0';
					dr.Cells["Numero"].Value = dr.Cells["Numero"].Value.ToString().PadLeft(9, Pad);
					
					int idDCCValN = 0;
					int idCompraValN = 0;

					if (dr.Cells["idGastosTarjetas"].Value != System.DBNull.Value) idDCCValN = (int)dr.Cells["idGastosTarjetas"].Value; 
					if (dr.Cells["idCompra"].Value != System.DBNull.Value) idCompraValN = (int)dr.Cells["idCompra"].Value; 

					string sSQLValNum = string.Format("Exec ValidaNumeracionCompraCajaChica {0}, {1}, {2}, '{3}', '{4}', '{5}'", 
						idDCCValN, idCompraValN, (int)dr.Cells["idProveedor"].Value, dr.Cells["Serie"].Value.ToString(), dr.Cells["Numero"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString());
					string sMensValNum = Funcion.sEscalarSQL(cdsSeteoF, sSQLValNum, true);
					if (sMensValNum.Length > 0)
					{
						MessageBox.Show(sMensValNum, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Numero
				}
				#endregion Ingreso Con Factura

				if (!(bool)dr.Cells["SinFactura"].Value)
				{
					#region Valida Ingreso de Ruc
					if (dr.Cells["Ruc"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese el Ruc del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida Ingreso de Ruc

					#region Valida identificación
					int idTipoRUC = 1;
					if (dr.Cells["Ruc"].Value.ToString().Length == 10) idTipoRUC = 2;
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["Ruc"].Value.ToString(), idTipoRUC);
					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.ultraGrid1.ActiveRow = dr;
						this.ultraGrid1.ActiveRow.Selected = true;
						return;
					}
					#endregion Valida identificación

					#region Busca Proveedor
					if (dr.Cells["Ruc"].Value.ToString().Length > 0)
					{
						string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["Ruc"].Value.ToString(), 4, 0, false);
						SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
						drC.Read();
						if (!drC.HasRows)
						{
							drC.Close();
							MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.ultraGrid1.ActiveRow = dr;
							this.ultraGrid1.ActiveRow.Selected = true;
							return;
						}
						drC.Close();
					}
					#endregion Busca Proveedor
				}

				#region Fecha
				if (dr.Cells["Fecha"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				#endregion Fecha

				#region Centro de Costo
				if (dr.Cells["idProyecto"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				#endregion Centro de Costo

				#region Detalle
				if (dr.Cells["Detalle"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Detalle de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Detalle

				#region Cuotas
				if ((bool)dr.Cells["Diferido"].Value && (int)dr.Cells["Cuotas"].Value == 1)	
				{
					MessageBox.Show("En Pago Diferido no puede ser una Cuota", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}

				if ((bool)dr.Cells["Diferido"].Value && (double)dr.Cells["ValorCuotas"].Value == 0)	
				{
					MessageBox.Show("Ingrese el Valor de las Cuotas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
				#endregion Cuotas

				#region Articulo
				if (dr.Cells["idArticulo1"].Value == System.DBNull.Value && dr.Cells["idArticulo2"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Articulo

				#region Valor del Voucher
				if ((double)dr.Cells["ValorTotal"].Value == 0)
				{
					MessageBox.Show("Ingrese el Valor del Voucher", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}					
				#endregion Valor del Voucher			
			}
			#endregion Validación Detalle

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
					if (bNuevo)
					{
						string sSQLNumero = "Exec GTCNumeracion";
						oCmdActualiza.CommandText = sSQLNumero;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
					}
					#endregion Numeracion

					#region Maestro
					DateTime dtFecha = (DateTime)this.dtFecha.Value;

					string sSQLMaestro = string.Format("Exec GTCGuardaMaestro {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', {7}, {8}, {9}",
						(int)this.txtIdGastos.Value, (int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, (int)this.cmbNombre.Value, dtFecha.ToString("yyyyMMdd"), this.txtNumero.Text.ToString(), this.txtNotas.Text.ToString(), (double)this.txtTotal.Value, (double)this.txtTotalVoucher.Value, (int)this.txtEstado.Value);
					oCmdActualiza.CommandText = sSQLMaestro;
					this.txtIdGastos.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Maestro

					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{	
						DateTime dtFechaDetalle = (DateTime)dr.Cells["Fecha"].Value;
//						int idArticulo1 = 0;
//						if (dr.Cells["idArticulo1"].Value != System.DBNull.Value) idArticulo1 = (int)dr.Cells["idArticulo1"].Value;
//						int idArticulo2 = 0;
//						if (dr.Cells["idArticulo2"].Value != System.DBNull.Value) idArticulo2 = (int)dr.Cells["idArticulo2"].Value;
						int idProveedor = 0;
						if (dr.Cells["idProveedor"].Value != System.DBNull.Value) idProveedor = (int)dr.Cells["idProveedor"].Value;
						
						string sSQLDetalle = 
							string.Format("Exec GTCGuardaDetalle {0}, {1}, {2}, {3}, '{4}', {5}, '{6}', '{7}', '{8}', '{9}', {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}",
							(int)dr.Cells["idDetalleGastosTarjetas"].Value, (int)this.txtIdGastos.Value, (bool)dr.Cells["SinFactura"].Value,
							idProveedor, dtFechaDetalle.ToString("yyyyMMdd"), (int)dr.Cells["idProyecto"].Value, 
							dr.Cells["Serie"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
							dr.Cells["Detalle"].Value.ToString(),							
							(int)dr.Cells["idArticulo1"].Value, (double)dr.Cells["IVA1"].Value, (double)dr.Cells["Subtotal1"].Value,
							(int)dr.Cells["idArticulo2"].Value, (double)dr.Cells["IVA2"].Value, (double)dr.Cells["Subtotal2"].Value, 
							(int)dr.Cells["idArticulo3"].Value, (double)dr.Cells["IVA3"].Value, (double)dr.Cells["Subtotal3"].Value,
							(int)dr.Cells["idArticulo4"].Value, (double)dr.Cells["IVA4"].Value, (double)dr.Cells["Subtotal4"].Value,
							(int)dr.Cells["idArticulo5"].Value, (double)dr.Cells["IVA5"].Value, (double)dr.Cells["Subtotal5"].Value,
							(int)dr.Cells["idArticulo6"].Value, (double)dr.Cells["IVA6"].Value, (double)dr.Cells["Subtotal6"].Value,
							(int)dr.Cells["idArticulo7"].Value, (double)dr.Cells["IVA7"].Value, (double)dr.Cells["Subtotal7"].Value,
							(int)dr.Cells["idArticulo8"].Value, (double)dr.Cells["IVA8"].Value, (double)dr.Cells["Subtotal8"].Value,
							(int)dr.Cells["idArticulo9"].Value, (double)dr.Cells["IVA9"].Value, (double)dr.Cells["Subtotal9"].Value,
							(int)dr.Cells["idArticulo10"].Value, (double)dr.Cells["IVA10"].Value, (double)dr.Cells["Subtotal10"].Value,
							(int)dr.Cells["idArticulo11"].Value, (double)dr.Cells["IVA11"].Value, (double)dr.Cells["Subtotal11"].Value,
							(int)dr.Cells["idArticulo12"].Value, (double)dr.Cells["IVA12"].Value, (double)dr.Cells["Subtotal12"].Value,
							(int)dr.Cells["idArticulo13"].Value, (double)dr.Cells["IVA13"].Value, (double)dr.Cells["Subtotal13"].Value,
							(int)dr.Cells["idArticulo14"].Value, (double)dr.Cells["IVA14"].Value, (double)dr.Cells["Subtotal14"].Value,
							(int)dr.Cells["idArticulo15"].Value, (double)dr.Cells["IVA15"].Value, (double)dr.Cells["Subtotal15"].Value,
							(double)dr.Cells["SubtotalIVA"].Value, (double)dr.Cells["SubtotalIVA0"].Value, (double)dr.Cells["IVA"].Value, 
							(double)dr.Cells["Total"].Value, (double)dr.Cells["Interes"].Value, (double)dr.Cells["ValorTotal"].Value,
							(bool)dr.Cells["Diferido"].Value, (int)dr.Cells["Cuotas"].Value, (double)dr.Cells["ValorCuotas"].Value, 
							(double)dr.Cells["Saldo"].Value, (int)dr.Cells["Estado"].Value, (bool)dr.Cells["Point"].Value);
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
						
						#region Orden de Parametros
						/*@idDetalleGastosTarjetas Int, @idGastosTarjetas Int, @SinFactura Bit, 0 1 2
						@idProveedor Int, @Fecha Datetime, @idProyecto Int, 3 4 5
						@Serie Varchar(10), @Autorizacion Varchar(40), @Numero Varchar(20), 6 7 8
						@Detalle Varchar(600),	9
						@idArticulo1 Int, @IVA1 Float, @Subtotal1 Float, 10 11 12
						@idArticulo2 Int, @IVA2 Float, @Subtotal2 Float, 13 14 15
						@idArticulo3 Int, @IVA3 Float, @Subtotal3 Float, 16 17 18
						@idArticulo4 Int, @IVA4 Float, @Subtotal4 Float, 19 20 21
						@idArticulo5 Int, @IVA5 Float, @Subtotal5 Float, 22 23 24
						@idArticulo6 Int, @IVA6 Float, @Subtotal6 Float, 25 26 27
						@idArticulo7 Int, @IVA7 Float, @Subtotal7 Float, 28 29 30
						@idArticulo8 Int, @IVA8 Float, @Subtotal8 Float, 31 32 33
						@idArticulo9 Int, @IVA9 Float, @Subtotal9 Float, 34 35 36
						@idArticulo10 Int, @IVA10 Float, @Subtotal10 Float, 37 38 39
						@idArticulo11 Int, @IVA11 Float, @Subtotal11 Float, 40 41 42
						@idArticulo12 Int, @IVA12 Float, @Subtotal12 Float, 43 44 45
						@idArticulo13 Int, @IVA13 Float, @Subtotal13 Float, 46 47 48
						@idArticulo14 Int, @IVA14 Float, @Subtotal14 Float, 49 50 51
						@idArticulo15 Int, @IVA15 Float, @Subtotal15 Float,	52 53 54 
						@SubtotalIVA Float, @SubtotalIVA0 Float, @IVA Float, @Total Float, @ValorTotal Float, 55 56 57 58 59
						@Diferido Bit, @Cuotas Int, @ValorCuotas Float, 60 61 62
						@Saldo Float, @Estado Int, @Point Bit 63 64*/
						#endregion Orden de Parametros
					}

					string sSQLActDetalle = string.Format("Exec GTCConsultaDetalle {0}", (int)this.txtIdGastos.Value);
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.ultraGrid1.DataSource = oDTDetalle;
					#endregion Detalle

					oTransaction.Commit();

					#region Controles
					bNuevo = false;
					bEdicion = false;

					this.cmbBanco.Enabled = false;
					this.cmbTarjeta.Enabled = false;
					this.cmbNombre.Enabled = false;
					this.dtFecha.Enabled = false;
					this.txtNotas.Enabled = false;
								
					ModificaGrilla(false);
														
					this.btnNuevo.Enabled = true;
					this.btnConsultar.Enabled = true;
					this.btnEditar.Enabled = true;
					this.btnImprimir.Enabled = true;
					this.btnGuardar.Enabled = false;
					this.btnProcesar.Enabled = true;
					this.btnExportar.Enabled = true;
					this.btnCalcular.Enabled = false;
					this.btnCancelar.Enabled = true;
					#endregion Controles						

					//this.btnProcesar_Click(sender, e);
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

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("¿Desea Procesar el Registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
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

						DateTime dtFecha = DateTime.Today;						
						double dBase0 = 0;
						double dBaseImp = 0;
						double dIVATotal = 0;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{	
							#region Factura
							dBase0 = 0;
							dBaseImp = 0;
							dIVATotal = 0;
							
							if (!(bool)dr.Cells["SinFactura"].Value)
							{						
								dtFecha = (DateTime)dr.Cells["Fecha"].Value;

//								int idArticulo1 = 0;
//								if (dr.Cells["idArticulo1"].Value != System.DBNull.Value) idArticulo1 = (int)dr.Cells["idArticulo1"].Value;
//								int idArticulo2 = 0;
//								if (dr.Cells["idArticulo2"].Value != System.DBNull.Value) idArticulo2 = (int)dr.Cells["idArticulo2"].Value;
								int idProveedor = 0;
								if (dr.Cells["idProveedor"].Value != System.DBNull.Value) idProveedor = (int)dr.Cells["idProveedor"].Value;
								int idCompra = 0;
								if (dr.Cells["idCompra"].Value != System.DBNull.Value)idCompra = (int)dr.Cells["idCompra"].Value;

								if ((double)dr.Cells["IVA1"].Value > 0)
								{
									dIVATotal = dIVATotal + (((double)dr.Cells["Subtotal1"].Value * 12)/100);
									dBaseImp = dBaseImp + (double)dr.Cells["Subtotal1"].Value;
								}
								else if ((double)dr.Cells["IVA1"].Value == 0) dBase0 = dBase0 + (double)dr.Cells["Subtotal1"].Value;

								if ((double)dr.Cells["IVA2"].Value > 0)
								{
									dIVATotal = dIVATotal + (((double)dr.Cells["Subtotal2"].Value * 12)/100);
									dBaseImp = dBaseImp + (double)dr.Cells["Subtotal2"].Value;
								}
								else if ((double)dr.Cells["IVA2"].Value == 0) dBase0 = dBase0 + (double)dr.Cells["Subtotal2"].Value;
								
								DateTime dtFechaGen = (DateTime)dr.Cells["Fecha"].Value;
								DateTime dtFechaCorte = (DateTime)this.dtFecha.Value;

								string sSQLGeneraCompra = string.Format("Exec CreaCompraGastosTarjetas {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, '{21}', {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, {59}, {60}, {61}, {62}, {63}, {64}, {65}, {66}, {67}, {68}, {69}, {70}, {71}, {72}, {73}, {74}, {75}, {76}, {77}, {78}, {79}, {80}, {81}", 
									(int)dr.Cells["idDetalleGastosTarjetas"].Value, idCompra, 
									(int)dr.Cells["idProyecto"].Value, idProveedor, "", "", dtFechaGen.ToString("yyyyMMdd"),
									dr.Cells["Autorizacion"].Value.ToString(), dr.Cells["Serie"].Value.ToString(), dr.Cells["Numero"].Value.ToString(), dr.Cells["Detalle"].Value.ToString(),
									(int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value, 
									(bool)dr.Cells["Diferido"].Value, (int)dr.Cells["Cuotas"].Value, (double)dr.Cells["ValorCuotas"].Value,
									(double)dr.Cells["SubtotalIVA0"].Value, (double)dr.Cells["SubtotalIVA"].Value, (double)dr.Cells["IVA"].Value, (double)dr.Cells["Total"].Value, (double)dr.Cells["ValorTotal"].Value, 
									this.txtNumero.Text.ToString(),							/*bases*/
									(int)dr.Cells["idDetCompra1"].Value, (int)dr.Cells["idArticulo1"].Value, (double)dr.Cells["IVA1"].Value, (double)dr.Cells["Subtotal1"].Value,
									(int)dr.Cells["idDetCompra2"].Value, (int)dr.Cells["idArticulo2"].Value, (double)dr.Cells["IVA2"].Value, (double)dr.Cells["Subtotal2"].Value,
									(int)dr.Cells["idDetCompra3"].Value, (int)dr.Cells["idArticulo3"].Value, (double)dr.Cells["IVA3"].Value, (double)dr.Cells["Subtotal3"].Value,
									(int)dr.Cells["idDetCompra4"].Value, (int)dr.Cells["idArticulo4"].Value, (double)dr.Cells["IVA4"].Value, (double)dr.Cells["Subtotal4"].Value,
									(int)dr.Cells["idDetCompra5"].Value, (int)dr.Cells["idArticulo5"].Value, (double)dr.Cells["IVA5"].Value, (double)dr.Cells["Subtotal5"].Value,
									(int)dr.Cells["idDetCompra6"].Value, (int)dr.Cells["idArticulo6"].Value, (double)dr.Cells["IVA6"].Value, (double)dr.Cells["Subtotal6"].Value,
									(int)dr.Cells["idDetCompra7"].Value, (int)dr.Cells["idArticulo7"].Value, (double)dr.Cells["IVA7"].Value, (double)dr.Cells["Subtotal7"].Value,
									(int)dr.Cells["idDetCompra8"].Value, (int)dr.Cells["idArticulo8"].Value, (double)dr.Cells["IVA8"].Value, (double)dr.Cells["Subtotal8"].Value,
									(int)dr.Cells["idDetCompra9"].Value, (int)dr.Cells["idArticulo9"].Value, (double)dr.Cells["IVA9"].Value, (double)dr.Cells["Subtotal9"].Value,
									(int)dr.Cells["idDetCompra10"].Value, (int)dr.Cells["idArticulo10"].Value, (double)dr.Cells["IVA10"].Value, (double)dr.Cells["Subtotal10"].Value,
									(int)dr.Cells["idDetCompra11"].Value, (int)dr.Cells["idArticulo11"].Value, (double)dr.Cells["IVA11"].Value, (double)dr.Cells["Subtotal11"].Value,
									(int)dr.Cells["idDetCompra12"].Value, (int)dr.Cells["idArticulo12"].Value, (double)dr.Cells["IVA12"].Value, (double)dr.Cells["Subtotal12"].Value,
									(int)dr.Cells["idDetCompra13"].Value, (int)dr.Cells["idArticulo13"].Value, (double)dr.Cells["IVA13"].Value, (double)dr.Cells["Subtotal13"].Value,
									(int)dr.Cells["idDetCompra14"].Value, (int)dr.Cells["idArticulo14"].Value, (double)dr.Cells["IVA14"].Value, (double)dr.Cells["Subtotal14"].Value,
									(int)dr.Cells["idDetCompra15"].Value, (int)dr.Cells["idArticulo15"].Value, (double)dr.Cells["IVA15"].Value, (double)dr.Cells["Subtotal15"].Value); 
									
								oCmdActualiza.CommandText = sSQLGeneraCompra;
								oCmdActualiza.ExecuteNonQuery();	
							}
							#endregion Factura

							#region Gasto Socios
							if (!(bool)dr.Cells["Point"].Value)
							{
								string sSQLAs = string.Format("Exec GTCCreaAsiento {0}", (int)dr.Cells["idDetalleGastosTarjetas"].Value);
								oCmdActualiza.CommandText = sSQLAs;
								oCmdActualiza.ExecuteNonQuery();
							}
							#endregion Gasto Socios
						}

						string sSQLActDetalle = string.Format("Exec GTCConsultaDetalle {0}", (int)this.txtIdGastos.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);
						this.ultraGrid1.DataSource = oDTDetalle;

						string sSQLActEstado = string.Format("Update GastosTarjetas Set Estado = 1 Where idGastosTarjetas = {0}", (int)this.txtIdGastos.Value);
						oCmdActualiza.CommandText = sSQLActEstado;
						oCmdActualiza.ExecuteNonQuery();

						#region Cuotas
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
						{	
							DateTime dtFechaGen = (DateTime)dr1.Cells["Fecha"].Value;

							if (!(bool)dr1.Cells["SinFactura"].Value)
							{	
								string sSQLGTCCuotas = string.Format("Exec GTCGuardaCuotas {0}, {1}, '{2}', {3}, {4}, {5}", 
									0, (int)dr1.Cells["idCompra"].Value, dtFechaGen.ToString("yyyyMMdd"), (int)dr1.Cells["Cuotas"].Value, (double)dr1.Cells["ValorCuotas"].Value, (bool)dr1.Cells["SinFactura"].Value);
								oCmdActualiza.CommandText = sSQLGTCCuotas;
								oCmdActualiza.ExecuteNonQuery();
							}
							else
							{
								string sSQLGTCCuotas = string.Format("Exec GTCGuardaCuotas {0}, {1}, '{2}', {3}, {4}, {5}", 
									0, (int)dr1.Cells["idDetalleGastosTarjetas"].Value, dtFechaGen.ToString("yyyyMMdd"), (int)dr1.Cells["Cuotas"].Value, (double)dr1.Cells["ValorCuotas"].Value, (bool)dr1.Cells["SinFactura"].Value);
								oCmdActualiza.CommandText = sSQLGTCCuotas;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
						#endregion Cuotas

						this.txtEstado.Value = 1;

						this.lblEstado.Text = "PROCESADO";

						oTransaction.Commit();								

						MessageBox.Show("Proceso Finalizado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

						this.btnProcesar.Enabled = false;
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

		private void Totales()
		{							
			this.txtTotal.Value = 0;
			this.txtTotalVoucher.Value = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{	
				if (dr.Cells["ValorCuotas"].Value == System.DBNull.Value || dr.Cells["Interes"].Value == System.DBNull.Value || dr.Cells["ValorTotal"].Value == System.DBNull.Value || dr.Cells["Total"].Value == System.DBNull.Value || dr.Cells["Cuotas"].Value == System.DBNull.Value) return;

				dr.Cells["ValorTotal"].Value = (double)dr.Cells["Total"].Value + (double)dr.Cells["Interes"].Value;

				dr.Cells["ValorCuotas"].Value = Math.Round((double)dr.Cells["ValorTotal"].Value/(int)dr.Cells["Cuotas"].Value, 2);

				this.txtTotal.Value = (double)this.txtTotal.Value + (double)dr.Cells["Total"].Value;
				this.txtTotalVoucher.Value = (double)this.txtTotalVoucher.Value + (double)dr.Cells["ValorTotal"].Value;
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SinFactura")
			{
				if (e.Cell.Row.Cells["Point"].Value != System.DBNull.Value)
				{
					if (!(bool)e.Cell.Row.Cells["Point"].Value && !(bool)e.Cell.Row.Cells["SinFactura"].Value)
					{
						MessageBox.Show("Si el Consumo no es de la Empresa no puede Ingresar la Factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["SinFactura"].Value = true;						
					}
				}
			}

			if (e.Cell.Column.ToString() == "Ruc")
			{
				#region Valida identificación
				if (e.Cell.Row.Cells["Ruc"].Value.ToString().Length > 0)
				{
					int idTipoRUC = 1;
					if (e.Cell.Row.Cells["Ruc"].Value.ToString().Length == 10) idTipoRUC = 2;
					string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", e.Cell.Row.Cells["Ruc"].Value.ToString(), idTipoRUC);
					string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
					if (VlCed.StartsWith("Error"))
					{
						MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						e.Cell.Row.Cells["Ruc"].Value = "";
						return;
					}
				}
				#endregion Valida identificación

				#region Busca Proveedor
				if (e.Cell.Row.Cells["Ruc"].Value.ToString().Length > 0)
				{
					string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", e.Cell.Row.Cells["Ruc"].Value.ToString(), 4, 0, false);
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
					drC.Read();
					if (drC.HasRows)
					{
						e.Cell.Row.Cells["idProveedor"].Value = (int) drC.GetValue(0);
						e.Cell.Row.Cells["Nombre"].Value = drC.GetString(2);	
			
						drC.Close();

						//						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor {0}", e.Cell.Row.Cells["Ruc"].Value.ToString())) == 0)
						//						{
						//							MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						//							e.Cell.Row.Cells["idProveedor"].Value = System.DBNull.Value;
						//							e.Cell.Row.Cells["Ruc"].Value = "";
						//							e.Cell.Row.Cells["Nombre"].Value = "";
						//							return;
						//						}
					}
					else
					{
						drC.Close();
						MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["Ruc"].Value = "";
						return;
					}
					drC.Close();
				}
				#endregion Busca Proveedor
			}

			#region Autorización
			if (e.Cell.Column.ToString() == "Autorización")
			{
				if (e.Cell.Row.Cells["Autorización"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["Autorización"].Value.ToString().Length != 10)
					{
						if (e.Cell.Row.Cells["Autorización"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización debe ser de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							e.Cell.Row.Cells["Autorización"].Value = "";
							return;
						}
					}
				}
			}
			#endregion Autorización

			#region Serie
			if (e.Cell.Column.ToString() == "Serie")
			{
				if (e.Cell.Row.Cells["Serie"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["Serie"].Value = "";
						return;
					}
				}
			}
			#endregion Serie

			#region Cuotas
			if (e.Cell.Column.ToString() == "Cuotas")
			{
				if ((bool)e.Cell.Row.Cells["Diferido"].Value && (int)e.Cell.Row.Cells["Cuotas"].Value < 2)
				{
					MessageBox.Show("En Diferido el Número de Cuotas debe ser Mayor a UNO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);												
					e.Cell.Row.Cells["Cuotas"].Value = 3;
				}
//				else if ((bool)e.Cell.Row.Cells["Diferido"].Value && (int)e.Cell.Row.Cells["Cuotas"].Value != 1)
//				{
//					MessageBox.Show("En Corriente el Número de Cuotas debe ser UNO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);												
//					e.Cell.Row.Cells["Cuotas"].Value = 1;
//				}

				Totales();
			}
			#endregion Cuotas

			#region Valor Total
			if (e.Cell.Column.ToString() == "Total")
			{
				Totales();
			}

			if (e.Cell.Column.ToString() == "Interes")
			{				
				Totales();
			}

			if (e.Cell.Column.ToString() == "ValorTotal")
			{
				
				if ((double)e.Cell.Row.Cells["ValorTotal"].Value < (double)e.Cell.Row.Cells["Total"].Value)
				{
					MessageBox.Show(string.Format("El Valor del Voucher {0} no puede ser menor al Valor de la Factura {1}", Math.Round((double)e.Cell.Row.Cells["ValorTotal"].Value, 2), Math.Round((double)e.Cell.Row.Cells["Total"].Value, 2)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);												
					e.Cell.Row.Cells["ValorTotal"].Value = (double)e.Cell.Row.Cells["Total"].Value;
				}
				
				Totales();
			}
			#endregion Valor Total
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "SinFactura")
			{
				if (this.ultraGrid1.ActiveRow.Cells["Ruc"].Value.ToString().Length > 0 || this.ultraGrid1.ActiveRow.Cells["Nombre"].Value.ToString().Length > 0 
					|| this.ultraGrid1.ActiveRow.Cells["Autorizacion"].Value.ToString().Length > 0 || this.ultraGrid1.ActiveRow.Cells["Serie"].Value.ToString().Length > 0 
					|| this.ultraGrid1.ActiveRow.Cells["Numero"].Value.ToString().Length > 0)
				{
					MessageBox.Show("No puede crear un gasto sin Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel= true;
					return;
				}
			}
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
			e.Layout.Bands[0].Columns["ValorCuotas"].Format = stFormato; 
			e.Layout.Bands[0].Columns["ValorCuotas"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["IVA1"].Format = stFormato; 
			e.Layout.Bands[0].Columns["IVA1"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["IVA2"].Format = stFormato; 
			e.Layout.Bands[0].Columns["IVA2"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Subtotal1"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Subtotal1"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Subtotal2"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Subtotal2"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Total"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Total"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["Interes"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Interes"].MaskInput = stInput;
			e.Layout.Bands[0].Columns["ValorTotal"].Format = stFormato; 
			e.Layout.Bands[0].Columns["ValorTotal"].MaskInput = stInput;
			#endregion Decimales para el total
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{				
				switch(e.KeyValue)
				{	
						#region Proveedores
					case (int) Keys.F3:
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "Ruc" && !(bool)this.ultraGrid1.ActiveRow.Cells["SinFactura"].Value && bEdicion)
						{
							using (BuscaClientes miBusqueda = new BuscaClientes(this.ultraGrid1.ActiveRow.Cells["Ruc"].Value.ToString(), 4, 0, 1, 0, false))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{	
									//this.ultraGrid1.ActiveRow.Cells["idProveedor"].Value = miBusqueda.idCliente;
									this.ultraGrid1.ActiveRow.Cells["Ruc"].Value = miBusqueda.Ruc;
									//this.ultraGrid1.ActiveRow.Cells["Nombre"].Value = miBusqueda.Nombre;
								}
							}
						}						
						break;
						#endregion Proveedores
						#region Page Down Abre lista desplegable
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
						break;
						#endregion Page Down Abre lista desplegable
						#region Flechas
					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology");
			}
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACION") Funcion.CamposNumericos(sender, e);
			
			if ((bool)this.ultraGrid1.ActiveRow.Cells["SinFactura"].Value && (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOMBRE" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACION" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO")) e.Handled = true;

			if (!(bool)this.ultraGrid1.ActiveRow.Cells["Diferido"].Value && (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CUOTAS" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "VALORCUOTAS")) e.Handled = true;

			//if ((this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACION" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE")) e.Handled = true;
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			this.cmbBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idBanco, Nombre From Banco Where ISNULL(Pais, '') = 'Ecuador' Order By Nombre");
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
//			this.cmbArticulo1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticulo, Codigo, Articulo From Articulo Where idTipoGrupo In (2, 3, 13)");
//			this.cmbArticulo2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idArticulo, Codigo, Articulo From Articulo Where idTipoGrupo In (2, 3, 13)");
		}

		private void btnCalcular_Click(object sender, System.EventArgs e)
		{
			Totales();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				
				ultraGridExcelExporter1.Export(this.ultraGrid1, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idGastosTarjetas";
			ValIdCompra.Value = (int)this.txtIdGastos.Value;
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);

			Reporte miRep = new Reporte("GastosTarjetas", "");
			miRep.MdiParent = this.MdiParent;			
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();

			this.Cursor = Cursors.Default;			
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetalleGastosTarjetas"].Value = 0;
				e.Row.Cells["SinFactura"].Value = false;
				e.Row.Cells["Point"].Value = true;
				e.Row.Cells["idCompra"].Value = 0;
				e.Row.Cells["idAsiento"].Value = 0;
				e.Row.Cells["idProveedor"].Value = 0;

				#region Servicios
				e.Row.Cells["idDetCompra1"].Value = 0;
				e.Row.Cells["idArticulo1"].Value = 0;	
				e.Row.Cells["IVA1"].Value = 0.00;
				e.Row.Cells["Subtotal1"].Value = 0.00;

				e.Row.Cells["idDetCompra2"].Value = 0;
				e.Row.Cells["idArticulo2"].Value = 0;	
				e.Row.Cells["IVA2"].Value = 0.00;
				e.Row.Cells["Subtotal2"].Value = 0.00;

				e.Row.Cells["idDetCompra3"].Value = 0;
				e.Row.Cells["idArticulo3"].Value = 0;	
				e.Row.Cells["IVA3"].Value = 0.00;
				e.Row.Cells["Subtotal3"].Value = 0.00;

				e.Row.Cells["idDetCompra4"].Value = 0;
				e.Row.Cells["idArticulo4"].Value = 0;	
				e.Row.Cells["IVA4"].Value = 0.00;
				e.Row.Cells["Subtotal4"].Value = 0.00;

				e.Row.Cells["idDetCompra5"].Value = 0;
				e.Row.Cells["idArticulo5"].Value = 0;	
				e.Row.Cells["IVA5"].Value = 0.00;
				e.Row.Cells["Subtotal5"].Value = 0.00;

				e.Row.Cells["idDetCompra6"].Value = 0;
				e.Row.Cells["idArticulo6"].Value = 0;	
				e.Row.Cells["IVA6"].Value = 0.00;
				e.Row.Cells["Subtotal6"].Value = 0.00;

				e.Row.Cells["idDetCompra7"].Value = 0;
				e.Row.Cells["idArticulo7"].Value = 0;	
				e.Row.Cells["IVA7"].Value = 0.00;
				e.Row.Cells["Subtotal7"].Value = 0.00;

				e.Row.Cells["idDetCompra8"].Value = 0;
				e.Row.Cells["idArticulo8"].Value = 0;	
				e.Row.Cells["IVA8"].Value = 0.00;
				e.Row.Cells["Subtotal8"].Value = 0.00;

				e.Row.Cells["idDetCompra9"].Value = 0;
				e.Row.Cells["idArticulo9"].Value = 0;	
				e.Row.Cells["IVA9"].Value = 0.00;
				e.Row.Cells["Subtotal9"].Value = 0.00;

				e.Row.Cells["idDetCompra10"].Value = 0;
				e.Row.Cells["idArticulo10"].Value = 0;	
				e.Row.Cells["IVA10"].Value = 0.00;
				e.Row.Cells["Subtotal10"].Value = 0.00;

				e.Row.Cells["idDetCompra11"].Value = 0;
				e.Row.Cells["idArticulo11"].Value = 0;	
				e.Row.Cells["IVA11"].Value = 0.00;
				e.Row.Cells["Subtotal11"].Value = 0.00;

				e.Row.Cells["idDetCompra12"].Value = 0;
				e.Row.Cells["idArticulo12"].Value = 0;	
				e.Row.Cells["IVA12"].Value = 0.00;
				e.Row.Cells["Subtotal12"].Value = 0.00;

				e.Row.Cells["idDetCompra13"].Value = 0;
				e.Row.Cells["idArticulo13"].Value = 0;	
				e.Row.Cells["IVA13"].Value = 0.00;
				e.Row.Cells["Subtotal13"].Value = 0.00;

				e.Row.Cells["idDetCompra14"].Value = 0;
				e.Row.Cells["idArticulo14"].Value = 0;	
				e.Row.Cells["IVA14"].Value = 0.00;
				e.Row.Cells["Subtotal14"].Value = 0.00;

				e.Row.Cells["idDetCompra15"].Value = 0;
				e.Row.Cells["idArticulo15"].Value = 0;	
				e.Row.Cells["IVA15"].Value = 0.00;
				e.Row.Cells["Subtotal15"].Value = 0.00;
				#endregion Servicios

				e.Row.Cells["SubtotalIVA"].Value = 0.00;	
				e.Row.Cells["SubtotalIVA0"].Value = 0.00;	
				e.Row.Cells["IVA"].Value = 0.00;	
				e.Row.Cells["Total"].Value = 0.00;	
				e.Row.Cells["Interes"].Value = 0.00;	
				e.Row.Cells["ValorTotal"].Value = 0.00;
				e.Row.Cells["Diferido"].Value = false;
				e.Row.Cells["Cuotas"].Value = 1;
				e.Row.Cells["ValorCuotas"].Value = 0.00;
				e.Row.Cells["Saldo"].Value = 0.00;
				e.Row.Cells["Estado"].Value = 0;
			}
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bEdicion) return;
		
			if (e.Cell.Column.ToString() == "UnboundColumn1")
			{	
				if (!(bool)e.Cell.Row.Cells["Point"].Value)
				{
					MessageBox.Show("Si no es de La Empresa el Consumo no puede ingresar Gastos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				using (frmDetalleGastosTarjetas DGT = new frmDetalleGastosTarjetas(
								 (int)e.Cell.Row.Cells["idArticulo1"].Value, (double)e.Cell.Row.Cells["IVA1"].Value, (double)e.Cell.Row.Cells["Subtotal1"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo2"].Value, (double)e.Cell.Row.Cells["IVA2"].Value, (double)e.Cell.Row.Cells["Subtotal2"].Value,
								 (int)e.Cell.Row.Cells["idArticulo3"].Value, (double)e.Cell.Row.Cells["IVA3"].Value, (double)e.Cell.Row.Cells["Subtotal3"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo4"].Value, (double)e.Cell.Row.Cells["IVA4"].Value, (double)e.Cell.Row.Cells["Subtotal4"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo5"].Value, (double)e.Cell.Row.Cells["IVA5"].Value, (double)e.Cell.Row.Cells["Subtotal5"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo6"].Value, (double)e.Cell.Row.Cells["IVA6"].Value, (double)e.Cell.Row.Cells["Subtotal6"].Value,
								 (int)e.Cell.Row.Cells["idArticulo7"].Value, (double)e.Cell.Row.Cells["IVA7"].Value, (double)e.Cell.Row.Cells["Subtotal7"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo8"].Value, (double)e.Cell.Row.Cells["IVA8"].Value, (double)e.Cell.Row.Cells["Subtotal8"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo9"].Value, (double)e.Cell.Row.Cells["IVA9"].Value, (double)e.Cell.Row.Cells["Subtotal9"].Value,
								 (int)e.Cell.Row.Cells["idArticulo10"].Value, (double)e.Cell.Row.Cells["IVA10"].Value, (double)e.Cell.Row.Cells["Subtotal10"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo11"].Value, (double)e.Cell.Row.Cells["IVA11"].Value, (double)e.Cell.Row.Cells["Subtotal11"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo12"].Value, (double)e.Cell.Row.Cells["IVA12"].Value, (double)e.Cell.Row.Cells["Subtotal12"].Value,
								 (int)e.Cell.Row.Cells["idArticulo13"].Value, (double)e.Cell.Row.Cells["IVA13"].Value, (double)e.Cell.Row.Cells["Subtotal13"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo14"].Value, (double)e.Cell.Row.Cells["IVA14"].Value, (double)e.Cell.Row.Cells["Subtotal14"].Value, 
								 (int)e.Cell.Row.Cells["idArticulo15"].Value, (double)e.Cell.Row.Cells["IVA15"].Value, (double)e.Cell.Row.Cells["Subtotal15"].Value,
								 (bool)e.Cell.Row.Cells["SinFactura"].Value))
				{
					if (DialogResult.OK == DGT.ShowDialog())
					{
						if (DGT.uGridGastos.Rows.Count > 0)
						{								
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in DGT.uGridGastos.Rows.All)
							{
								if ((int)dr.Cells["Posicion"].Value == 1)
								{
									e.Cell.Row.Cells["idArticulo1"].Value = (int)dr.Cells["idArticulo"].Value; 
									e.Cell.Row.Cells["IVA1"].Value = (double)dr.Cells["IVA"].Value; 
									e.Cell.Row.Cells["Subtotal1"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 2)
								{
									e.Cell.Row.Cells["idArticulo2"].Value = (int)dr.Cells["idArticulo"].Value;  
									e.Cell.Row.Cells["IVA2"].Value = (double)dr.Cells["IVA"].Value;  
									e.Cell.Row.Cells["Subtotal2"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 3)
								{
									e.Cell.Row.Cells["idArticulo3"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA3"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal3"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								if ((int)dr.Cells["Posicion"].Value == 4)
								{
									e.Cell.Row.Cells["idArticulo4"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA4"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal4"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 5)
								{
									e.Cell.Row.Cells["idArticulo5"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA5"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal5"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 6)
								{
									e.Cell.Row.Cells["idArticulo6"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA6"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal6"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								if ((int)dr.Cells["Posicion"].Value == 7)
								{
									e.Cell.Row.Cells["idArticulo7"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA7"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal7"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 8)
								{
									e.Cell.Row.Cells["idArticulo8"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA8"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal8"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 9)
								{
									e.Cell.Row.Cells["idArticulo9"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA9"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal9"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								if ((int)dr.Cells["Posicion"].Value == 10)
								{
									e.Cell.Row.Cells["idArticulo10"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA10"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal10"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								if ((int)dr.Cells["Posicion"].Value == 11)
								{
									e.Cell.Row.Cells["idArticulo11"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA11"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal11"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 12)
								{
									e.Cell.Row.Cells["idArticulo12"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA12"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal12"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								if ((int)dr.Cells["Posicion"].Value == 13)
								{
									e.Cell.Row.Cells["idArticulo13"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA13"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal13"].Value = (double) dr.Cells["Subtotal"].Value; 
								}

								if ((int)dr.Cells["Posicion"].Value == 14)
								{
									e.Cell.Row.Cells["idArticulo14"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA14"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal14"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								if ((int)dr.Cells["Posicion"].Value == 15)
								{
									e.Cell.Row.Cells["idArticulo15"].Value = (int)dr.Cells["idArticulo"].Value;   
									e.Cell.Row.Cells["IVA15"].Value = (double)dr.Cells["IVA"].Value;   
									e.Cell.Row.Cells["Subtotal15"].Value = (double) dr.Cells["Subtotal"].Value;
								}

								e.Cell.Row.Cells["SubtotalIVA0"].Value = (double) DGT.txtSubtotalIVA0.Value;
								e.Cell.Row.Cells["SubtotalIVA"].Value = (double) DGT.txtSubtotalIVA.Value;
								e.Cell.Row.Cells["IVA"].Value = (double) DGT.txtIVA.Value;
								e.Cell.Row.Cells["Total"].Value = (double) DGT.txtTotal.Value;
							}
						}
					}
				}
			}
		}

		private void cmbTarjeta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTarjeta.ActiveRow == null) return;

			string sSQL = string.Format("Exec TCCListaUsuarios {0}, {1}", (int)this.cmbBanco.Value, (int)this.cmbTarjeta.Value);
			this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void cmbTarjeta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void cmbNombre_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbNombre.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Usuario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbNombre.Focus();
					e.Cancel = true;
					return;
				}
			}
		}

		private void cmbNombre_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Point")
			{
				this.ultraGrid1.UpdateData();

				if (!(bool)e.Cell.Row.Cells["Point"].Value) e.Cell.Row.Cells["SinFactura"].Value = true;
				else e.Cell.Row.Cells["SinFactura"].Value = false;
			}

			if (e.Cell.Column.ToString() == "Diferido")
			{
				this.ultraGrid1.UpdateData();

				if (!(bool)e.Cell.Row.Cells["Diferido"].Value) e.Cell.Row.Cells["Cuotas"].Value = 1;
				else e.Cell.Row.Cells["Cuotas"].Value = 3;
			}
		}
	}
}

