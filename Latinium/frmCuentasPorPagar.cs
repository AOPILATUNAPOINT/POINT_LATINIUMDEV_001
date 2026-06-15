using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCuentasPorPagar.
	/// </summary>
	public class frmCuentasPorPagar : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.ContextMenu mnuVer;
		private System.Windows.Forms.MenuItem mnuVerEgreso;
		private System.Windows.Forms.MenuItem mnuVerICheque;
		private System.Windows.Forms.MenuItem mnuCobrosNC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.MenuItem mnuVerRetencion;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoNotasDeCredito;
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalSaldo;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.MenuItem mnuCobrosAnticipos;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoAnticipos;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox2;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.GroupBox groupBox3;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoFacturas;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoNotasDeDebito;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.ContextMenu mnuPagos;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnPagos;
		private System.Windows.Forms.Button btnActualizarSaldos;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem mnuImprimirFactura;
		private System.Windows.Forms.MenuItem mnuImprimirRIDE;

		string sRUC = "";
		public frmCuentasPorPagar(string SRUC)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			sRUC = SRUC;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmCuentasPorPagar()
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

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenM");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasMora");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAsiento", 1);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 12, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 13, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OrigenM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinEditors.EditorButton editorButton1 = new Infragistics.Win.UltraWinEditors.EditorButton();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHistorialAutorizacionDePago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnEliminar", 1);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 12, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idHistorialAutorizacionDePago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuentaCorriente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PointBeneficiario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaCheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pendiente");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblCliente = new System.Windows.Forms.Label();
			this.mnuVer = new System.Windows.Forms.ContextMenu();
			this.mnuVerEgreso = new System.Windows.Forms.MenuItem();
			this.mnuVerICheque = new System.Windows.Forms.MenuItem();
			this.mnuVerRetencion = new System.Windows.Forms.MenuItem();
			this.mnuPagos = new System.Windows.Forms.ContextMenu();
			this.mnuCobrosNC = new System.Windows.Forms.MenuItem();
			this.mnuCobrosAnticipos = new System.Windows.Forms.MenuItem();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSaldoNotasDeCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtTotalSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.txtSaldoAnticipos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtSaldoNotasDeDebito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSaldoFacturas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnPagos = new System.Windows.Forms.Button();
			this.btnActualizarSaldos = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.mnuImprimirFactura = new System.Windows.Forms.MenuItem();
			this.mnuImprimirRIDE = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNotasDeCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAnticipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNotasDeDebito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoFacturas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance1;
			this.uGridDocumentos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 21;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "Origen";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 45;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.Caption = "Tipo";
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 40;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 62;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Número";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Format = "";
			ultraGridColumn9.Header.Caption = "Emisión";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 65;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Format = "";
			ultraGridColumn10.Header.Caption = "Vence";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Recepción";
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Width = 65;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Retención";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance7;
			ultraGridColumn15.Header.Caption = "D Mora";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 60;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance8;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 100;
			ultraGridColumn17.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn17.Header.Caption = "Doc";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn17.Width = 30;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.Header.Caption = "A C";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn18.Width = 30;
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
																										 ultraGridColumn18});
			ultraGridBand1.Header.Caption = "Documentos";
			ultraGridBand1.HeaderVisible = true;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance9;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance10;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance11;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.25F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance15;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 88);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(872, 224);
			this.uGridDocumentos.TabIndex = 5;
			this.uGridDocumentos.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridDocumentos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridDocumentos_BeforeSelectChange);
			this.uGridDocumentos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_CellChange);
			this.uGridDocumentos.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDocumentos_BeforeCellUpdate);
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout);
			this.uGridDocumentos.Click += new System.EventHandler(this.uGridDocumentos_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn9.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn11.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(int);
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
																																 ultraDataColumn16});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 12);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(71, 16);
			this.label3.TabIndex = 130;
			this.label3.Text = "Identificación";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBuscar
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance19;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Center;
			editorButton1.Appearance = appearance20;
			editorButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			editorButton1.Text = "...";
			this.txtBuscar.ButtonsRight.Add(editorButton1);
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(96, 9);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(144, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.EditorButtonClick += new Infragistics.Win.UltraWinEditors.EditorButtonEventHandler(this.txtBuscar_EditorButtonClick);
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource3;
			appearance21.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance21;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 59;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "Forma de pago";
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 120;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Width = 44;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.Caption = "Forma De Pago";
			ultraGridColumn24.Header.VisiblePosition = 5;
			ultraGridColumn24.Width = 140;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 7;
			ultraGridColumn25.Width = 70;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.Caption = "N. Documento";
			ultraGridColumn26.Header.VisiblePosition = 6;
			ultraGridColumn26.Width = 100;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "N. CH/TR";
			ultraGridColumn27.Header.VisiblePosition = 8;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "Fecha de Pago";
			ultraGridColumn28.Header.VisiblePosition = 9;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 10;
			ultraGridColumn29.Width = 180;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 11;
			ultraGridColumn30.Width = 233;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance22;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 12;
			ultraGridColumn31.Width = 80;
			ultraGridColumn32.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn32.Header.Caption = "D";
			ultraGridColumn32.Header.VisiblePosition = 13;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn32.Width = 31;
			ultraGridColumn33.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn33.Header.Caption = "E";
			ultraGridColumn33.Header.VisiblePosition = 14;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn33.Width = 31;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn33});
			ultraGridBand2.Header.Caption = "Movimientos";
			ultraGridBand2.HeaderVisible = true;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance23;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8.25F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance26;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 320);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1120, 208);
			this.uGridDetalle.TabIndex = 7;
			this.uGridDetalle.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_ClickCellButton);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn17.DefaultValue = 0;
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn18.DefaultValue = 0;
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn19.DefaultValue = 0;
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn20.DefaultValue = 0;
			ultraDataColumn23.DataType = typeof(System.DateTime);
			ultraDataColumn23.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn26.DataType = typeof(System.DateTime);
			ultraDataColumn26.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn29.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn29});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(System.DateTime);
			ultraDataColumn35.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn41.DataType = typeof(System.DateTime);
			ultraDataColumn41.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn43.DataType = typeof(System.Double);
			ultraDataColumn43.DefaultValue = 0;
			ultraDataColumn44.DataType = typeof(bool);
			ultraDataColumn44.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn44});
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1136, 8);
			this.groupBox1.TabIndex = 620;
			this.groupBox1.TabStop = false;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblCliente.Location = new System.Drawing.Point(8, 40);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(0, 23);
			this.lblCliente.TabIndex = 624;
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// mnuVer
			// 
			this.mnuVer.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuVerEgreso,
																																					 this.mnuVerICheque,
																																					 this.mnuVerRetencion});
			// 
			// mnuVerEgreso
			// 
			this.mnuVerEgreso.Index = 0;
			this.mnuVerEgreso.Text = "Asiento - Egreso";
			this.mnuVerEgreso.Click += new System.EventHandler(this.mnuVerEgreso_Click);
			// 
			// mnuVerICheque
			// 
			this.mnuVerICheque.Index = 1;
			this.mnuVerICheque.Text = "Reimprimir Cheque";
			this.mnuVerICheque.Click += new System.EventHandler(this.mnuVerICheque_Click);
			// 
			// mnuVerRetencion
			// 
			this.mnuVerRetencion.Index = 2;
			this.mnuVerRetencion.Text = "Retención";
			// 
			// mnuPagos
			// 
			this.mnuPagos.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																						 this.mnuCobrosNC,
																																						 this.mnuCobrosAnticipos});
			this.mnuPagos.Popup += new System.EventHandler(this.mnuPagos_Popup);
			// 
			// mnuCobrosNC
			// 
			this.mnuCobrosNC.Index = 0;
			this.mnuCobrosNC.Text = "Nota de Crédito";
			this.mnuCobrosNC.Click += new System.EventHandler(this.mnuCobrosNC_Click);
			// 
			// mnuCobrosAnticipos
			// 
			this.mnuCobrosAnticipos.Index = 1;
			this.mnuCobrosAnticipos.Text = "Anticipos";
			this.mnuCobrosAnticipos.Click += new System.EventHandler(this.mnuCobrosAnticipos_Click);
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(576, 9);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 626;
			this.txtNumIdCliente.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(896, 184);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(104, 16);
			this.label4.TabIndex = 627;
			this.label4.Text = "(-) Notas de Crédito";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoNotasDeCredito
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoNotasDeCredito.Appearance = appearance30;
			this.txtSaldoNotasDeCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoNotasDeCredito.Enabled = false;
			this.txtSaldoNotasDeCredito.FormatString = "#,##0.00";
			this.txtSaldoNotasDeCredito.Location = new System.Drawing.Point(1024, 184);
			this.txtSaldoNotasDeCredito.Name = "txtSaldoNotasDeCredito";
			this.txtSaldoNotasDeCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoNotasDeCredito.PromptChar = ' ';
			this.txtSaldoNotasDeCredito.Size = new System.Drawing.Size(96, 22);
			this.txtSaldoNotasDeCredito.TabIndex = 628;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(896, 256);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(53, 16);
			this.label5.TabIndex = 629;
			this.label5.Text = "(=) Total ";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalSaldo
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalSaldo.Appearance = appearance31;
			this.txtTotalSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalSaldo.Enabled = false;
			this.txtTotalSaldo.FormatString = "#,##0.00";
			this.txtTotalSaldo.Location = new System.Drawing.Point(1024, 256);
			this.txtTotalSaldo.Name = "txtTotalSaldo";
			this.txtTotalSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalSaldo.PromptChar = ' ';
			this.txtTotalSaldo.Size = new System.Drawing.Size(96, 22);
			this.txtTotalSaldo.TabIndex = 630;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(1088, 112);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(31, 16);
			this.label6.TabIndex = 631;
			this.label6.Text = "Saldo";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(864, 43);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(72, 16);
			this.lblDescuento.TabIndex = 633;
			this.lblDescuento.Text = "Seleccionado";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(896, 208);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(65, 16);
			this.label8.TabIndex = 636;
			this.label8.Text = "(-) Anticipos";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoAnticipos
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoAnticipos.Appearance = appearance32;
			this.txtSaldoAnticipos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoAnticipos.Enabled = false;
			this.txtSaldoAnticipos.FormatString = "#,##0.00";
			this.txtSaldoAnticipos.Location = new System.Drawing.Point(1024, 208);
			this.txtSaldoAnticipos.Name = "txtSaldoAnticipos";
			this.txtSaldoAnticipos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoAnticipos.PromptChar = ' ';
			this.txtSaldoAnticipos.Size = new System.Drawing.Size(96, 22);
			this.txtSaldoAnticipos.TabIndex = 637;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(896, 160);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(105, 16);
			this.label1.TabIndex = 640;
			this.label1.Text = "(+) Notas de Debito";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoNotasDeDebito
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoNotasDeDebito.Appearance = appearance33;
			this.txtSaldoNotasDeDebito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoNotasDeDebito.Enabled = false;
			this.txtSaldoNotasDeDebito.FormatString = "#,##0.00";
			this.txtSaldoNotasDeDebito.Location = new System.Drawing.Point(1024, 160);
			this.txtSaldoNotasDeDebito.Name = "txtSaldoNotasDeDebito";
			this.txtSaldoNotasDeDebito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoNotasDeDebito.PromptChar = ' ';
			this.txtSaldoNotasDeDebito.Size = new System.Drawing.Size(96, 22);
			this.txtSaldoNotasDeDebito.TabIndex = 641;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(896, 136);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(67, 16);
			this.label2.TabIndex = 642;
			this.label2.Text = "(+) Facturas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoFacturas
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoFacturas.Appearance = appearance34;
			this.txtSaldoFacturas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoFacturas.Enabled = false;
			this.txtSaldoFacturas.FormatString = "#,##0.00";
			this.txtSaldoFacturas.Location = new System.Drawing.Point(1024, 136);
			this.txtSaldoFacturas.Name = "txtSaldoFacturas";
			this.txtSaldoFacturas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoFacturas.PromptChar = ' ';
			this.txtSaldoFacturas.Size = new System.Drawing.Size(96, 22);
			this.txtSaldoFacturas.TabIndex = 643;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(888, 240);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(240, 8);
			this.groupBox2.TabIndex = 646;
			this.groupBox2.TabStop = false;
			// 
			// txtValor
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance35;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(952, 40);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 22);
			this.txtValor.TabIndex = 647;
			// 
			// groupBox3
			// 
			this.groupBox3.Location = new System.Drawing.Point(888, 288);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(240, 8);
			this.groupBox3.TabIndex = 648;
			this.groupBox3.TabStop = false;
			// 
			// txtNumero
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance36;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(360, 9);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 649;
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(264, 12);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(76, 16);
			this.label7.TabIndex = 650;
			this.label7.Text = "N. Documento";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(488, 9);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 651;
			this.btnBuscar.Text = "Ver";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnPagos
			// 
			this.btnPagos.Location = new System.Drawing.Point(1056, 40);
			this.btnPagos.Name = "btnPagos";
			this.btnPagos.TabIndex = 652;
			this.btnPagos.Text = "Pagos";
			this.btnPagos.Click += new System.EventHandler(this.btnPagos_Click);
			// 
			// btnActualizarSaldos
			// 
			this.btnActualizarSaldos.Location = new System.Drawing.Point(952, 9);
			this.btnActualizarSaldos.Name = "btnActualizarSaldos";
			this.btnActualizarSaldos.Size = new System.Drawing.Size(96, 23);
			this.btnActualizarSaldos.TabIndex = 653;
			this.btnActualizarSaldos.Text = "Actualizar Saldos";
			this.btnActualizarSaldos.Click += new System.EventHandler(this.btnActualizarSaldos_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Location = new System.Drawing.Point(1056, 9);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 654;
			this.btnExcel.Text = "Excel";
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// optEstados
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance37;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 0;
			this.optEstados.ItemAppearance = appearance38;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 7;
			valueListItem1.DisplayText = "Por Liquidar";
			valueListItem2.DataValue = 8;
			valueListItem2.DisplayText = "Liquidados";
			valueListItem3.DataValue = 0;
			valueListItem3.DisplayText = "Todos";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(600, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(240, 24);
			this.optEstados.TabIndex = 655;
			this.optEstados.Text = "Por Liquidar";
			this.optEstados.ValueChanged += new System.EventHandler(this.optEstados_ValueChanged);
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuImprimirFactura,
																																								this.mnuImprimirRIDE});
			// 
			// mnuImprimirFactura
			// 
			this.mnuImprimirFactura.Index = 0;
			this.mnuImprimirFactura.Text = "Documentos";
			this.mnuImprimirFactura.Click += new System.EventHandler(this.mnuImprimirFactura_Click);
			// 
			// mnuImprimirRIDE
			// 
			this.mnuImprimirRIDE.Index = 1;
			this.mnuImprimirRIDE.Text = "Movimientos";
			this.mnuImprimirRIDE.Click += new System.EventHandler(this.mnuImprimirRIDE_Click);
			// 
			// frmCuentasPorPagar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1136, 534);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnActualizarSaldos);
			this.Controls.Add(this.btnPagos);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtSaldoFacturas);
			this.Controls.Add(this.txtSaldoNotasDeDebito);
			this.Controls.Add(this.txtSaldoAnticipos);
			this.Controls.Add(this.txtTotalSaldo);
			this.Controls.Add(this.txtSaldoNotasDeCredito);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.uGridDocumentos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCuentasPorPagar";
			this.Text = "Cuentas por Pagar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCuentasPorPagar_KeyDown);
			this.Load += new System.EventHandler(this.frmCuentasPorPagar_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNotasDeCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoAnticipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNotasDeDebito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoFacturas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void BusquedaCliente()
		{
			this.txtNumIdCliente.Value = 0;
			this.lblCliente.Text = "";			
			this.txtValor.Value = 0.00m;

			this.txtSaldoFacturas.Value = 0.00m;
			this.txtSaldoNotasDeDebito.Value = 0.00m;
			this.txtSaldoNotasDeCredito.Value = 0.00m;
			this.txtSaldoAnticipos.Value = 0.00m;
			this.txtTotalSaldo.Value = 0.00m;
						
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(idCliente, 0), ISNULL(Nombre, '') From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtBuscar.Text.ToString()), true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtNumIdCliente.Value = dr.GetInt32(0);
				this.lblCliente.Text = dr.GetString(1);
			}
			dr.Close();
			
			string sSQLResumen = string.Format("Exec CXPSaldos {0}", (int)this.txtNumIdCliente.Value);
			SqlDataReader drResumen = Funcion.rEscalarSQL(cdsSeteoF, sSQLResumen, true);
			while (drResumen.Read())
			{
				if (drResumen.GetInt32(0) == 4) this.txtSaldoFacturas.Value = drResumen.GetDecimal(1);
				if (drResumen.GetInt32(0) == 28) this.txtSaldoNotasDeDebito.Value = drResumen.GetDecimal(1);
				if (drResumen.GetInt32(0) == 6) this.txtSaldoNotasDeCredito.Value = drResumen.GetDecimal(1);
				if (drResumen.GetInt32(0) == 25) this.txtSaldoAnticipos.Value = drResumen.GetDecimal(1);
			}
			drResumen.Close();

			this.txtTotalSaldo.Value = 
				Convert.ToDecimal(this.txtSaldoFacturas.Value) +
				Convert.ToDecimal(this.txtSaldoNotasDeDebito.Value) -
				Convert.ToDecimal(this.txtSaldoNotasDeCredito.Value) -
				Convert.ToDecimal(this.txtSaldoAnticipos.Value);
		}

		private void Busqueda()
		{
			string sSQL = string.Format("Exec CuentasPorPagar '{0}', {1}, '{2}'", 
				this.txtBuscar.Text.ToString(), (int)this.optEstados.Value, this.txtNumero.Text.ToString());
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridDocumentos);

			#region Mensaje
			string sEstado = "";
			if ((int)this.optEstados.Value == 0 && this.uGridDocumentos.Rows.Count == 1) sEstado = "DOCUMENTO";
			if ((int)this.optEstados.Value == 0 && this.uGridDocumentos.Rows.Count != 1) sEstado = "DOCUMENTOS";
			if ((int)this.optEstados.Value == 7 && this.uGridDocumentos.Rows.Count == 1) sEstado = "DOCUMENTO POR LIQUIDAR";
			if ((int)this.optEstados.Value == 7 && this.uGridDocumentos.Rows.Count != 1) sEstado = "DOCUMENTOS POR LIQUIDAR";
			if ((int)this.optEstados.Value == 8 && this.uGridDocumentos.Rows.Count == 1) sEstado = "DOCUMENTO LIQUIDADO";
			if ((int)this.optEstados.Value == 8 && this.uGridDocumentos.Rows.Count != 1) sEstado = "DOCUMENTOS LIQUIDADOS";
			this.uGridDocumentos.DisplayLayout.Bands[0].Header.Caption = this.uGridDocumentos.Rows.Count.ToString() + " " + sEstado;
			#endregion Mensaje
		}

		private void frmCuentasPorPagar_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0611");

			if (sRUC.Length > 0)
			{
				this.txtBuscar.Text = sRUC;
				this.btnBuscar_Click(sender, e);
			}
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{				
				this.btnBuscar_Click(sender, e);
				this.txtNumero.Focus();
			}
		}

		private void uGridDocumentos_Click(object sender, System.EventArgs e)
		{
//			if (this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value)
//				DetalleCuentas();
		}
	
		private void frmCuentasPorPagar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaClientes miBusqueda = new BuscaClientes(this.txtBuscar.Text.ToString(), 4, 1, 1, 1, false))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						this.txtBuscar.Text = miBusqueda.Ruc;
						BusquedaCliente();
						Busqueda();
					}
				}
			}
		}

		private void DetalleCuentas()
		{
			if (uGridDocumentos.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value)
			{				
				string sSQL = string.Format("Exec DetalleCuentasPorPagar {0}, {1}", 
					(int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, 
					(int)this.uGridDocumentos.ActiveRow.Cells["idTipoFactura"].Value);
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridDetalle);
			}
		}

		private void uGridDocumentos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value)
				DetalleCuentas();
		}

		private void ImpresionCheque(decimal fValorCheque, string sNombre, DateTime dtFechaCheque, int idAsiento, int idCuenta)
		{			
			string strFiltro = "{Asiento.idAsiento} = " + idAsiento;
			
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
						
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "ValorCheque";
			discreteVal.Value = fValorCheque;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
			
			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "Nombre";
			discreteVal1.Value = sNombre;
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "Fecha";
			discreteVal2.Value = dtFechaCheque;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			string sSQL = string.Format("Select Reporte From CuentaCorriente Where idCuentaCorriente = {0}", idCuenta);
			string sReporte = Funcion.sEscalarSQL(cdsSeteoF, sSQL);
			
			Reporte miRep = new Reporte(string.Format("{0}.rpt", sReporte), strFiltro);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = MdiParent;
			miRep.Show();
		}

		private void mnuVerEgreso_Click(object sender, System.EventArgs e)
		{
			frmAsientos Asiento = new frmAsientos((int)this.uGridDetalle.ActiveRow.Cells["idAsiento"].Value);
			Asiento.ShowDialog();
		}

		private void mnuVerICheque_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.uGridDetalle.ActiveRow.Cells["FechaCheque"].Value;

			decimal dValor = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Haber From AsientoDetalle ad Inner Join CuentaCorriente cc On cc.idCuenta = ad.idCuenta Where idAsiento = {0} And idCuentaCorriente = {1}", 
											(int)this.uGridDetalle.ActiveRow.Cells["idAsiento"].Value, (int)this.uGridDetalle.ActiveRow.Cells["idCuentaCorriente"].Value));

			ImpresionCheque(dValor, this.uGridDetalle.ActiveRow.Cells["Concepto"].Value.ToString(), (DateTime)this.uGridDetalle.ActiveRow.Cells["FechaCheque"].Value, (int)this.uGridDetalle.ActiveRow.Cells["idAsiento"].Value, (int)this.uGridDetalle.ActiveRow.Cells["idCuentaCorriente"].Value);
		}

		private void mnuCobrosNC_Click(object sender, System.EventArgs e)
		{	
			#region Validacion
			if ((int)this.optEstados.Value != 7) 
			{
				MessageBox.Show("Debe seleccionar el estado Por Liquidar para generar pagos con Notas de Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optEstados.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtSaldoNotasDeCredito.Value) == 0.00m)
			{
				MessageBox.Show("Saldo de Notas de Crédito en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDocumentos.Focus();				
				return;
			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				if ((bool) dr.Cells["Sel"].Value) iCont ++;

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione la(s) Factura(s).", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDocumentos.Focus();				
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{	
					if ((int)dr.Cells["Origen"].Value > 0)
					{
						string sOrigen = "CAJA CHICA";
						if ((int)dr.Cells["Origen"].Value == 2) sOrigen = "TARJETAS DE CREDITO";

						MessageBox.Show(string.Format("No puede generar pagos de facturas generadas desde '{1}',\n\nFactura '{0}'", dr.Cells["Numero"].Value, sOrigen), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}

					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CXPValidaEstadoAutorizacion {0}", (int)dr.Cells["idCompra"].Value)) > 0)
					{
						MessageBox.Show(string.Format("La Factura '{0}' tiene pagos autorizados\n\nComuniquese con el Administrador", dr.Cells["Numero"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Validacion

			bool bAsientoEgresoAnulado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("select ISNULL (CruzeAnticiposVariosProveedores, 0)  from Seteo"), true);

//			if(bAsientoEgresoAnulado)
//			{
//				if (miAcceso.BModificarGastos)
//				{
//          this.txtNumIdCliente.Value = 0;
//				}
//			}

			#region Abre Notas de credito
			using (frmBuscaNCPendientes miNCPendientes = new frmBuscaNCPendientes((int)this.txtNumIdCliente.Value, Convert.ToDecimal(this.txtValor.Value), 6))
			{
				if (DialogResult.OK == miNCPendientes.ShowDialog())
				{			
					#region Documentos seleccionados
					decimal dSaldo = 0.00m;
		
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNC in miNCPendientes.uGridNotasDeCredito.Rows.All)
					{
						if ((bool) drNC.Cells["Sel"].Value)
						{
							int IdCompraNC = (int)drNC.Cells["idCompra"].Value;
							dSaldo = dSaldo + Convert.ToDecimal(drNC.Cells["SaldoAnticipo"].Value);
							string sNumero = drNC.Cells["Numero"].Value.ToString();

							#region Actualiza Saldos documentos seleccionados
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
							{
								if ((bool) dr.Cells["Sel"].Value)
								{
									if (dSaldo > 0)
									{
										if (Convert.ToDecimal(dr.Cells["Saldo"].Value) > 0)
										{									
											string sSQL = string.Format("Exec CruzaNotasCreditoCompra {0}, {1}, {2}, '{3}'", (int)dr.Cells["idCompra"].Value, IdCompraNC, dSaldo, sNumero);
											dSaldo = Funcion.decEscalarSQL(cdsSeteoF, sSQL);

											Funcion.EjecutaSQL(cdsSeteoF, string.Format("ActualizaSaldoFacturas {0}, 4, 6", (int)dr.Cells["idCompra"].Value));
										}
									}
								}
							}
							#endregion Actualiza Saldos documentos seleccionados
					
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("ActualizaSaldoFacturas {0}, 6, 6", IdCompraNC));
						}
					}
					#endregion Documentos seleccionados

					MessageBox.Show("Proceso Terminado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnBuscar_Click(sender, e);
				}
			}
			#endregion Abre Notas de credito			
		}

		private void txtBuscar_EditorButtonClick(object sender, Infragistics.Win.UltraWinEditors.EditorButtonEventArgs e)
		{
			using (BuscaClientes miBusqueda = new BuscaClientes(this.txtBuscar.Text.ToString(), 4, 1, 1, 1, false))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					this.txtBuscar.Text = miBusqueda.Ruc;
					BusquedaCliente();
					Busqueda();
				}
			}
		}

		private void Suma()
		{
			this.txtValor.Value = 0;
						
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDocumentos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value && dr.Cells["idCompra"].Value != System.DBNull.Value)
				{
					this.txtValor.Value = Convert.ToDecimal(this.txtValor.Value) + Convert.ToDecimal(dr.Cells["Saldo"].Value);
				}
			}
		}

		private void CuentaFacSel()
		{
			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDocumentos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value) iCont ++;
			}

			//iCont.ToString() + " DOCUMENTOS SELECCIONADOS.";
		}

		private void mnuCobrosAnticipos_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if ((int)this.optEstados.Value != 7) 
			{
				MessageBox.Show("Debe seleccionar el estado Por Liquidar para generar pagos con Notas de Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optEstados.Focus();
				return;
			}

//			if (Convert.ToDecimal(this.txtSaldoAnticipos.Value) == 0.00m)
//			{
//				MessageBox.Show("Saldo de Anticipos en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.uGridDocumentos.Focus();				
//				return;
//			}

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				if ((bool) dr.Cells["Sel"].Value) iCont ++;

			if (iCont == 0)
			{
				MessageBox.Show("Seleccione la(s) Factura(s).", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDocumentos.Focus();				
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{					
					if ((int)dr.Cells["Origen"].Value > 0)
					{
						string sOrigen = "CAJA CHICA";
						if ((int)dr.Cells["Origen"].Value == 2) sOrigen = "TARJETAS DE CREDITO";

						MessageBox.Show(string.Format("No puede generar pagos de facturas generadas desde '{1}',\n\nFactura '{0}'", dr.Cells["Numero"].Value, sOrigen), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}

					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CXPValidaEstadoAutorizacion {0}", (int)dr.Cells["idCompra"].Value)) > 0)
					{
						MessageBox.Show(string.Format("La Factura '{0}' tiene pagos autorizados\n\nComuniquese con el Administrador", dr.Cells["Numero"].Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Validacion

			bool bAsientoEgresoAnulado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("select ISNULL (CruzeAnticiposVariosProveedores, 0)  from Seteo"), true);

			if(bAsientoEgresoAnulado)
			{
				if (miAcceso.BModificarGastos)
				{
					this.txtNumIdCliente.Value = 0;
				}
			}
			
			using (frmListaAnticiposPendientesProveedor AnticiposPendientes = new frmListaAnticiposPendientesProveedor((int)this.txtNumIdCliente.Value, this.lblCliente.Text.ToString(), Convert.ToDecimal(this.txtValor.Value)))
			{
				if (DialogResult.OK == AnticiposPendientes.ShowDialog())
				{	
					#region Actualiza Saldos en Compras Autorizaciones Anticipos
					decimal dSaldo = 0.00m;
		
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drAnticipos in AnticiposPendientes.uGridAnticipos.Rows.All)
					{
						if ((bool) drAnticipos.Cells["Sel"].Value)
						{
							int IdAnticipo = (int)drAnticipos.Cells["idAnticipo"].Value;
							dSaldo = dSaldo + Convert.ToDecimal(drAnticipos.Cells["Saldo"].Value);
							string sNumero = drAnticipos.Cells["Secuencial"].Value.ToString();
							DateTime dtFecha = DateTime.Today;
							decimal dSaldoInicial = dSaldo;
							decimal dCruzado = 0;

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
							{
								if ((bool) dr.Cells["Sel"].Value)
								{
									if (dSaldo > 0)
									{
										int idTipo = 4;																			
										if ((int)dr.Cells["idTipoFactura"].Value == 28) idTipo = 15;

										if (Convert.ToDecimal(dr.Cells["Saldo"].Value) > 0)
										{		
											if (dSaldo > Convert.ToDecimal(dr.Cells["Saldo"].Value)) dSaldo = Convert.ToDecimal(dr.Cells["Saldo"].Value);

											string sSQL = string.Format("Exec CruzaAnticiposCompra {0}, {1}, {2}, '{3}', {4}", (int)dr.Cells["idCompra"].Value, IdAnticipo, dSaldo, sNumero, idTipo);
											dSaldo = Funcion.decEscalarSQL(cdsSeteoF, sSQL);

											dCruzado = dSaldoInicial - dSaldo;
											dSaldoInicial = dSaldo;											
										}											
									}
								}
							}

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("ActualizaSaldoFacturas {0}, 25, 0", IdAnticipo));

							#region Asiento
							string sSQLAsiento = string.Format("Exec AsientoAnticipoProveedor {0}", IdAnticipo);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
							#endregion Asiento								
						}
					}
					#endregion Actualiza Saldos en Compras Autorizaciones Anticipos
					
					MessageBox.Show("Proceso Terminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnBuscar_Click(sender, e);
				}
			}
		}

		private void uGridDocumentos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value)
			{
				if (e.Cell.Column.ToString() == "Sel")
				{
					this.uGridDocumentos.UpdateData();
					Suma();
					CuentaFacSel();								
				}
			}
		}

		private void uGridDocumentos_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 6 || (int)e.Cell.Row.Cells["idTipoFactura"].Value == 25)
				{
					MessageBox.Show("Solo puede seleccionar Facturas y Notas de Debito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					e.Cancel = true;
				}			
			}
		}
		
		private void uGridDetalle_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnDocumento")
			{
				if (e.Cell.Row.Cells["idHistorialAutorizacionDePago"].Value != System.DBNull.Value)
				{
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 2)
					{
						frmAsientos Asiento = new frmAsientos((int)e.Cell.Row.Cells["Origen"].Value);
						Asiento.MdiParent = MdiParent;
						Asiento.Show();				
					}

					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 4)
					{
						frmCompra Compra = new frmCompra((int)e.Cell.Row.Cells["Origen"].Value, true);
						Compra.MdiParent = MdiParent;
						Compra.Show();
					}

					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 6)
					{
						frmDevolucionCompra NotaDeCredito = new frmDevolucionCompra((int)e.Cell.Row.Cells["Origen"].Value);
						NotaDeCredito.MdiParent = MdiParent;
						NotaDeCredito.Show();
					}
					
					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 8)
					{
						frmAnticiposProveedor Anticipos = new frmAnticiposProveedor((int)e.Cell.Row.Cells["Origen"].Value);
						Anticipos.MdiParent = MdiParent;
						Anticipos.Show();						
					}					

					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 10)
					{
						string sSQLRet = string.Format("Select dbo.ValidaRetencionCompra ({0})", (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						bool iRes = Funcion.bEscalarSQL(cdsSeteoF, sSQLRet, true);
						if (iRes)
						{
							frmRetenciones Retencion = new frmRetenciones((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, this.uGridDocumentos.ActiveRow.Cells["Numero"].Value.ToString(), Convert.ToDecimal(this.uGridDocumentos.ActiveRow.Cells["Total"].Value), true, (DateTime)this.uGridDetalle.ActiveRow.Cells["Fecha"].Value);
							Retencion.MdiParent = MdiParent;
							Retencion.Show();				
						}
					}

					if ((int)e.Cell.Row.Cells["idFormaPago"].Value == 28)
					{
						frmNotaDeDebitoProveedor NotaDeDebito = new frmNotaDeDebitoProveedor((int)e.Cell.Row.Cells["Origen"].Value);
						NotaDeDebito.MdiParent = MdiParent;
						NotaDeDebito.Show();
					}
				}
			}
			
			if (e.Cell.Column.ToString() == "btnEliminar")
			{
				if (e.Cell.Row.Cells["idHistorialAutorizacionDePago"].Value != System.DBNull.Value)
				{
					if (!miAcceso.BAnulaCrucesProveedores)
					{
						MessageBox.Show("No tiene acceso a anular cruces de proveedores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de anular el cruce seleccionado", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
					{
						using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
						{
							if (DialogResult.OK == Nota.ShowDialog())
							{
								try
								{	
									int iTipo = 2;
									if (e.Cell.Row.Cells["Doc"].Value.ToString() == "ANT") iTipo = 3;
									
									string sSQL = string.Format("Exec Prov_AnulaCruces {0}, {1}, '{2}'", (int)e.Cell.Row.Cells["idHistorialAutorizacionDePago"].Value, iTipo, Nota.txtNotas.Text.ToString());
									Funcion.EjecutaSQL(cdsSeteoF, sSQL);
								}
								catch(Exception Exc)
								{
									MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
								}
							}
						}
					}
				}
			}
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Abre Documentos
			if (e.Cell.Column.ToString() == "btnDocumento")
			{
				if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 4)
					{
						frmCompra Compra = new frmCompra((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, true);
						Compra.MdiParent = MdiParent;
						Compra.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 6)
					{
						frmDevolucionCompra NotaDeCredito = new frmDevolucionCompra((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						NotaDeCredito.MdiParent = MdiParent;
						NotaDeCredito.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 25)
					{
						frmAnticiposProveedor Anticipos = new frmAnticiposProveedor((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						Anticipos.MdiParent = MdiParent;
						Anticipos.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 28)
					{
						frmNotaDeDebitoProveedor NotaDeDebito = new frmNotaDeDebitoProveedor((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						NotaDeDebito.MdiParent = MdiParent;
						NotaDeDebito.Show();
					}
				}
			}
			#endregion Abre Documentos

			#region Abre Asientos
			if (e.Cell.Column.ToString() == "btnAsiento")
			{
				if (e.Cell.Row.Cells["idAsiento"].Value != System.DBNull.Value) 
				{
					frmAsientos Asiento = new frmAsientos((int)this.uGridDocumentos.ActiveRow.Cells["idAsiento"].Value);
					Asiento.MdiParent = MdiParent;
					Asiento.Show();
				}
			}
			#endregion Abre Asientos
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.txtNumIdCliente.Value = 0;
			this.lblCliente.Text = "";			
			this.txtValor.Value = 0.00m;

			this.txtSaldoFacturas.Value = 0.00m;
			this.txtSaldoNotasDeDebito.Value = 0.00m;
			this.txtSaldoNotasDeCredito.Value = 0.00m;
			this.txtSaldoAnticipos.Value = 0.00m;
			this.txtTotalSaldo.Value = 0.00m;				

			this.uGridDocumentos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CuentasPorPagar '', 0, ''");
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DetalleCuentasPorPagar 0, 6");
			this.uGridDocumentos.DisplayLayout.Bands[0].Header.Caption = "Documentos";

			if (this.txtBuscar.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba el número de RUC o pulse F3 para ver la lista de proveedores.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();				
			}
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Cliente Where Ruc = '{0}' And Proveedor = 1", this.txtBuscar.Text.ToString())) == 0)
			{
				MessageBox.Show("Proveedor no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();			
			}
			else
			{
				BusquedaCliente();

				Busqueda();
			}
		}

		private void btnPagos_Click(object sender, System.EventArgs e)
		{
			this.mnuPagos.Show(this.btnPagos, new Point(1, 25));
		}

		private void btnActualizarSaldos_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No hay documentos para actualizar saldos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();		
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Actualizar Saldos?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3))
			{	
				this.Cursor = Cursors.WaitCursor;
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				{
					if ((int)this.optEstados.Value == 7)
						if ((int)dr.Cells["Origen"].Value == 0)
							if ((int)dr.Cells["idTipoFactura"].Value == 4)
								if (Convert.ToDecimal(dr.Cells["Saldo"].Value) > 0.00m)
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CXPValidaPagos {0}", (int)dr.Cells["idCompra"].Value));

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, {1}, 0", (int)dr.Cells["idCompra"].Value, (int)dr.Cells["idTipoFactura"].Value));
				}
				
				this.Cursor = Cursors.Default;

				MessageBox.Show("Saldos Actualizados Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnBuscar_Click(sender, e);
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			//FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
			this.mnuImprimir.Show(this.btnExcel, new Point(1, 20));
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnBuscar_Click(sender, e);
		}

		private void optEstados_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtValor.Value = 0.00m;

			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DetalleCuentasPorPagar 0, 6");

			this.btnPagos.Enabled = false;
			this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;

			if ((int)this.txtNumIdCliente.Value > 0)
			{
				if ((int)this.optEstados.Value == 7)
				{
					this.btnPagos.Enabled = true;
					this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				}				
			}

			if (this.txtBuscar.Text.ToString().Length == 0) return;

			Busqueda();
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void mnuImprimirFactura_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void mnuImprimirRIDE_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDetalle);
		}

		private void mnuPagos_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}

