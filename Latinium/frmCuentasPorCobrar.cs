using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmCuentasPorCobrar.
	/// </summary>
	public class frmCuentasPorCobrar : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.RadioGroup rdgEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPlanDePagos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesAbono;
		private System.Windows.Forms.MenuItem mnuAccionesRetencion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.ContextMenu mnuAbrir;
		private System.Windows.Forms.MenuItem mnuAbrirAbono;
		private System.Windows.Forms.MenuItem mnuAbrirRetencion;
		private System.Windows.Forms.MenuItem mnuAbrirAsiento;
		private System.Windows.Forms.MenuItem mnuAbrirImpresión;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label lblFacturas;
		private System.Windows.Forms.Label lblDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label lblBodega;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private System.Windows.Forms.MenuItem mnuAccionesNotaDeCredito;
		private System.Windows.Forms.ContextMenu mnuVer;
		private System.Windows.Forms.MenuItem mnuAccionesFactura;
		private System.Windows.Forms.MenuItem mnuAccionesAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscarDocumento;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.MenuItem mnuAccionesAnticipos;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnCobro;

		CultureInfo us = new CultureInfo("en-US");
		private Acceso miAcceso;
		int idTipoRuc = 0;
		bool bConsulta = false;
		private System.Windows.Forms.MenuItem mnuAccionesCruceCuentas; 
		string sDocumento = "";
		DateTime dtFechaE;

		public frmCuentasPorCobrar(string SDocumento)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			sDocumento = SDocumento; 
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		public frmCuentasPorCobrar()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cartera");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRetencion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retenido");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobro");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MotivoNDC");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDocumento", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAsiento", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnActualizaSaldos", 2);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 16, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 19, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retenido", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retenido", 18, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 21, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cobro", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cobro", 23, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 20, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 21, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cartera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRetencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retenido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MotivoNDC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinEditors.EditorButton editorButton1 = new Infragistics.Win.UltraWinEditors.EditorButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Planes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gastos");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("sTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnAnularCruce", 0);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 11, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 12, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 13, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Mora", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Mora", 14, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Gastos", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Gastos", 15, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Planes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gastos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("sTipo");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPlanPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPlanPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			this.rdgEstado = new DevExpress.XtraEditors.RadioGroup();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridPlanDePagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesAbono = new System.Windows.Forms.MenuItem();
			this.mnuAccionesRetencion = new System.Windows.Forms.MenuItem();
			this.mnuAccionesNotaDeCredito = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAnticipos = new System.Windows.Forms.MenuItem();
			this.mnuAccionesCruceCuentas = new System.Windows.Forms.MenuItem();
			this.lblCliente = new System.Windows.Forms.Label();
			this.mnuAbrir = new System.Windows.Forms.ContextMenu();
			this.mnuAbrirAbono = new System.Windows.Forms.MenuItem();
			this.mnuAbrirRetencion = new System.Windows.Forms.MenuItem();
			this.mnuAbrirAsiento = new System.Windows.Forms.MenuItem();
			this.mnuAbrirImpresión = new System.Windows.Forms.MenuItem();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.lblFacturas = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblBodega = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.mnuVer = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesFactura = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAsiento = new System.Windows.Forms.MenuItem();
			this.txtBuscarDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnCobro = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanDePagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarDocumento)).BeginInit();
			this.SuspendLayout();
			// 
			// rdgEstado
			// 
			this.rdgEstado.EditValue = 1;
			this.rdgEstado.Location = new System.Drawing.Point(568, 0);
			this.rdgEstado.Name = "rdgEstado";
			// 
			// rdgEstado.Properties
			// 
			this.rdgEstado.Properties.Appearance.BackColor = System.Drawing.Color.Transparent;
			this.rdgEstado.Properties.Appearance.Options.UseBackColor = true;
			this.rdgEstado.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Pendiente"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Cancelado"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Todos")});
			this.rdgEstado.Size = new System.Drawing.Size(80, 72);
			this.rdgEstado.TabIndex = 2;
			this.rdgEstado.SelectedIndexChanged += new System.EventHandler(this.rdgEstado_SelectedIndexChanged);
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 44;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 35;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 50;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 40;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance5;
			ultraGridColumn8.Header.Caption = "Número";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "dd/MM/yyyy";
			ultraGridColumn9.Header.Caption = "Emisión";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 65;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Format = "dd/MM/yyyy";
			ultraGridColumn10.Header.Caption = "Vence";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 180;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 64;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Header.Caption = "Forma de Pago";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.Header.VisiblePosition = 25;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance9;
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 90;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance10;
			ultraGridColumn16.Header.Caption = "Lote";
			ultraGridColumn16.Header.VisiblePosition = 14;
			ultraGridColumn16.Width = 35;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance11;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 15;
			ultraGridColumn17.Width = 70;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 16;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 66;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance12;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.Width = 70;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance13;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Width = 70;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance14;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Width = 70;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance15;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 20;
			ultraGridColumn22.Width = 70;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance16;
			ultraGridColumn23.Header.VisiblePosition = 21;
			ultraGridColumn23.Width = 90;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance17;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 22;
			ultraGridColumn24.Width = 70;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance18;
			ultraGridColumn25.Header.Caption = "Tipo";
			ultraGridColumn25.Header.VisiblePosition = 23;
			ultraGridColumn25.Width = 126;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn26.Header.VisiblePosition = 24;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn27.Header.Caption = "D";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 20;
			ultraGridColumn28.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn28.Header.Caption = "A";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn28.Width = 20;
			ultraGridColumn29.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn29.Header.Caption = "S";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn29.Width = 20;
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
																										 ultraGridColumn29});
			ultraGridBand1.Header.Caption = "Documentos";
			ultraGridBand1.HeaderVisible = true;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance19;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance20;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance21;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance22;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.Hidden = true;
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance23;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance24;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance25;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance26;
			this.uGridDocumentos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance27;
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance29.FontData.BoldAsString = "True";
			appearance29.FontData.Name = "Arial";
			appearance29.FontData.SizeInPoints = 8F;
			appearance29.ForeColor = System.Drawing.Color.White;
			appearance29.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance29;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance30.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance30.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.uGridDocumentos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(4, 88);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(1182, 224);
			this.uGridDocumentos.TabIndex = 605;
			this.uGridDocumentos.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridDocumentos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridDocumentos_BeforeSelectChange);
			this.uGridDocumentos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_CellChange);
			this.uGridDocumentos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDocumentos_KeyPress);
			this.uGridDocumentos.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDocumentos_BeforeCellUpdate);
			this.uGridDocumentos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDocumentos_KeyDown);
			this.uGridDocumentos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_ClickCellButton);
			this.uGridDocumentos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDocumentos_AfterCellUpdate);
			this.uGridDocumentos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDocumentos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn9.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(int);
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
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(int);
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
																																 ultraDataColumn26});
			// 
			// txtBuscar
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance33;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Center;
			editorButton1.Appearance = appearance34;
			editorButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			editorButton1.Text = "...";
			this.txtBuscar.ButtonsRight.Add(editorButton1);
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(104, 8);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(144, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			this.txtBuscar.EditorButtonClick += new Infragistics.Win.UltraWinEditors.EditorButtonEventHandler(this.txtBuscar_EditorButtonClick);
			this.txtBuscar.Validated += new System.EventHandler(this.txtBuscar_Validated);
			this.txtBuscar.TextChanged += new System.EventHandler(this.txtBuscar_TextChanged);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(71, 16);
			this.label1.TabIndex = 609;
			this.label1.Text = "Identificación";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource3;
			appearance35.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance35;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Width = 59;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Header.VisiblePosition = 3;
			ultraGridColumn33.Width = 97;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn34.Header.VisiblePosition = 4;
			ultraGridColumn34.Width = 88;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 5;
			ultraGridColumn35.Width = 88;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn36.Header.VisiblePosition = 6;
			ultraGridColumn36.Width = 88;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn37.Header.VisiblePosition = 7;
			ultraGridColumn37.Width = 88;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.VisiblePosition = 10;
			ultraGridColumn38.Width = 67;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.VisiblePosition = 8;
			ultraGridColumn39.Width = 78;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn40.Header.VisiblePosition = 9;
			ultraGridColumn40.Width = 67;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance36;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 11;
			ultraGridColumn41.Width = 82;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance37;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 12;
			ultraGridColumn42.Width = 65;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance38;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 13;
			ultraGridColumn43.Width = 65;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance39;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 14;
			ultraGridColumn44.Width = 65;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance40;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.VisiblePosition = 15;
			ultraGridColumn45.Width = 65;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn46.Header.VisiblePosition = 16;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn47.Header.Caption = "Tipo";
			ultraGridColumn47.Header.VisiblePosition = 17;
			ultraGridColumn47.Width = 99;
			ultraGridColumn48.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn48.Header.Caption = "A";
			ultraGridColumn48.Header.VisiblePosition = 18;
			ultraGridColumn48.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn48.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn48});
			ultraGridBand2.Header.Caption = "Movimientos";
			ultraGridBand2.HeaderVisible = true;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance41;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance42;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance43;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance44;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance45;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.CellAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance49.FontData.BoldAsString = "True";
			appearance49.FontData.Name = "Arial";
			appearance49.FontData.SizeInPoints = 8F;
			appearance49.ForeColor = System.Drawing.Color.White;
			appearance49.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance49;
			appearance50.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance50.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance52;
			this.uGridDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(4, 312);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1182, 184);
			this.uGridDetalle.TabIndex = 611;
			this.uGridDetalle.MouseUp += new System.Windows.Forms.MouseEventHandler(this.uGridDetalle_MouseUp);
			this.uGridDetalle.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridDetalle_BeforeSelectChange);
			this.uGridDetalle.AfterRowActivate += new System.EventHandler(this.uGridDetalle_AfterRowActivate);
			this.uGridDetalle.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_ClickCellButton);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(System.DateTime);
			ultraDataColumn37.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn39.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn40.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn41.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn42.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn43.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
			this.groupBox1.Size = new System.Drawing.Size(1192, 8);
			this.groupBox1.TabIndex = 619;
			this.groupBox1.TabStop = false;
			// 
			// uGridPlanDePagos
			// 
			this.uGridPlanDePagos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPlanDePagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPlanDePagos.DataSource = this.ultraDataSource2;
			appearance53.BackColor = System.Drawing.Color.White;
			this.uGridPlanDePagos.DisplayLayout.Appearance = appearance53;
			this.uGridPlanDePagos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "_DetCompra";
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 51;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Width = 16;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.Caption = "Número";
			ultraGridColumn52.Header.VisiblePosition = 2;
			ultraGridColumn52.Width = 8;
			ultraGridColumn53.Header.VisiblePosition = 3;
			ultraGridColumn53.Width = 8;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance54;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 6;
			ultraGridColumn54.Width = 8;
			ultraGridColumn55.Header.VisiblePosition = 5;
			ultraGridColumn55.Width = 8;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			ultraGridBand3.Header.Caption = "Plan de pagos";
			ultraGridBand3.HeaderVisible = true;
			this.uGridPlanDePagos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridPlanDePagos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridPlanDePagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPlanDePagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance55.BackColor = System.Drawing.Color.Transparent;
			this.uGridPlanDePagos.DisplayLayout.Override.CardAreaAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance56.FontData.BoldAsString = "True";
			appearance56.FontData.Name = "Arial";
			appearance56.FontData.SizeInPoints = 10F;
			appearance56.ForeColor = System.Drawing.Color.White;
			appearance56.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPlanDePagos.DisplayLayout.Override.HeaderAppearance = appearance56;
			appearance57.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance57.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanDePagos.DisplayLayout.Override.RowAlternateAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanDePagos.DisplayLayout.Override.RowSelectorAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPlanDePagos.DisplayLayout.Override.SelectedRowAppearance = appearance59;
			this.uGridPlanDePagos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPlanDePagos.Location = new System.Drawing.Point(1160, 568);
			this.uGridPlanDePagos.Name = "uGridPlanDePagos";
			this.uGridPlanDePagos.Size = new System.Drawing.Size(22, 32);
			this.uGridPlanDePagos.TabIndex = 620;
			this.uGridPlanDePagos.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(System.DateTime);
			ultraDataColumn47.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn50.DataType = typeof(System.Double);
			ultraDataColumn50.DefaultValue = 0;
			ultraDataColumn51.DataType = typeof(System.Double);
			ultraDataColumn51.DefaultValue = 0;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47,
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50,
																																 ultraDataColumn51});
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesAbono,
																																								this.mnuAccionesRetencion,
																																								this.mnuAccionesNotaDeCredito,
																																								this.mnuAccionesAnticipos,
																																								this.mnuAccionesCruceCuentas});
			// 
			// mnuAccionesAbono
			// 
			this.mnuAccionesAbono.Index = 0;
			this.mnuAccionesAbono.Text = "Abono";
			this.mnuAccionesAbono.Click += new System.EventHandler(this.mnuAccionesAbono_Click);
			// 
			// mnuAccionesRetencion
			// 
			this.mnuAccionesRetencion.Index = 1;
			this.mnuAccionesRetencion.Text = "Retención";
			this.mnuAccionesRetencion.Click += new System.EventHandler(this.mnuAccionesRetencion_Click);
			// 
			// mnuAccionesNotaDeCredito
			// 
			this.mnuAccionesNotaDeCredito.Index = 2;
			this.mnuAccionesNotaDeCredito.Text = "Nota de Crédito";
			this.mnuAccionesNotaDeCredito.Click += new System.EventHandler(this.mnuAccionesNotaDeCredito_Click);
			// 
			// mnuAccionesAnticipos
			// 
			this.mnuAccionesAnticipos.Index = 3;
			this.mnuAccionesAnticipos.Text = "Anticipo";
			this.mnuAccionesAnticipos.Visible = false;
			this.mnuAccionesAnticipos.Click += new System.EventHandler(this.mnuAccionesAnticipos_Click);
			// 
			// mnuAccionesCruceCuentas
			// 
			this.mnuAccionesCruceCuentas.Index = 4;
			this.mnuAccionesCruceCuentas.Text = "Cruce de Cuentas";
			this.mnuAccionesCruceCuentas.Click += new System.EventHandler(this.mnuAccionesCruceCuentas_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblCliente.Location = new System.Drawing.Point(8, 39);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(0, 23);
			this.lblCliente.TabIndex = 623;
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// mnuAbrir
			// 
			this.mnuAbrir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																						 this.mnuAbrirAbono,
																																						 this.mnuAbrirRetencion,
																																						 this.mnuAbrirAsiento,
																																						 this.mnuAbrirImpresión});
			// 
			// mnuAbrirAbono
			// 
			this.mnuAbrirAbono.Index = 0;
			this.mnuAbrirAbono.Text = "Abono";
			this.mnuAbrirAbono.Click += new System.EventHandler(this.mnuAbrirAbono_Click);
			// 
			// mnuAbrirRetencion
			// 
			this.mnuAbrirRetencion.Index = 1;
			this.mnuAbrirRetencion.Text = "Retención";
			this.mnuAbrirRetencion.Click += new System.EventHandler(this.mnuAbrirRetencion_Click);
			// 
			// mnuAbrirAsiento
			// 
			this.mnuAbrirAsiento.Index = 2;
			this.mnuAbrirAsiento.Text = "Asiento";
			this.mnuAbrirAsiento.Click += new System.EventHandler(this.mnuAbrirAsiento_Click);
			// 
			// mnuAbrirImpresión
			// 
			this.mnuAbrirImpresión.Index = 3;
			this.mnuAbrirImpresión.Text = "Impresión";
			this.mnuAbrirImpresión.Click += new System.EventHandler(this.mnuAbrirImpresión_Click);
			// 
			// lblFacturas
			// 
			this.lblFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblFacturas.AutoSize = true;
			this.lblFacturas.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFacturas.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFacturas.Location = new System.Drawing.Point(664, 43);
			this.lblFacturas.Name = "lblFacturas";
			this.lblFacturas.Size = new System.Drawing.Size(0, 17);
			this.lblFacturas.TabIndex = 630;
			this.lblFacturas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblFacturas.Click += new System.EventHandler(this.lblFacturas_Click);
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(904, 11);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(84, 16);
			this.lblDescuento.TabIndex = 629;
			this.lblDescuento.Text = "Monto A Cobrar";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblDescuento.Click += new System.EventHandler(this.lblDescuento_Click);
			// 
			// txtValor
			// 
			this.txtValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance60;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.Font = new System.Drawing.Font("Tahoma", 9F);
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(992, 8);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(96, 23);
			this.txtValor.TabIndex = 632;
			this.txtValor.ValueChanged += new System.EventHandler(this.txtValor_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(656, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(26, 16);
			this.lblBodega.TabIndex = 633;
			this.lblBodega.Text = "Caja";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblBodega.Click += new System.EventHandler(this.lblBodega_Click);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn56.Header.VisiblePosition = 2;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Width = 300;
			ultraGridColumn58.Header.VisiblePosition = 0;
			ultraGridColumn58.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(696, 9);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(192, 21);
			this.cmbBodega.TabIndex = 3;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(88, 8);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 637;
			this.txtNumIdCliente.Visible = false;
			// 
			// mnuVer
			// 
			this.mnuVer.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuAccionesFactura,
																																					 this.mnuAccionesAsiento});
			// 
			// mnuAccionesFactura
			// 
			this.mnuAccionesFactura.Index = 0;
			this.mnuAccionesFactura.Text = "Factura";
			this.mnuAccionesFactura.Click += new System.EventHandler(this.mnuAccionesFactura_Click);
			// 
			// mnuAccionesAsiento
			// 
			this.mnuAccionesAsiento.Index = 1;
			this.mnuAccionesAsiento.Text = "Asiento";
			// 
			// txtBuscarDocumento
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscarDocumento.Appearance = appearance61;
			this.txtBuscarDocumento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscarDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscarDocumento.Enabled = false;
			this.txtBuscarDocumento.Location = new System.Drawing.Point(352, 8);
			this.txtBuscarDocumento.Name = "txtBuscarDocumento";
			this.txtBuscarDocumento.Size = new System.Drawing.Size(96, 22);
			this.txtBuscarDocumento.TabIndex = 640;
			this.txtBuscarDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscarDocumento_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(264, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(76, 16);
			this.label2.TabIndex = 641;
			this.label2.Text = "N. Documento";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Enabled = false;
			this.btnExcel.Location = new System.Drawing.Point(1104, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 655;
			this.btnExcel.Text = "Excel";
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(472, 8);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.TabIndex = 656;
			this.btnBuscar.Text = "Ver";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnCobro
			// 
			this.btnCobro.Enabled = false;
			this.btnCobro.Location = new System.Drawing.Point(1104, 8);
			this.btnCobro.Name = "btnCobro";
			this.btnCobro.TabIndex = 657;
			this.btnCobro.Text = "Cobro";
			this.btnCobro.Click += new System.EventHandler(this.btnCobro_Click);
			// 
			// frmCuentasPorCobrar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1190, 504);
			this.Controls.Add(this.btnCobro);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblFacturas);
			this.Controls.Add(this.txtBuscarDocumento);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.uGridPlanDePagos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.rdgEstado);
			this.Controls.Add(this.uGridDocumentos);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmCuentasPorCobrar";
			this.Text = "Cuentas Por Cobrar";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCuentasPorCobrar_KeyDown);
			this.Load += new System.EventHandler(this.frmCuentasPorCobrar_Load);
			((System.ComponentModel.ISupportInitialize)(this.rdgEstado.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPlanDePagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscarDocumento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void LimpiaControles()
		{
      this.lblFacturas.Text = "0 Documentos Seleccionados";						
			this.txtNumIdCliente.Value = 0;
			this.lblCliente.Text = "";
			this.txtValor.Value = 0.00m;
			this.txtBuscarDocumento.Text = "";
			//this.cmbBodega.Value = System.DBNull.Value;
			idTipoRuc = 0;
			bConsulta = false;

			this.uGridDocumentos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CuentasPorCobrar 0, 0, ''");
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DetalleCuentasPorCobrar 0, 0");	
			
			this.btnExcel.Enabled = false;
			this.btnCobro.Enabled = false;
			this.txtBuscarDocumento.Enabled = false;
			this.cmbBodega.Enabled = false;

			this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Retenido"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Cobro"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
		}

		private void Mensaje()
		{
			#region Mensaje
			this.lblFacturas.Text	= "";
			string sEstado = "DOCUMENTO(S)";			
			if ((int)this.rdgEstado.EditValue == 1) sEstado = sEstado + " PENDIENTE(S)";			
			if ((int)this.rdgEstado.EditValue == 2) sEstado = sEstado + " CANCELADO(S)";			
			this.lblFacturas.Text	= this.uGridDocumentos.Rows.Count.ToString() + " " + sEstado;
			#endregion Mensaje
		}

		private void Busqueda()
		{						
			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DetalleCuentasPorCobrar 0, 0");

			string sSQL = string.Format("Exec CuentasPorCobrar {0}, {1}, '{2}'", 
				(int)this.txtNumIdCliente.Value, (int)this.rdgEstado.EditValue, this.txtBuscarDocumento.Text.ToString());
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridDocumentos);
			
			Mensaje();

			#region Habilita Controles
			if ((int)this.rdgEstado.EditValue == 1)
			{
				if (miAcceso.BGenerarCobro)
				{
					this.btnCobro.Enabled = true;
					this.cmbBodega.Enabled = true;
					this.cmbBodega.Value = 1;
					this.txtValor.Value = 0.00m;

					this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Retenido"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Cobro"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

					if (idTipoRuc == 4) this.mnuAcciones.MenuItems[1].Enabled = true; else this.mnuAcciones.MenuItems[1].Enabled = false;
				}
			}
			else
			{
				this.btnCobro.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.txtValor.Value = 0.00m;
				//this.cmbBodega.Value = System.DBNull.Value;

				this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Retenido"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
				this.uGridDocumentos.DisplayLayout.Bands[0].Columns["Cobro"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			}	
		
			if (miAcceso.BExportar) this.btnExcel.Enabled = true;			
			#endregion Habilita Controles

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{				
				if ((int)dr.Cells["idTipoFactura"].Value == 6 || (int)dr.Cells["idTipoFactura"].Value == 9)
				{
					Funcion.EjecutaSQL(cdsSeteoF, 
						string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", 
						(int)dr.Cells["idCompra"].Value, (int)dr.Cells["idTipoFactura"].Value, (int)dr.Cells["idFormaPago"].Value));
				}
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCuentasPorCobrar_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0724");
			
			if (!Funcion.Permiso("53", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Cuentas por Cobrar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				if (miAcceso.BExportar) this.btnExcel.Enabled = true;
				if (miAcceso.BCruzaAnticiposCC) this.mnuAccionesAnticipos.Visible = true;

				Funcion.Resolucion(this, this.Height, this.Width);

				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Codigo From Bodega Where Bodega In (1)");
				
				this.txtBuscar.Focus();

				if (sDocumento.Length > 0)
				{
					this.txtBuscar.Text = sDocumento;
					CargaCliente();		
					Busqueda();//this.btnBuscar_Click(sender, e);
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void DetalleCuentas()
		{
			if (uGridDocumentos.ActiveRow != null)
			{
				if (uGridDocumentos.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value)
				{					
					if ((int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value == 9) this.mnuAcciones.MenuItems[0].Enabled = false;
					else this.mnuAcciones.MenuItems[0].Enabled = true;						

					if (this.uGridDocumentos.ActiveRow.Cells["Estado"].Value.ToString() == "PENDIENTE" && ((int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value == 6 ||(int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value == 23)) this.mnuAcciones.MenuItems[0].Enabled = true;
					else this.mnuAcciones.MenuItems[0].Enabled = false;

					if (this.uGridDocumentos.ActiveRow.Cells["Estado"].Value.ToString() == "CANCELADO") this.mnuAcciones.MenuItems[0].Enabled = false;

					string sSQL = string.Format("Exec DetalleCuentasPorCobrar {0}, {1}", (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, (int)this.uGridDocumentos.ActiveRow.Cells["idTipo"].Value);
					FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridDetalle);
				}
			}
		}

		private void rdgEstado_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			Busqueda();
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
		
		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 1);
				if (sIdentificacion.Length > 0)
				{						
					this.txtBuscar.Text = sIdentificacion;
					
					CargaCliente();				
				}
			}

			if (e.KeyCode == Keys.Enter) 
			{
				this.btnBuscar.Focus();
			}
		}

		private void frmCuentasPorCobrar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridDocumentos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{	
			DetalleCuentas();
		}

		private void uGridDocumentos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Cobro", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Retenido", 2);
		}

		private void mnuAccionesAbono_Click(object sender, System.EventArgs e)
		{			
			if (this.cmbBodega.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Caja", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();				
			}
			else if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();				
			}
			else if (Convert.ToDecimal(this.txtValor.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese el Valor a Cobrar en los Documentos Seleccionados en la Columna Cobro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbBodega.Focus();
				return;
			}
			else
			{
				int iContSel = 0;			

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				{
					if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
					{
						if ((bool)dr.Cells["Sel"].Value) 
						{
							iContSel ++;
						}	
					}
				}

				if (iContSel == 0) 
				{
					MessageBox.Show("Seleccione el o los Documentos que va a Cobrar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.Focus();				
				}			
				else
				{
					bool bValido = false;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
					{
						if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
						{
							if ((bool)dr.Cells["Sel"].Value) 
							{
								if (Convert.ToDecimal(dr.Cells["Cobro"].Value) == 0.00m)
								{
									MessageBox.Show(string.Format("Ingrese el Monto a Cobrar del Documento '{0}'.", dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.uGridDocumentos.ActiveRow = dr;
									this.uGridDocumentos.ActiveRow.Selected = true;
									bValido = true;
								}
							}	
						}
					}
					
					if (!bValido)
					{
						#region Registro de Cobro
						using (VentaFormaDePago miCobro = new VentaFormaDePago (2, Convert.ToDecimal(this.txtValor.Value), 4, true, DateTime.Today, 1, miAcceso.BFiniquitos)) //Convert.ToDecimal(this.uGridDocumentos.ActiveRow.Cells["Saldo"].Value
						{
							if (DialogResult.OK == miCobro.ShowDialog())
							{
								#region Genera Abono Factura 24
								DateTime dtFecha = DateTime.Today;

								string sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 24, 1, 0"));

								string sDetalle = "Cobro Facturas : ";
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
								{
									if ((bool)dr.Cells["Sel"].Value) sDetalle = sDetalle + " - " + dr.Cells["Numero"].Value;
								}

								#region Crea Anticipo
								DateTime dtFechaAbono = (DateTime)miCobro.dtFecha.Value;
								string sSQLPago = string.Format("Exec GuardaAnticipoAbonoClienteCredito {0}, '{1}', {2}, {3}, '{4}', '{5}', '{6}'", 
									0, this.txtBuscar.Text.ToString(), (int)this.cmbBodega.Value, Convert.ToDecimal(miCobro.spnTotal.Value), sNumero, sDetalle, dtFechaAbono.ToString("yyyyMMdd"));
								int idCompraPago = Funcion.iEscalarSQL(cdsSeteoF, sSQLPago);
								#endregion Crea Anticipo
								#endregion Genera Abono Factura 24

								#region Guarda Cobro
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in miCobro.ultraGrid1.Rows.All)
								{
									int idDetFormaPago = 0;
									int idBanco = 0;
									int idTarjeta = 0;
									int idCuenta = 0;
									int idPlan = 0;
									int idPlazo = 0;
									decimal dComTarjeta = 0;
									int idDocumento = 0;
									DateTime dtFecCaducidad = DateTime.Today;
							
									if (dr.Cells["idDetFormaPago"].Value != System.DBNull.Value) idDetFormaPago = (int)dr.Cells["idDetFormaPago"].Value;
									if (dr.Cells["idBanco"].Value != System.DBNull.Value) idBanco = (int)dr.Cells["idBanco"].Value;
									if (dr.Cells["idTarjeta"].Value != System.DBNull.Value) idTarjeta = (int)dr.Cells["idTarjeta"].Value;
									if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;
									if (dr.Cells["idPlan"].Value != System.DBNull.Value) idPlan = (int)dr.Cells["idPlan"].Value;							
									if (dr.Cells["idPlazo"].Value != System.DBNull.Value) idPlazo = (int)dr.Cells["idPlazo"].Value;
									if ((int)dr.Cells["idFormaPago"].Value == 3 || (int)dr.Cells["idFormaPago"].Value == 11) if (dr.Cells["FechaCaducidad"].Value != System.DBNull.Value) dtFecCaducidad = (DateTime)dr.Cells["FechaCaducidad"].Value;
									if (dr.Cells["ValorUsoTarjeta"].Value != System.DBNull.Value) dComTarjeta = Convert.ToDecimal(dr.Cells["ValorUsoTarjeta"].Value);
									if (dr.Cells["idDocumento"].Value != System.DBNull.Value) idDocumento = (int)dr.Cells["idDocumento"].Value;

									string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, {17}, {18}", 
										idDetFormaPago, 
										idCompraPago, 4, 
										(int)dr.Cells["idFormaPago"].Value, Convert.ToDecimal(dr.Cells["Valor"].Value),
										idBanco, idTarjeta,
										dr.Cells["CuentaCorriente"].Value.ToString(), dr.Cells["Numero"].Value.ToString(),
										dtFecCaducidad.ToString("yyyyMMdd"),								
										dr.Cells["NumVoucher"].Value.ToString(), dr.Cells["Autorizacion"].Value.ToString(),
										idCuenta, 
										dr.Cells["Lote"].Value.ToString(), dr.Cells["Referencia"].Value.ToString(),
										idPlan, idPlazo, dComTarjeta, idDocumento);
									Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
								}
								#endregion Guarda Cobro

								#region Variables 
								decimal dValor = Convert.ToDecimal(miCobro.spnTotal.Value);		
								decimal dSaldoFac = 0.00m;
								decimal dAbono = 0.00m;
								#endregion Variables 

								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
								{
									dSaldoFac = 0.00m;
									dAbono = 0.00m;
									if ((bool)dr.Cells["Sel"].Value)
									{
										/* GUARDA HISTORICO DE COBRO EN TABLA ANTICIPOSCOMPRAS DE LA FACTURA ACTUAL */
										string sSQLInsPag = string.Format("Insert Into AnticiposCompras (idAnticipo, idCompra, Valor, Fecha, idTipoFactura) Values ({0}, {1}, {2}, '{3}', {4})", 
											idCompraPago, (int)dr.Cells["idCompra"].Value, Convert.ToDecimal(dr.Cells["Cobro"].Value), dtFechaAbono.ToString("yyyyMMdd"), (int)dr.Cells["idTipoFactura"].Value);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLInsPag);
										/* ACTUALIZA PLAN DE PAGOS DE LA FACTURA ACTUAL */
										string sSQLPP = string.Format("Exec ActualizaPlanDePagos {0}, {1}, {2}, 1", (int)dr.Cells["idCompra"].Value, Convert.ToDecimal(dr.Cells["Cobro"].Value), idCompraPago);
										Funcion.EjecutaSQL(cdsSeteoF, sSQLPP);
										/* ACTUALIZA SALDO DE LA FACTURA ACTUAL */										
										Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 1, 6", (int)dr.Cells["idCompra"].Value));
									}
								}
					
								#region Asiento
								string sSQLAsiento = string.Format("Exec AsientoCXCAbono {0}, 0", idCompraPago);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLAsiento, true);
								#endregion Asiento

								#region Impresion
								ParameterFields paramFields = new ParameterFields ();
				
								ParameterField paramFieldCompra = new ParameterField ();
								ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
								paramFieldCompra.ParameterFieldName = "@idAnticipo";
								ValIdCompra.Value = int.Parse(idCompraPago.ToString());
								paramFieldCompra.CurrentValues.Add (ValIdCompra);
								paramFields.Add (paramFieldCompra);
				
								Reporte miRep = new Reporte("AbonoCXC.rpt", "");
								miRep.MdiParent = this.MdiParent;
								miRep.Titulo("Abono");
								miRep.crVista.ParameterFieldInfo = paramFields;
								miRep.Show();					
								#endregion Impresion

								Busqueda();				
							}
						}
						#endregion Registro de Cobro

						DetalleCuentas();
					}
				}
			}
		}

		private void mnuAccionesRetencion_Click(object sender, System.EventArgs e)
		{			
			#region Validacion
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}

			int iContSel = 0;			

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					if ((bool)dr.Cells["Sel"].Value) 
					{
						iContSel ++;
					}	
				}
			}

			if (iContSel == 0) 
			{
				MessageBox.Show("Debe Seleccionar al menos un Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDocumentos.Focus();
				return;
			}			
			#endregion Validacion

			#region Valida Existencia - Periodo
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec CxCValidaRetencionFactura {0}", (int)dr.Cells["idCompra"].Value), true))
					{
						MessageBox.Show(string.Format("La Factura '{0}' ya tiene Retención por un Valor de {1}", dr.Cells["Numero"].Value.ToString(), Convert.ToDecimal(dr.Cells["Retenido"].Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}

					DateTime dtFechaMaxima = Convert.ToDateTime(dr.Cells["Fecha"].Value).AddDays(5);
					if (Funcion.ValidaPeriodo(dtFechaMaxima, cdsSeteoF, "Venta"))
					{						
						MessageBox.Show(string.Format("No puede ingresar retenciones en esta fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaMaxima.ToString("MMMM").ToUpper()), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);											
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}
				}
			}
			#endregion Valida Existencia - Periodo

			int idCompra = 0;			
			DateTime dtFecha = DateTime.Today;
			if (this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value) idCompra = (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value;
			string sNumero = this.uGridDocumentos.ActiveRow.Cells["Numero"].Value.ToString();
			if (this.uGridDocumentos.ActiveRow.Cells["Fecha"].Value != System.DBNull.Value) dtFecha = Convert.ToDateTime(this.uGridDocumentos.ActiveRow.Cells["Fecha"].Value);

			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				if ((bool)dr.Cells["Sel"].Value)iCont++;
						
			using (frmRetencionesVenta Retencion = new frmRetencionesVenta(this.uGridDocumentos, idCompra, (int)this.txtNumIdCliente.Value, sNumero, dtFecha, Convert.ToDecimal(this.uGridDocumentos.ActiveRow.Cells["Total"].Value), iCont))
			{
				if (DialogResult.OK == Retencion.ShowDialog())
				{
					#region Retenciones
					DateTime dtFechaRet = (DateTime)Retencion.dtFecha.Value;				
					DateTime dtFechaCadR = (DateTime)Retencion.dtCaducidad.Value;

 					#region Guarda Maestro
					string sSQLGR = string.Format("Exec GuardaRetenciones {0}, {1}, 0, 0, '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}", 
						(int) Retencion.txtIdRetencion.Value, idCompra, Retencion.txtSerieFact.Text.ToString(), Retencion.txtNumero.Text.ToString(), Retencion.txtAutFactura.Text.ToString(), 
						dtFechaCadR.ToString("yyyyMMdd"), dtFechaRet.ToString("yyyyMMdd"), "", Convert.ToDecimal(Retencion.txtRetenidoIVA.Value), Convert.ToDecimal(Retencion.txtRetenidoRenta.Value), Convert.ToDecimal(Retencion.txtRetenido.Value));
					Retencion.txtIdRetencion.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLGR);
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

					if (iCont == 1)
					{
						#region Asiento de retencion
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetalleFormaPago Where idCompra = {0} And idFormaPago = 10 And Origen = 1", idCompra)) == 0)
						{
							int idProyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idProyecto From Bodega Where Codigo = '{0}'", this.uGridDocumentos.ActiveRow.Cells["Local"].Value.ToString()));

							string sSQLAsientoR = string.Format("Exec CreaAsientoRetencionVenta 0, {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}, '{9}'", 
								(int)Retencion.txtIdRetencion.Value, (int)this.txtNumIdCliente.Value, (int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value, this.lblCliente.Text.ToString(), Retencion.txtNumero.Text.ToString(), 
								this.uGridDocumentos.ActiveRow.Cells["Numero"].Value.ToString(), dtFechaRet.ToString("yyyyMMdd"), 19, Convert.ToDecimal(Retencion.txtRetenido.Value), this.txtBuscar.Text.ToString());
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAsientoR, true);
						}
						#endregion Asiento de retencion

						string sSQLActValRet = string.Format("Update Compra Set Retenido = {0}, RetenidoIVA = {1}, RetenidoRenta = {2} Where idCompra = {3}", 
							Convert.ToDecimal(Retencion.txtRetenido.Value), Convert.ToDecimal(Retencion.txtRetenidoIVA.Value), Convert.ToDecimal(Retencion.txtRetenidoRenta.Value), (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLActValRet);											

						#region Detalle Forma de Pago
						DateTime dtFecCaducidad = DateTime.Today;
						
						int iFP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFormaPago From Compra Where idCompra = {0}", (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value));//drRetencion["idCompra"])); 
									
						if (iFP != 6 && iFP != 23) 
						{
							string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
								0, (int)Retencion.txtIdRetencion.Value, 6, 1, Convert.ToDecimal(Retencion.txtRetenido.Value), 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
						}
						#endregion Detalle Forma de Pago

						#region Guarda Retencion en el detalleformapago Codigo Anterior
						//					int iFP = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idFormaPago From Compra Where idCompra = {0}", (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value));//drRetencion["idCompra"])); 
						//			
						//					if (iFP == 6) /* SI ES CREDITO GUARDA CON EL ID DE LA FACTURA EL CAMPO IDCOMPRA */ 
						//					{	
						//						//string sSQL1 = string.Format("Exec GeneraPlanDePagosCredito {0}", (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);//(int)drRetencion["idCompra"]); //GeneraPagos 
						//						//Funcion.EjecutaSQL(cdsSeteoF, sSQL1, true);	
						//
						//						DateTime dtFecCaducidad = DateTime.Today;
						//						string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
						//							0, (int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, 4, 10, Convert.ToDecimal(Retencion.txtRetenido.Value, 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "",	0, 0);
						//						Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);		
						//					}
						//					else /* SI NO ES CREDITO GUARDA CON EL ID DE LA RETENCION EL CAMPO IDCOMPRA */ 
						//					{
						//						DateTime dtFecCaducidad = DateTime.Today;
						//						string sSQLCobro = string.Format("Exec GuardaDetalleFormaPago {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', {12}, '{13}', '{14}', {15}, {16}, 0, 0", 
						//							0, (int)Retencion.txtIdRetencion.Value, 4, 1, Convert.ToDecimal(Retencion.txtRetenido.Value, 0, 0, "", "", dtFecCaducidad.ToString("yyyyMMdd"), "", "", 0, "", "", 0, 0);
						//						Funcion.EjecutaSQL(cdsSeteoF, sSQLCobro, true);
						//					}
						#endregion Guarda Retencion en el detalleformapago Codigo Anterior

						#region Actualiza Plan de pagos
						string sSQLPP = string.Format("Exec ActualizaPlanDePagos {0}, {1}, {2}, 4", 
							(int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value, Convert.ToDecimal(Retencion.txtRetenido.Value), (int)Retencion.txtIdRetencion.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLPP);
						#endregion Actualiza Plan de pagos

						string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 1, {1}", idCompra, (int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo, true);						
					}
					else if (iCont > 1)
					{	
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Retenciones Set idCompra = 0 Where idRetencion = {0}", (int)Retencion.txtIdRetencion.Value));

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drRV in Retencion.ultraGrid2.Rows.All)
						{							
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec GuardaRetencionesVenta {0}, {1}, {2}", 
								(int)Retencion.txtIdRetencion.Value, (int)drRV.Cells["idCompra"].Value, Convert.ToDecimal(drRV.Cells["Retenido"].Value)));						
						}

						string sFactura = "";
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drFact in this.uGridDocumentos.Rows.All)
						{
							if ((bool)drFact.Cells["Sel"].Value)
							{
								if (sFactura.Length == 0) sFactura = drFact.Cells["Numero"].Value.ToString();
								else sFactura = sFactura + " - " + drFact.Cells["Numero"].Value.ToString();

								string sSQLSaldo = string.Format("Exec ActualizaSaldoFacturas {0}, 1, {1}", (int)drFact.Cells["idCompra"].Value, (int)drFact.Cells["idFormaPago"].Value);
								Funcion.EjecutaSQL(cdsSeteoF, sSQLSaldo, true);
							}
						}

						string sSQLAsientoR = string.Format("Exec CreaAsientoRetencionVenta 0, {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}, '{9}'", 
							(int)Retencion.txtIdRetencion.Value, (int)this.txtNumIdCliente.Value, 6, this.lblCliente.Text.ToString(), Retencion.txtNumero.Text.ToString(), sFactura, dtFechaRet.ToString("yyyyMMdd"), 19, Convert.ToDecimal(Retencion.txtRetenido.Value), this.txtBuscar.Text.ToString());
						Funcion.EjecutaSQL(cdsSeteoF, sSQLAsientoR, true);
					}

					this.btnBuscar_Click(sender, e);					//DetalleCuentas();					
					#endregion Retenciones
				}
			}
		}

		private void mnuAccionesNotaDeDebito_Click(object sender, System.EventArgs e)
		{
			
		}

		private void uGridDetalle_AfterRowActivate(object sender, System.EventArgs e)
		{

		}

		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridDetalle_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{

		}

		private void mnuAbrirAsiento_Click(object sender, System.EventArgs e)
		{
			frmAsientos Asiento = new frmAsientos((int)this.uGridDetalle.ActiveRow.Cells["idAsiento"].Value);
			Asiento.ShowDialog();
		}

		private void mnuAbrirRetencion_Click(object sender, System.EventArgs e)
		{
			int idRetencion = 0;
			if (this.uGridDetalle.ActiveRow.Cells["idAnticipo"].Value != System.DBNull.Value) idRetencion = (int)this.uGridDetalle.ActiveRow.Cells["idAnticipo"].Value;
			
			frmRetencionesVenta Retencion = new frmRetencionesVenta (idRetencion);
			Retencion.ShowDialog();
		}

		private void mnuAbrirAbono_Click(object sender, System.EventArgs e)
		{
			frmAnticiposClientes miAbono = new frmAnticiposClientes((int)this.uGridDetalle.ActiveRow.Cells["idAnticipo"].Value);
			miAbono.ShowDialog();
		}

		private void mnuAbrirImpresión_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;
			int idAnticipo = (int)this.uGridDetalle.ActiveRow.Cells["idAnticipo"].Value;
			int Tipo = (int)this.uGridDetalle.ActiveRow.Cells["Tipo"].Value;
			
			if (Tipo == 1)
			{
				#region Anticipos
				ParameterFields paramFields = new ParameterFields ();

				ParameterField paramFieldBodega = new ParameterField ();
				ParameterDiscreteValue discreteValBodega = new ParameterDiscreteValue ();
				paramFieldBodega.ParameterFieldName = "@idAnticipo";
				discreteValBodega.Value = idAnticipo; 
				paramFieldBodega.CurrentValues.Add (discreteValBodega);
				paramFields.Add (paramFieldBodega);

				Reporte miRep = new Reporte("ReporteAnticipo.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Anticipo");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
				#endregion Anticipos
			}
			else if (Tipo == 2)
			{
				#region Cobranza
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfAbono = new ParameterField ();
				ParameterDiscreteValue idAbono = new ParameterDiscreteValue ();
				pfAbono.ParameterFieldName = "@idAbono";
				idAbono.Value = idAnticipo;
				pfAbono.CurrentValues.Add (idAbono);
				paramFields.Add (pfAbono);
				
				ParameterField pfAbonoDet = new ParameterField ();
				ParameterDiscreteValue idAbonoDet = new ParameterDiscreteValue ();
				pfAbonoDet.ParameterFieldName = "@idAbonoDet";
				idAbonoDet.Value = idAnticipo;
				pfAbonoDet.CurrentValues.Add (idAbonoDet);
				paramFields.Add (pfAbonoDet);
				
				Reporte miRepor = new Reporte("Cobranzas.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Abono");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.Show();	
				#endregion Cobranza
			}
			else if (Tipo == 3)
			{
				#region Crédito
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idAnticipo";
				ValIdCompra.Value = int.Parse(idAnticipo.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);
				
				Reporte miRep = new Reporte("AbonoCXC.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Abono");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();					
				#endregion Impresion
				#endregion Crédito
			}
			Cursor = Cursors.Default;
		}

		private void uGridDetalle_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
						
		}

		private void Suma()
		{
			this.txtValor.Value = 0.00m;
						
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDocumentos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value && dr.Cells["idCompra"].Value != System.DBNull.Value)
				{
					this.txtValor.Value = Convert.ToDecimal(this.txtValor.Value) + Convert.ToDecimal(dr.Cells["Cobro"].Value);
				}
			}
		}

		private void CuentaFacSel()
		{
			Mensaje();

			int iCont = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridDocumentos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value) iCont ++;
			}

			this.lblFacturas.Text = this.lblFacturas.Text + " - " + iCont.ToString() + " DOCUMENTO(S) SELECCIONADO(S)";
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

		private void uGridDocumentos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCompra"].Value == System.DBNull.Value) return;
			
			if (e.Cell.Column.ToString() == "Retenido")
			{
				e.Cell.Row.Cells["Saldo"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Retenido"].Value);

				Suma();
			}

			if (e.Cell.Column.ToString() == "Cobro")
			{
				decimal dCobro = 0.00m;
				decimal dSaldo = 0.00m;

				dCobro = Convert.ToDecimal(e.Cell.Row.Cells["Cobro"].Value);
				dSaldo = Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value);
 
				if (Math.Round(dCobro, 2) > Math.Round(dSaldo, 2))
				{
					e.Cell.Value = Convert.ToDecimal(e.Cell.OriginalValue);

					MessageBox.Show(string.Format("El Cobro [{0}] no puede ser Mayor al Saldo [{1}]", dCobro, dSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					return;
				}
				
				if (dCobro > 0) e.Cell.Row.Cells["Sel"].Value = true; else if (dCobro == 0) e.Cell.Row.Cells["Sel"].Value = false;

				Suma();
			}
		}

		private void uGridDocumentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDocumentos);
		}

		private void uGridDocumentos_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Retenido")
			{
				if (!(bool)e.Cell.Row.Cells["Sel"].Value) e.Cancel = true;
				if (e.Cell.Row.Cells["idRetencion"].Value != System.DBNull.Value) e.Cancel = true;
			}

			if (e.Cell.Column.ToString() == "Sel")
			{
				if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 5) e.Cancel = true;
				if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 24) e.Cancel = true;
				if (((int)e.Cell.Row.Cells["idTipoFactura"].Value == 1 || (int)e.Cell.Row.Cells["idTipoFactura"].Value == 27)
					&& (int)e.Cell.Row.Cells["idFormaPago"].Value != 6 && (int)e.Cell.Row.Cells["idFormaPago"].Value != 23) e.Cancel = true;
			}
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtBuscar.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtNumIdCliente.Value = int.Parse(sParametros[0]);			
		  if (sParametros[1] != null) idTipoRuc = int.Parse(sParametros[1]);
			if (sParametros[3] != null) this.txtBuscar.Text = sParametros[3];
			if (sParametros[4] != null) this.lblCliente.Text = sParametros[4];			
			
			if ((int)this.txtNumIdCliente.Value > 0) 
			{
				this.txtBuscarDocumento.Enabled = true;
				bConsulta = true; 
				this.Busqueda();
			}
		}

		private void txtBuscar_EditorButtonClick(object sender, Infragistics.Win.UltraWinEditors.EditorButtonEventArgs e)
		{			
			this.LimpiaControles();

			string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 1);

			if (sIdentificacion.Length > 0)
			{						
				this.txtBuscar.Text = sIdentificacion;				

				CargaCliente();
			}
		}

		private void mnuAccionesNotaDeCredito_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();				
				return;
			}

			int iContSel = 0;			

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					if ((bool)dr.Cells["Sel"].Value) iContSel ++;
				}
			}


			if (iContSel == 0) 
			{
				MessageBox.Show("Debe Seleccionar al menos un Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDocumentos.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					if ((bool)dr.Cells["Sel"].Value == true){
					 dtFechaE = Convert.ToDateTime(dr.Cells["Fecha"].Value);
					}
				}
			}


				
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					if ((int)dr.Cells["idFormaPago"].Value != 6 && (int)dr.Cells["idFormaPago"].Value != 23)
					{
						MessageBox.Show("Solo para Formas de Pago Credito - CrediPOINT", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDocumentos.ActiveRow = dr;
						this.uGridDocumentos.ActiveRow.Selected = true;
						return;
					}

					if ((int)dr.Cells["idFormaPago"].Value == 9)
					{
						if ((int)dr.Cells["idEntidadFinanciera"].Value != 2 && (int)dr.Cells["idEntidadFinanciera"].Value != 3)
						{
							MessageBox.Show("Solo para Creditos Solidario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.uGridDocumentos.ActiveRow = dr;
							this.uGridDocumentos.ActiveRow.Selected = true;
							return;
						}
					}
				}
			}		
			#endregion Validacion

			#region Abre Notas de credito
			using (frmBuscaNCCXC miNCPendientes = new frmBuscaNCCXC((int)this.txtNumIdCliente.Value,Convert.ToDateTime(dtFechaE)))
			{
				if (DialogResult.OK == miNCPendientes.ShowDialog())
				{			
					decimal dSaldo = 0.00m;
					decimal dSaldoFactura = 0.00m;
	
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNC in miNCPendientes.uGridNotasDeCredito.Rows.All)
					{
						if ((bool) drNC.Cells["Sel"].Value)
						{
							int IdCompraNC = (int)drNC.Cells["idCompra"].Value;
							DateTime dtFechaNC = (DateTime)drNC.Cells["Fecha"].Value;
							dSaldo = dSaldo + Convert.ToDecimal(drNC.Cells["SaldoAnticipo"].Value);
							
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
							{
								if ((bool) dr.Cells["Sel"].Value)
								{
									if (dSaldo > 0.00m)
									{
										dSaldoFactura = Convert.ToDecimal(dr.Cells["Cobro"].Value); //Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ROUND(SaldoAnticipo, 2) From Compra Where idCompra = {0}", (int)dr.Cells["idCompra"].Value));
										if (dSaldoFactura > 0.00m)
										{								
											decimal dValor = 0.00m;
											if (dSaldo == dSaldoFactura)
											{
												dValor = dSaldo;
												dSaldo = 0.00m;
											}
											else if (dSaldo > dSaldoFactura)
											{
												dValor = dSaldoFactura;
												dSaldo = dSaldo - dSaldoFactura;
											}
											else if (dSaldo < dSaldoFactura)
											{
												dValor = dSaldo;
												dSaldo = 0.00m;
											}

											string sSQL = string.Format("Exec GuardaFacturasNotasDeCredito {0}, {1}, {2}, 0, '{3}', {4}", (int)dr.Cells["idCompra"].Value, IdCompraNC, dValor, dtFechaNC.ToString("yyyyMMdd"), (int)dr.Cells["idTipoFactura"].Value);
											Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

											Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 1, 6", (int)dr.Cells["idCompra"].Value));
										}
									}
								}
							}
					
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 5, 0", IdCompraNC));
						}
					}				

					MessageBox.Show("Proceso Terminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnBuscar_Click(sender, e);
				}
			}
			#endregion Abre Notas de credito			
		}

		private void txtValor_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void mnuAccionesFactura_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros en la lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.uGridDocumentos.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else if (this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{			
				frmVenta Venta = new frmVenta((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
				Venta.MdiParent = MdiParent;
				Venta.Show();
			}
		}

		private void txtBuscarDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) Busqueda();
		}

		private void mnuAccionesAnticipos_Click(object sender, System.EventArgs e)
		{
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();				
			}
			else
			{
				int iContSel = 0;			

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
				{
					if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
					{
						if ((bool)dr.Cells["Sel"].Value) iContSel ++;
					}
				}

				if (iContSel == 0) 
				{
					MessageBox.Show("Debe Seleccionar al menos un Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDocumentos.Focus();
					return;
				}
				else			
				{
					#region Abre Anticipos
					using (frmListaAnticiposCliente AntClte = new frmListaAnticiposCliente((int)this.txtNumIdCliente.Value))
					{
						if (DialogResult.OK == AntClte.ShowDialog())
						{			
							decimal dSaldo = 0;
							decimal dSaldoFactura = 0;
	
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNC in AntClte.uGrid.Rows.All)
							{
								if ((bool)drNC.Cells["Sel"].Value)
								{
									int IdAnticipo = (int)drNC.Cells["idAnticipo"].Value;
									DateTime dtFechaA = (DateTime)drNC.Cells["Fecha"].Value;
									dSaldo = dSaldo + Convert.ToDecimal(drNC.Cells["Saldo"].Value);
							
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
									{
										if ((bool) dr.Cells["Sel"].Value)
										{
											if (dSaldo > 0.00m)
											{
												dSaldoFactura = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ROUND(SaldoAnticipo, 2) From Compra Where idCompra = {0}", (int)dr.Cells["idCompra"].Value));
												if (dSaldoFactura > 0.00m)
												{								
													decimal dValor = 0.00m;
													if (dSaldo == dSaldoFactura)
													{
														dValor = dSaldo;
														dSaldo = 0.00m;
													}
													else if (dSaldo > dSaldoFactura)
													{
														dValor = dSaldoFactura;
														dSaldo = dSaldo - dSaldoFactura;
													}
													else if (dSaldo < dSaldoFactura)
													{
														dValor = dSaldo;
														dSaldo = 0.00m;
													}
												
													string sSQLInsPag = string.Format("Insert Into AnticiposCompras (idAnticipo, idCompra, Valor, Fecha, idTipoFactura) Values ({0}, {1}, {2}, '{3}', {4})", 
														(int)drNC.Cells["idAnticipo"].Value, (int)dr.Cells["idCompra"].Value, dValor, DateTime.Today.ToString("yyyyMMdd"), (int)dr.Cells["idTipoFactura"].Value);
													Funcion.EjecutaSQL(cdsSeteoF, sSQLInsPag);

													Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, {1}, 6", (int)dr.Cells["idCompra"].Value, (int)dr.Cells["idTipoFactura"].Value));
												}
											}
										}
									}
								
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoAnticipos {0}, 1", (int)drNC.Cells["idAnticipo"].Value));								
								}
							}

							MessageBox.Show("Proceso Terminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

							this.btnBuscar_Click(sender, e);
						}
					}
					#endregion Abre Anticipos
				}
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDocumentos);
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			if (this.txtBuscar.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Escriba el Número de Identificación o pulse F3 para ver la lista de clientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();
			}
			else if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Cliente Where Ruc = '{0}' And Proveedor = 0", this.txtBuscar.Text.ToString())) == 0)
			{
				MessageBox.Show("Cliente no existe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscar.Focus();				
			}
			else
			{
				Busqueda();

//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
//				{				
//					if ((int)dr.Cells["idTipoFactura"].Value == 6 || (int)dr.Cells["idTipoFactura"].Value == 9)
//					{
//						Funcion.EjecutaSQL(cdsSeteoF, 
//							string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", 
//							(int)dr.Cells["idCompra"].Value, (int)dr.Cells["idTipoFactura"].Value, (int)dr.Cells["idFormaPago"].Value));
//					}
//				}
			}
		}

		private void btnCobro_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnCobro, new Point(1, 21));
		}

		private void lblFacturas_Click(object sender, System.EventArgs e)
		{
		
		}

		private void lblDescuento_Click(object sender, System.EventArgs e)
		{
		
		}

		private void lblBodega_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_TextChanged(object sender, System.EventArgs e)
		{
			if (bConsulta) this.LimpiaControles();
		}

		private void txtBuscar_Validated(object sender, System.EventArgs e)
		{
			CargaCliente();
		}

		private void uGridDocumentos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Abre Documentos
			if (e.Cell.Column.ToString() == "btnDocumento")
			{
				if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 1)
					{
						frmVenta Venta = new frmVenta((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						Venta.MdiParent = MdiParent;
						Venta.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 5)
					{
						frmDevolucionVenta NotaDeCredito = new frmDevolucionVenta((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						NotaDeCredito.MdiParent = MdiParent;
						NotaDeCredito.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 24)
					{
						frmAnticiposClientes Anticipos = new frmAnticiposClientes((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						Anticipos.MdiParent = MdiParent;
						Anticipos.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 27)
					{
						frmNotaDeDebitoCliente NotaDeDebito = new frmNotaDeDebitoCliente((int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						NotaDeDebito.MdiParent = MdiParent;
						NotaDeDebito.Show();
					}
					if ((int)e.Cell.Row.Cells["idTipoFactura"].Value == 56)
					{
						frmDevolucionAnticipo DevolucionAnticipo = new frmDevolucionAnticipo();//(int)this.uGridDocumentos.ActiveRow.Cells["idCompra"].Value);
						DevolucionAnticipo.MdiParent = MdiParent;
						DevolucionAnticipo.Show();
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

			#region Actualiza Saldos
			if (e.Cell.Column.ToString() == "btnActualizaSaldos")
			{
				if (e.Cell.Row.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					string sSQL = string.Format("Exec ActualizaSaldoFacturas {0}, {1}, {2}", 
						(int)e.Cell.Row.Cells["idCompra"].Value, (int)e.Cell.Row.Cells["idTipoFactura"].Value, (int)e.Cell.Row.Cells["idFormaPago"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);

					MessageBox.Show("Saldo Actualizado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					Busqueda();
				}
			}
			#endregion Actualiza Saldos
		}

		private void mnuAccionesCruceCuentas_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.uGridDocumentos.Rows.Count == 0)
			{
				MessageBox.Show("No existen documentos pendientes.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();				
				return;
			}

			int iContSel = 0;			

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
			{
				if (dr.Cells["idCompra"].Value != System.DBNull.Value) 
				{
					if ((bool)dr.Cells["Sel"].Value) iContSel ++;
				}
			}

			if (iContSel == 0) 
			{
				MessageBox.Show("Debe Seleccionar al menos un Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.uGridDocumentos.Focus();
				return;
			}
			#endregion Validacion

			#region Abre Cruce de Cuentas
			using (frmCruceCuentas CC = new frmCruceCuentas((int)this.txtNumIdCliente.Value, (decimal)this.txtValor.Value))
			{
				if (DialogResult.OK == CC.ShowDialog())
				{			
					decimal dSaldo = 0.00m;
					decimal dSaldoFactura = 0.00m;
					decimal dValorTotal = 0.00m;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drNC in CC.uGridNotasDeCredito.Rows.All)
					{
						if ((bool) drNC.Cells["Sel"].Value)
						{
							#region Cruce de cuentas
							int IdCompraNC = (int)drNC.Cells["idCompra"].Value;
							DateTime dtFechaNC = (DateTime)drNC.Cells["Fecha"].Value;
							dSaldo = dSaldo + Convert.ToDecimal(drNC.Cells["SaldoAnticipo"].Value);

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDocumentos.Rows.All)
							{
								if ((bool) dr.Cells["Sel"].Value)
								{
									if (dSaldo > 0.00m)
									{
										dSaldoFactura = Convert.ToDecimal(dr.Cells["Cobro"].Value); //Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ROUND(SaldoAnticipo, 2) From Compra Where idCompra = {0}", (int)dr.Cells["idCompra"].Value));
										if (dSaldoFactura > 0.00m)
										{								
											decimal dValor = 0.00m;
											if (dSaldo >= dSaldoFactura)
											{
												dValor = dSaldoFactura;
												dSaldo = dSaldo - dSaldoFactura;
											}											
											else if (dSaldo < dSaldoFactura)
											{
												dValor = dSaldo;
												dSaldo = 0.00m;
											}

											string sSQL = string.Format("Exec GuardaFacturaCruceCuentas {0}, {1}, {2}, 0, '{3}', 4", 
												(int)dr.Cells["idCompra"].Value, IdCompraNC, dValor, dtFechaNC.ToString("yyyyMMdd")); //(int)dr.Cells["idTipoFactura"].Value
											Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);	
										
											dValorTotal = dValorTotal + dValor;
										}
									}
								}
							}
							#endregion Cruce de cuentas							
						}
					}

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AsientoCruceDeCuentas {0}", dValorTotal));					

					MessageBox.Show("Proceso Terminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

					this.btnBuscar_Click(sender, e);
				} 
			}
			#endregion Abre Cruce de Cuentas
		}

		private void uGridDocumentos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//if (((int)this.uGridDocumentos.ActiveRow.Cells["idTipoFactura"].Value == 1 || (int)this.uGridDocumentos.ActiveRow.Cells["idTipoFactura"].Value == 27)	&&
			if (((int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value != 6 && (int)this.uGridDocumentos.ActiveRow.Cells["idFormaPago"].Value != 23)
				&& (this.uGridDocumentos.ActiveCell.Column.Key.ToString().ToUpper() == "COBRO")) e.Handled = true;
		}

		private void uGridDetalle_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}
	}
}

