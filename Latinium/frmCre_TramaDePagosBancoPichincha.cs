using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_TramaDePagosBancoPichincha.
	/// </summary>
	public class frmCre_TramaDePagosBancoPichincha : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnContabilizar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPagos;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroDoc;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFechaDoc;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSeleccionados;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipoDoc;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFMora;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIMora;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Button btnImportarCV;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipos;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnRescates;
		private System.ComponentModel.IContainer components;

		public frmCre_TramaDePagosBancoPichincha()
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
		private Acceso miAcceso;
		int iEstado = 0;
		int idCre_LiquidacionPagosPichincha = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_TramaDePagosBancoPichincha));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPagoCuotaCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoLiquidacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Doc");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeOperacion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Secuencial");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroCuota");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorTotal");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debitado");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEnvio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioEnvio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaLiqBco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaLiquida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioLiquida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Error_Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Error_Sistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasMora");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnVer", 0);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 19, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 20, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Mora", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Mora", 21, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorTotal", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorTotal", 22, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debitado", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debitado", 23, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Diferencia", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencia", 24, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "DiasMora", 33, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DiasMora", 33, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPagoCuotaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoLiquidacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Doc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeOperacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Secuencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroCuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEnvio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioEnvio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaLiqBco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaLiquida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioLiquida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Error_Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Error_Sistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasMora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			this.btnImportar = new System.Windows.Forms.Button();
			this.uGridPagos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumeroDoc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtpFechaDoc = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.optTipoPago = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblEstado = new System.Windows.Forms.Label();
			this.chkSeleccionados = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.optTipoDoc = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label3 = new System.Windows.Forms.Label();
			this.dtpDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtFMora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIMora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.btnImportarCV = new System.Windows.Forms.Button();
			this.optTipos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnRescates = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFechaDoc)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optTipoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoDoc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFMora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIMora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipos)).BeginInit();
			this.SuspendLayout();
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Enabled = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(716, 8);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(128, 23);
			this.btnImportar.TabIndex = 731;
			this.btnImportar.Text = "Importar Liquidación";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// uGridPagos
			// 
			this.uGridPagos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPagos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPagos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridPagos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.Caption = "";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 20;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.Caption = "A/M";
			ultraGridColumn5.Header.VisiblePosition = 30;
			ultraGridColumn5.Width = 30;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 31;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 15;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Header.VisiblePosition = 32;
			ultraGridColumn9.Width = 40;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Header.VisiblePosition = 34;
			ultraGridColumn10.Width = 40;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Header.Caption = "Lote";
			ultraGridColumn11.Header.VisiblePosition = 33;
			ultraGridColumn11.Width = 50;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Header.Caption = "Operacion";
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Fecha Doc";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Fecha Cruce";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 50;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance9;
			ultraGridColumn16.Header.Caption = "Numero";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Header.Caption = "Cédula";
			ultraGridColumn17.Header.VisiblePosition = 5;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 6;
			ultraGridColumn18.Width = 200;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn19.CellAppearance = appearance11;
			ultraGridColumn19.Header.Caption = "N. Cuota";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.Width = 60;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance12;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Width = 70;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance13;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 16;
			ultraGridColumn21.Width = 70;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance14;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Width = 70;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance15;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Total";
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn23.Width = 70;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance16;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 20;
			ultraGridColumn24.Width = 70;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance17;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Width = 70;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn26.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn26.Header.Caption = "Fecha Envio";
			ultraGridColumn26.Header.VisiblePosition = 22;
			ultraGridColumn26.Width = 90;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn27.Header.Caption = "Usuario Envio";
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn28.CellAppearance = appearance18;
			ultraGridColumn28.Header.Caption = "N. Documento";
			ultraGridColumn28.Header.VisiblePosition = 24;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn29.Header.Caption = "Fecha Debito";
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn30.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn30.Header.Caption = "Fecha Liquida";
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn30.Width = 100;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.Caption = "Usuario Liquida";
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn31.Width = 100;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn32.Header.Caption = "Error Banco";
			ultraGridColumn32.Header.VisiblePosition = 28;
			ultraGridColumn32.Width = 200;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn33.Header.Caption = "Error Sistema";
			ultraGridColumn33.Header.VisiblePosition = 29;
			ultraGridColumn33.Width = 200;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance19;
			ultraGridColumn34.Header.Caption = "Dias Mora";
			ultraGridColumn34.Header.VisiblePosition = 14;
			ultraGridColumn34.Width = 75;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn35.Header.Caption = "Fecha Registro";
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn35.Width = 120;
			ultraGridColumn36.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance20.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance20.Image = ((object)(resources.GetObject("appearance20.Image")));
			appearance20.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellButtonAppearance = appearance20;
			ultraGridColumn36.Header.Caption = "Ver";
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn36.Width = 26;
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
																										 ultraGridColumn36});
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance21;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance22;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance23;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance24;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance25;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance26;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Center;
			summarySettings7.Appearance = appearance27;
			summarySettings7.DisplayFormat = "{0: #,##0}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridPagos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridPagos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPagos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPagos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance28.BackColor = System.Drawing.Color.Transparent;
			this.uGridPagos.DisplayLayout.Override.CardAreaAppearance = appearance28;
			appearance29.ForeColor = System.Drawing.Color.Black;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagos.DisplayLayout.Override.CellAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance30.FontData.BoldAsString = "True";
			appearance30.FontData.Name = "Arial";
			appearance30.FontData.SizeInPoints = 8.5F;
			appearance30.ForeColor = System.Drawing.Color.White;
			appearance30.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPagos.DisplayLayout.Override.HeaderAppearance = appearance30;
			this.uGridPagos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.RowSelectorAppearance = appearance31;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagos.DisplayLayout.Override.SelectedRowAppearance = appearance32;
			this.uGridPagos.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPagos.Location = new System.Drawing.Point(8, 112);
			this.uGridPagos.Name = "uGridPagos";
			this.uGridPagos.Size = new System.Drawing.Size(1236, 248);
			this.uGridPagos.TabIndex = 732;
			this.uGridPagos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_CellChange);
			this.uGridPagos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridPagos_KeyDown);
			this.uGridPagos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_ClickCellButton);
			this.uGridPagos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridPagos_AfterCellUpdate);
			this.uGridPagos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPagos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn13.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn14.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn19.DataType = typeof(int);
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
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn25.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn26.DataType = typeof(System.DateTime);
			ultraDataColumn26.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn29.DataType = typeof(System.DateTime);
			ultraDataColumn30.DataType = typeof(System.DateTime);
			ultraDataColumn30.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn35});
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(420, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 733;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 74);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 17);
			this.label7.TabIndex = 738;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(200, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(33, 17);
			this.label6.TabIndex = 734;
			this.label6.Text = "Hasta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label6.Visible = false;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 72);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(216, 20);
			this.txtBuscar.TabIndex = 737;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(293, 71);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 736;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtpFecha
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance33;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(256, 8);
			this.dtpFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 735;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			this.dtpFecha.Visible = false;
			this.dtpFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtpFecha_BeforeDropDown);
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(620, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(88, 23);
			this.btnGenerar.TabIndex = 740;
			this.btnGenerar.Text = "Generar TXT";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.Location = new System.Drawing.Point(1148, 55);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 23);
			this.btnContabilizar.TabIndex = 739;
			this.btnContabilizar.Text = "Contabilizar";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
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
			this.groupBox1.Location = new System.Drawing.Point(0, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1252, 8);
			this.groupBox1.TabIndex = 741;
			this.groupBox1.TabStop = false;
			// 
			// optEstado
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance34;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 1;
			this.optEstado.ItemAppearance = appearance35;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Pendiente";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Enviado";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Liquidado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(8, 40);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(304, 26);
			this.optEstado.TabIndex = 742;
			this.optEstado.Text = "Pendiente";
			this.optEstado.ValueChanged += new System.EventHandler(this.optEstado_ValueChanged);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(488, 38);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 743;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(956, 58);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 17);
			this.label1.TabIndex = 747;
			this.label1.Text = "Numero";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(956, 34);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 17);
			this.label2.TabIndex = 744;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumeroDoc
			// 
			this.txtNumeroDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumeroDoc.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNumeroDoc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroDoc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroDoc.Enabled = false;
			this.txtNumeroDoc.Location = new System.Drawing.Point(1020, 56);
			this.txtNumeroDoc.MaxLength = 15;
			this.txtNumeroDoc.Name = "txtNumeroDoc";
			this.txtNumeroDoc.Size = new System.Drawing.Size(112, 20);
			this.txtNumeroDoc.TabIndex = 746;
			// 
			// dtpFechaDoc
			// 
			this.dtpFechaDoc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFechaDoc.Appearance = appearance36;
			this.dtpFechaDoc.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpFechaDoc.DateButtons.Add(dateButton2);
			this.dtpFechaDoc.Format = "dd/MM/yyyy";
			this.dtpFechaDoc.Location = new System.Drawing.Point(1020, 32);
			this.dtpFechaDoc.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFechaDoc.Name = "dtpFechaDoc";
			this.dtpFechaDoc.NonAutoSizeHeight = 23;
			this.dtpFechaDoc.Size = new System.Drawing.Size(112, 21);
			this.dtpFechaDoc.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpFechaDoc.SpinButtonsVisible = true;
			this.dtpFechaDoc.TabIndex = 745;
			this.dtpFechaDoc.Value = ((object)(resources.GetObject("dtpFechaDoc.Value")));
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Controls.Add(this.optTipoPago);
			this.groupBox2.Location = new System.Drawing.Point(948, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(296, 80);
			this.groupBox2.TabIndex = 748;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Liquidación";
			// 
			// optTipoPago
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipoPago.Appearance = appearance37;
			this.optTipoPago.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipoPago.CheckedIndex = 0;
			this.optTipoPago.Enabled = false;
			this.optTipoPago.ItemAppearance = appearance38;
			this.optTipoPago.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem5.DataValue = false;
			valueListItem5.DisplayText = "A";
			valueListItem6.DataValue = true;
			valueListItem6.DisplayText = "M";
			this.optTipoPago.Items.Add(valueListItem5);
			this.optTipoPago.Items.Add(valueListItem6);
			this.optTipoPago.ItemSpacingVertical = 10;
			this.optTipoPago.Location = new System.Drawing.Point(192, 16);
			this.optTipoPago.Name = "optTipoPago";
			this.optTipoPago.Size = new System.Drawing.Size(96, 24);
			this.optTipoPago.TabIndex = 750;
			this.optTipoPago.Text = "A";
			this.optTipoPago.ValueChanged += new System.EventHandler(this.optTipoPago_ValueChanged);
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(736, 64);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 17);
			this.lblEstado.TabIndex = 749;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// chkSeleccionados
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSeleccionados.Appearance = appearance39;
			this.chkSeleccionados.BackColor = System.Drawing.Color.Transparent;
			this.chkSeleccionados.Location = new System.Drawing.Point(320, 42);
			this.chkSeleccionados.Name = "chkSeleccionados";
			this.chkSeleccionados.Size = new System.Drawing.Size(96, 22);
			this.chkSeleccionados.TabIndex = 751;
			this.chkSeleccionados.Text = "Seleccionados";
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.Location = new System.Drawing.Point(56, 112);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(20, 20);
			this.btnSeleccionar.TabIndex = 752;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								 this.menuItem1});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "";
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(852, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 754;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// optTipoDoc
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipoDoc.Appearance = appearance40;
			this.optTipoDoc.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipoDoc.CheckedIndex = 0;
			this.optTipoDoc.ItemAppearance = appearance41;
			this.optTipoDoc.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = -1;
			valueListItem7.DisplayText = "Todos";
			valueListItem8.DataValue = 0;
			valueListItem8.DisplayText = "Abonos";
			valueListItem9.DataValue = 1;
			valueListItem9.DisplayText = "Notas De Crédito";
			this.optTipoDoc.Items.Add(valueListItem7);
			this.optTipoDoc.Items.Add(valueListItem8);
			this.optTipoDoc.Items.Add(valueListItem9);
			this.optTipoDoc.ItemSpacingVertical = 10;
			this.optTipoDoc.Location = new System.Drawing.Point(888, 72);
			this.optTipoDoc.Name = "optTipoDoc";
			this.optTipoDoc.Size = new System.Drawing.Size(32, 26);
			this.optTipoDoc.TabIndex = 755;
			this.optTipoDoc.Text = "Todos";
			this.optTipoDoc.Visible = false;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(35, 17);
			this.label3.TabIndex = 756;
			this.label3.Text = "Desde";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Visible = false;
			// 
			// dtpDesde
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpDesde.Appearance = appearance42;
			this.dtpDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtpDesde.DateButtons.Add(dateButton3);
			this.dtpDesde.Format = "dd/MM/yyyy";
			this.dtpDesde.Location = new System.Drawing.Point(8, 8);
			this.dtpDesde.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.NonAutoSizeHeight = 23;
			this.dtpDesde.Size = new System.Drawing.Size(112, 21);
			this.dtpDesde.SpinButtonsVisible = true;
			this.dtpDesde.TabIndex = 757;
			this.dtpDesde.Value = ((object)(resources.GetObject("dtpDesde.Value")));
			this.dtpDesde.Visible = false;
			// 
			// txtFMora
			// 
			this.txtFMora.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFMora.Location = new System.Drawing.Point(648, 62);
			this.txtFMora.MaxValue = 160;
			this.txtFMora.MinValue = 0;
			this.txtFMora.Name = "txtFMora";
			this.txtFMora.PromptChar = ' ';
			this.txtFMora.Size = new System.Drawing.Size(72, 22);
			this.txtFMora.TabIndex = 961;
			// 
			// txtIMora
			// 
			this.txtIMora.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIMora.Location = new System.Drawing.Point(472, 62);
			this.txtIMora.MaxValue = 105;
			this.txtIMora.MinValue = 0;
			this.txtIMora.Name = "txtIMora";
			this.txtIMora.PromptChar = ' ';
			this.txtIMora.Size = new System.Drawing.Size(72, 22);
			this.txtIMora.TabIndex = 960;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(384, 65);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 17);
			this.label4.TabIndex = 963;
			this.label4.Text = "Dias Mora Inicial";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(560, 65);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(81, 17);
			this.label5.TabIndex = 962;
			this.label5.Text = "Dias Mora Final";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnImportarCV
			// 
			this.btnImportarCV.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportarCV.CausesValidation = false;
			this.btnImportarCV.Enabled = false;
			this.btnImportarCV.Image = ((System.Drawing.Image)(resources.GetObject("btnImportarCV.Image")));
			this.btnImportarCV.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportarCV.Location = new System.Drawing.Point(516, 8);
			this.btnImportarCV.Name = "btnImportarCV";
			this.btnImportarCV.Size = new System.Drawing.Size(96, 23);
			this.btnImportarCV.TabIndex = 964;
			this.btnImportarCV.Text = "Importar XLS";
			this.btnImportarCV.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportarCV.Click += new System.EventHandler(this.btnImportarCV_Click);
			// 
			// optTipos
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipos.Appearance = appearance43;
			this.optTipos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipos.CheckedIndex = 0;
			this.optTipos.ItemAppearance = appearance44;
			this.optTipos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem10.DataValue = 1;
			valueListItem10.DisplayText = "Recaudado";
			valueListItem11.DataValue = 2;
			valueListItem11.DisplayText = "Rescates";
			valueListItem12.DataValue = 3;
			valueListItem12.DisplayText = "Saldos";
			valueListItem13.DataValue = 4;
			valueListItem13.DisplayText = "Retiros";
			this.optTipos.Items.Add(valueListItem10);
			this.optTipos.Items.Add(valueListItem11);
			this.optTipos.Items.Add(valueListItem12);
			this.optTipos.Items.Add(valueListItem13);
			this.optTipos.ItemSpacingVertical = 10;
			this.optTipos.Location = new System.Drawing.Point(8, 8);
			this.optTipos.Name = "optTipos";
			this.optTipos.Size = new System.Drawing.Size(304, 26);
			this.optTipos.TabIndex = 965;
			this.optTipos.Text = "Recaudado";
			this.optTipos.ValueChanged += new System.EventHandler(this.optTipos_ValueChanged);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(324, 8);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(88, 23);
			this.btnNuevo.TabIndex = 966;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnRescates
			// 
			this.btnRescates.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnRescates.CausesValidation = false;
			this.btnRescates.Enabled = false;
			this.btnRescates.Image = ((System.Drawing.Image)(resources.GetObject("btnRescates.Image")));
			this.btnRescates.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRescates.Location = new System.Drawing.Point(840, 40);
			this.btnRescates.Name = "btnRescates";
			this.btnRescates.Size = new System.Drawing.Size(88, 23);
			this.btnRescates.TabIndex = 967;
			this.btnRescates.Text = "Rescates";
			this.btnRescates.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRescates.Click += new System.EventHandler(this.btnRescates_Click);
			// 
			// frmCre_TramaDePagosBancoPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1252, 366);
			this.Controls.Add(this.btnRescates);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.optTipos);
			this.Controls.Add(this.btnImportarCV);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtFMora);
			this.Controls.Add(this.txtIMora);
			this.Controls.Add(this.dtpDesde);
			this.Controls.Add(this.optTipoDoc);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.chkSeleccionados);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtNumeroDoc);
			this.Controls.Add(this.dtpFechaDoc);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.uGridPagos);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.groupBox2);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCre_TramaDePagosBancoPichincha";
			this.Text = "Pagos Banco Pichincha";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_TramaDePagosBancoPichincha_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_TramaDePagosBancoPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridPagos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFechaDoc)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optTipoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipoDoc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFMora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIMora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec Cre_ListaPagosCarteraPichincha '{0}', '{1}', '{2}', {3}, {4}, {5}, {6}, {7}", 
				Convert.ToDateTime(dtpDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(dtpFecha.Value).ToString("yyyyMMdd"), this.txtBuscar.Text.ToString(), this.chkSeleccionados.Checked, 
				(int)this.optEstado.Value, (int)this.optTipos.Value, (int)this.txtIMora.Value, (int)this.txtFMora.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPagos);
						
			this.lblContador.Text = this.uGridPagos.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void frmCre_TramaDePagosBancoPichincha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ContPagos()
		{			      
			int iContPagosPorLiquidar = Funcion.iEscalarSQL(cdsSeteoF, "Exec Cre_ValidaPagosPichinchaPorLiquidar");

			if (iContPagosPorLiquidar > 0) this.lblEstado.Text = string.Format("{0} PAGOS PENDIENTES DE LIQUIDAR", iContPagosPorLiquidar);
		}

		private void frmCre_TramaDePagosBancoPichincha_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320202");

			if (!Funcion.Permiso("962", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso a PAGOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.dtpDesde.Value = DateTime.Today;
			this.dtpFecha.Value = DateTime.Today;	
		
			this.txtIMora.Value = 0;
			this.txtFMora.Value = 100;
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtpFecha, cdsSeteoF, false);

			iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select Top 1 Estado From Cre_LiquidacionPagosPichincha Where CONVERT(Date, Fecha) <= '{0}' Order By Fecha Desc) Select Top 1 Estado From Cre_LiquidacionPagosPichincha Where CONVERT(Date, Fecha) <= '{0}' Order By Fecha Desc Else Select CONVERT(INT, 0)", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd")));

			if (iEstado == 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Cre_LiquidacionPagosPichincha Where CONVERT(Date, Fecha) < '{0}' And Estado <> 2", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"))) == 0)
				{
					this.btnNuevo.Enabled = true;
					this.btnNuevo.Text = "Nuevo";
				}
				else
				{
					this.lblEstado.Text = "Tiene Liquidaciones Abiertas\n\nCierre Estas Liquidaciones Para Generar La Liquidación de Hoy";
				}				
			}
			if (iEstado == 1)
			{
				idCre_LiquidacionPagosPichincha = Funcion.iEscalarSQL(cdsSeteoF, 
					string.Format("Select Top 1 idCre_LiquidacionPagosPichincha From Cre_LiquidacionPagosPichincha Where CONVERT(Date, Fecha) <= '{0}' Order By Fecha Desc", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd")));

				this.btnNuevo.Enabled = true;
				this.btnNuevo.Text = "Cerrar";

				this.btnImportarCV.Enabled = true;
				this.btnGenerar.Enabled = true;
				this.btnImportar.Enabled = true;
				this.btnRescates.Enabled = true;
			}
			if (iEstado == 2)
			{
				this.btnNuevo.Enabled = true;
				this.btnNuevo.Text = "Nuevo";
			}

			//ContPagos();
		}

		private void optEstado_ValueChanged(object sender, System.EventArgs e)
		{
			this.dtpFechaDoc.Value = DateTime.Today;
			this.txtNumeroDoc.Text = "";
			this.optTipoPago.Value = false;
			this.chkSeleccionados.Checked = false;
			this.btnSeleccionar.Enabled = false;
			this.btnGenerar.Enabled = false;
			this.btnImportar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnImportarCV.Enabled = false;
			//this.dtpFechaDoc.Enabled = false;
			this.txtNumeroDoc.Enabled = false;
			this.btnContabilizar.Enabled = false;			
			this.optTipoPago.Enabled = false;
			this.chkSeleccionados.Enabled = false;
			this.btnRescates.Enabled = false;

			this.uGridPagos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			this.uGridPagos.DisplayLayout.Bands[0].Columns["Debitado"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
						
			if ((int)this.optEstado.Value == 1)
			{
				if (iEstado == 1)
				{
					this.chkSeleccionados.Enabled = true;
					this.btnSeleccionar.Enabled = true;
					this.btnGenerar.Enabled = true;
					this.uGridPagos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;

					if ((int)this.optTipos.Value == 2 || (int)this.optTipos.Value == 3)
					{
						this.btnImportarCV.Enabled = true;
						this.btnRescates.Enabled = true;
					}
				}
			}
			else if ((int)this.optEstado.Value == 2)
			{	
				this.chkSeleccionados.Enabled = true;
				this.btnSeleccionar.Enabled = true;				
								
				this.txtNumeroDoc.Enabled = true;
				this.btnContabilizar.Enabled = true;					
				this.btnImportar.Enabled = true;
				this.optTipoPago.Enabled = true;
				this.optTipoPago.Value = false;

				this.uGridPagos.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				this.uGridPagos.DisplayLayout.Bands[0].Columns["Debitado"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
								
				FuncionesProcedimientos.RetornaFechaServidor(this.dtpFechaDoc, cdsSeteoF, false);
			}
					
			this.btnVer_Click(sender, e);
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{	
			//FuncionesProcedimientos.ExportaExcel(this.uGridPagos);

//			if (DialogResult.OK == MessageBox.Show("Esta seguro de GENERAR EL ARCHIVO DE PAGOS", "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
//			{
//				SaveFileDialog sArchivo = new SaveFileDialog();
//
//				sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
//				sArchivo.FileName = "TRAPNT.txt";
//				sArchivo.OverwritePrompt = true;
//				sArchivo.Title = "Genera Archivo de Pagos Banco Pichincha";
//
//				#region Pagos
//				if (sArchivo.ShowDialog() == DialogResult.OK)
//				{
//					using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
//					{
//						#region Genera Archivo
//						int i = 0;
//						foreach (DataRow dr in FuncionesProcedimientos.dtGeneral("Exec Cre_GeneraTramasEnvioPagosManualesNoCobrados").Rows)//FuncionesProcedimientos.dtGeneral("Exec EnvioTempoPagos").Rows) //
//						{
//							decimal dMora = decimal.Parse(dr["Mora"].ToString().Substring(16, 4));
//							string sMora = "";
//							if (dMora > 0) sMora = dr["Mora"].ToString() + ";";
//
//							string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
//								+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
//								+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
//								+ dr["NumeroDeOperacion"].ToString() + ";" + dr["TipoDeTicket"].ToString() + ";" + dr["FechaValor"].ToString() + ";"
//								+ dr["ValorPago"].ToString() + ";" + dr["Capital"].ToString() + ";" + dr["Interes"].ToString() + ";" + sMora.Trim() + dr["ValorRubro"].ToString() + "|";
//
//							file.WriteLine(iLinea);
//
//							i++;					
//						}
//						#endregion Genera Archivo
//
//						MessageBox.Show(string.Format("{0} Tramas de Pagos Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
//					}
//				}
//				#endregion Pagos
//			}
//			return;

//			#region Validacion
//			if (this.uGridPagos.Rows.Count == 0)
//			{
//				MessageBox.Show("No hay Registros para crear el Archivo", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.dtpFecha.Focus();
//				return;
//			}
//
//			if (!Validacion.vbValidaSelGrids(this.uGridPagos, "Sel", "Seleccione el o los pagos a enviar")) return;
//
//			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows)
//			{
//				if ((bool)dr.Cells["Sel"].Value)
//				{
//					if (dr.Cells["NumeroDeOperacion"].Value.ToString().Length == 0)
//					{
//						MessageBox.Show("El o los pagos resaltados no tienen numero de operación", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.uGridPagos.ActiveRow = dr;
//						this.uGridPagos.ActiveRow.Selected = true;
//						return;
//					}
//				}
//			}			
//			#endregion Validacion

			if (DialogResult.OK == MessageBox.Show("Esta seguro de GENERAR EL ARCHIVO DE PAGOS", "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				try
				{
					this.txtBuscar.Text = "";
					decimal dRescates = 0;
															
					SaveFileDialog sArchivo = new SaveFileDialog();

					sArchivo.Filter = "Documentos de texto (*.txt)|*.txt";
					sArchivo.FileName = "TRAPNT.txt";
					sArchivo.OverwritePrompt = true;
					sArchivo.Title = "Genera Archivo de Pagos Banco Pichincha";
				                                
					if (sArchivo.ShowDialog() == DialogResult.OK)
					{
						#region Recaudado
						if ((int)this.optTipos.Value == 1)
						{							
							#region Reporte Recaudacion
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridPagos.Rows.All)
								if (dr.Cells["Debitado"].Value != System.DBNull.Value) dRescates += (decimal) dr.Cells["Debitado"].Value;

							string sSQLAct = string.Format("Exec Cre_InsertaDetalleLiquidacion {0}, 0, 0, 'RECAUDACION', {1}",
								idCre_LiquidacionPagosPichincha, dRescates); 
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
							#endregion Reporte Recaudacion             
							
							#region Pagos
							using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
							{
								#region Normales
								int i = 0;
								foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_TramasPagosCartera '{0}'", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"))).Rows)//FuncionesProcedimientos.dtGeneral("Exec EnvioTempoPagos").Rows) //
								{
									decimal dMora = decimal.Parse(dr["Mora"].ToString().Substring(16, 4));
									string sMora = "";
									if (dMora > 0) sMora = dr["Mora"].ToString() + ";";
									
									//MessageBox.Show(dr["NumeroDeOperacion"].ToString()); + dr["Gastos"].ToString() + ";"

									string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
										+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
										+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
										+ dr["NumeroDeOperacion"].ToString() + ";" + dr["TipoDeTicket"].ToString() + ";" + dr["FechaValor"].ToString() + ";"
										+ dr["ValorPago"].ToString() + ";" + dr["Capital"].ToString() + ";" + dr["Interes"].ToString() + ";" + sMora.Trim() + dr["ValorRubro"].ToString() + "|";

									file.WriteLine(iLinea);

									i++;

									//									string sSQLRef = string.Format("Update PagosCuotasCredito Set Referencia = '{0}' Where idPagoCuotaCredito = {1}", dr["Referencia"].ToString(), Convert.ToInt32(dr["idPagoCuotaCredito"]));
									//									Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);

									string sSQLRef = string.Format("Update PagosCuotasCredito Set Referencia = '{0}', EstadoBanco = 2, UsuarioEnvio = SUSER_SNAME(), FechaEnvio = GETDATE(), Enviado = 1, idPlanPago = 77 Where idCre_TablaDeAmortizacion = {1} And EstadoBanco = 1", 
										dr["Referencia"].ToString(), Convert.ToInt32(dr["idPagoCuotaCredito"]));
									Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);
								}
								return;
								#endregion Normales

								#region Precancelaciones
//								int j = 0;
//								foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_TramasPagosCarteraPreCancelaciones '{0}'", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"))).Rows)
//								{
//									decimal dMora = decimal.Parse(dr["Mora"].ToString().Substring(16, 4));
//									string sMora = "";
//									if (dMora > 0) sMora = dr["Mora"].ToString() + ";";
//          
//									string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
//										+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
//										+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
//										+ dr["NumeroDeOperacion"].ToString() + ";" + dr["TipoDeTicket"].ToString() + ";" + dr["FechaValor"].ToString() + ";"
//										+ dr["ValorPago"].ToString() + ";" + dr["Capital"].ToString() + ";" + dr["Interes"].ToString() + ";" + sMora.Trim() + dr["ValorRubro"].ToString() + "|";
//
//									file.WriteLine(iLinea);
//
//									j++;
//
//									if (Convert.ToInt32(dr["idTipoFactura"]) == 24)
//									{
//										string sSQLRef = string.Format("Update PagosCuotasCredito Set Referencia = '{0}' Where idAnticipo = {1} And Pago = 1", dr["Referencia"].ToString(), Convert.ToInt32(dr["idAnticipo"]));
//										Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);
//									}
//									else if (Convert.ToInt32(dr["idTipoFactura"]) == 5)
//									{
//										string sSQLRef = string.Format("Update PagosCuotasCredito Set Referencia = '{0}' Where NumeroDeOperacion = '{1}' And Pago = 1", dr["Referencia"].ToString(), Convert.ToInt32(dr["OperacionBco"]));
//										Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);
//									}									
//								}
								#endregion Precancelaciones
								
								//MessageBox.Show(string.Format("{0} Tramas de Pagos con {1} Precancelaciones Generadas Correctamente", i, j), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							#endregion Pagos

							#region Actualiza Enviado
							foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows)
								if ((bool)dr.Cells["Sel"].Value)
									Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaEnviadoPagosCartera {0}, {1}, 2", Convert.ToInt32(dr.Cells["idPagoCuotaCredito"].Value), Convert.ToInt32(dr.Cells["idAnticipo"].Value)));
							#endregion Actualiza Enviado
						}
						#endregion Recaudado

						#region Pagos FDG
						if ((int)this.optTipos.Value == 2)
						{   
							#region Reporte FDG
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridPagos.Rows.All)
								if (dr.Cells["Debitado"].Value != System.DBNull.Value) dRescates += (decimal) dr.Cells["Debitado"].Value;

							string sSQLAct = string.Format("Exec Cre_InsertaDetalleLiquidacion {0}, 0, 0, 'RESCATES', {1}",
								idCre_LiquidacionPagosPichincha, dRescates); 
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
							#endregion Reporte FDG

							#region Pagos
							using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
							{
								#region Genera Archivo
								int i = 0;
								foreach (DataRow dr in FuncionesProcedimientos.dtGeneral("Exec Cre_TramasPagosRescates").Rows)//FuncionesProcedimientos.dtGeneral("Exec EnvioTempoPagos").Rows) //
								{
									decimal dMora = decimal.Parse(dr["Mora"].ToString().Substring(16, 4));
									string sMora = "";
									if (dMora > 0) sMora = dr["Mora"].ToString() + ";";

									string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
										+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
										+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
										+ dr["NumeroDeOperacion"].ToString() + ";" + dr["TipoDeTicket"].ToString() + ";" + dr["FechaValor"].ToString() + ";"
										+ dr["ValorPago"].ToString() + ";" + dr["Capital"].ToString() + ";" + dr["Interes"].ToString() + ";" + sMora.Trim() + dr["ValorRubro"].ToString() + "|";

									file.WriteLine(iLinea);

									i++;

									string sSQLRef = string.Format("Exec Cre_ActualizaEstadoEnviosRescates {0}, 2, 0, 1", Convert.ToInt32(dr["idPagoCuotaCredito"]));
									Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);
								}
								#endregion Genera Archivo

								MessageBox.Show(string.Format("{0} Tramas de Pagos Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							#endregion Pagos
						}
						#endregion Pagos FDG

						#region Saldos
						if ((int)this.optTipos.Value == 3)
						{
							#region Reporte Saldos
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridPagos.Rows.All)
								if (dr.Cells["Debitado"].Value != System.DBNull.Value) dRescates += (decimal) dr.Cells["Debitado"].Value;

							string sSQLAct = string.Format("Exec Cre_InsertaDetalleLiquidacion {0}, 0, 0, 'SALDOS Y PRECANCELACIONES', {1}",
								idCre_LiquidacionPagosPichincha, dRescates); 
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
							#endregion Reporte Saldos

							#region Pagos
							using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
							{
								#region Genera Archivo
								int i = 0;
								foreach (DataRow dr in FuncionesProcedimientos.dtGeneral("Exec Cre_TramasPagosSaldos").Rows) // FuncionesProcedimientos.dtGeneral("Exec EnvioTempoPagos").Rows) //FuncionesProcedimientos.dtGeneral("Exec Cre_TramaSaldos").Rows)
								{
									decimal dMora = decimal.Parse(dr["Mora"].ToString().Substring(16, 4));
									string sMora = "";
									if (dMora > 0) sMora = dr["Mora"].ToString() + ";";

									string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
										+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
										+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
										+ dr["NumeroDeOperacion"].ToString() + ";" + dr["TipoDeTicket"].ToString() + ";" + dr["FechaValor"].ToString() + ";"
										+ dr["ValorPago"].ToString() + ";" + dr["Capital"].ToString() + ";" + dr["Interes"].ToString() + ";" + sMora.Trim() + dr["ValorRubro"].ToString() + "|";

									file.WriteLine(iLinea);

									i++;
								
									string sSQLRef = string.Format("Exec Cre_ActualizaEstadoEnviosRescates {0}, 2, 0, 2", Convert.ToInt32(dr["idPagoCuotaCredito"]));
									Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);
								}
								#endregion Genera Archivo
								
								MessageBox.Show(string.Format("{0} Tramas de Pagos Generadas Correctamente", i), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							#endregion Pagos
						}
						#endregion Saldos

						#region Precancelaciones de Retiro
						if ((int)this.optTipos.Value == 4)
						{
							#region Pagos
							using (StreamWriter file = new StreamWriter(sArchivo.FileName, true))
							{
								#region Precancelaciones
								int j = 0;
								foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec Cre_TramasPagosCarteraPreCancelaciones '{0}'", Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"))).Rows)
								{
									decimal dMora = decimal.Parse(dr["Mora"].ToString().Substring(16, 4));
									string sMora = "";
									if (dMora > 0) sMora = dr["Mora"].ToString() + ";";
          
									string iLinea = dr["Fecha"].ToString() + "\t" + dr["Peticion"].ToString() + "\t" + dr["NumeroLote"].ToString() + "\t" + dr["Producto"].ToString() + "\t"
										+ dr["Referencia"].ToString() + "\t" + dr["AgenciaIngreso"].ToString() + "\t" + dr["Usuario"].ToString() + "\t"
										+ dr["FechaIngreso"].ToString() + "\t" + dr["HoraIngreso"].ToString() + "\t" + dr["SistemaIngreso"].ToString() + "\t"
										+ dr["NumeroDeOperacion"].ToString() + ";" + dr["TipoDeTicket"].ToString() + ";" + dr["FechaValor"].ToString() + ";"
										+ dr["ValorPago"].ToString() + ";" + dr["Capital"].ToString() + ";" + dr["Interes"].ToString() + ";" + sMora.Trim() + dr["ValorRubro"].ToString() + "|";

									file.WriteLine(iLinea);

									j++;

									string sSQLRef = string.Format("Update Compra Set RetornaPoint = 2 Where idCompra = '{0}'", dr["idAnticipo"].ToString());
									Funcion.EjecutaSQL(cdsSeteoF, sSQLRef);									
								}
								#endregion Precancelaciones
								
								MessageBox.Show(string.Format("{0} Tramas de Precancelaciones Generadas Correctamente", j), "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
							#endregion Precancelaciones
						}
						#endregion Precancelaciones de Retiro							

						this.btnVer_Click(sender, e);
					}
				}
				catch (Exception exc)
				{
					MessageBox.Show(exc.Message);
				}
			}    
		}
    
		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{                       
			if (!Validacion.vbTexto(this.txtNumeroDoc, 5, 15, "Numero de documento")) return;

			if (!Validacion.vbValidaSelGrids(this.uGridPagos, "Sel", "Seleccione el o los pagos a contabilizar")) return;

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de contabilizar los pagos seleccionados", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				Cursor = Cursors.WaitCursor;

				#region Crea Asiento
				decimal dRescates = 0;
				decimal dCapital = 0;
				decimal dInteres = 0;
				decimal dMora = 0;
								
				#region Asiento Pagos
				if ((int)this.optTipos.Value == 1)
				{
					#region Liquidacion Recaudacion
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridPagos.Rows.All)
						if (dr.Cells["Debitado"].Value != System.DBNull.Value) dRescates += (decimal) dr.Cells["Debitado"].Value;

					string sSQLAct = string.Format("Exec Cre_InsertaDetalleLiquidacion {0}, 1, {1}, 'RECAUDACION', {2}",
						idCre_LiquidacionPagosPichincha, (bool)this.optTipoPago.Value, dRescates); 
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
					#endregion Liquidacion Recaudacion

					int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EncabezadoAsientoDevolucionCobranzaPichincha '{0}', '{1}', 50", 
						Convert.ToDateTime(this.dtpFechaDoc.Value).ToString("yyyyMMdd"), this.txtNumeroDoc.Text.ToString()));
					
					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows)
					{
						if (Convert.ToBoolean(dr.Cells["Sel"].Value))
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec DetalleAsientoDevolucionCobranzaPichincha {0}, {1}, {2}, '{3}', 1",
								idAsiento, 
								Convert.ToInt32(dr.Cells["idPagoCuotaCredito"].Value),
								Convert.ToDecimal(dr.Cells["Debitado"].Value), 
								(Convert.ToString(dr.Cells["Ruc"].Value) + "-" + Convert.ToString(dr.Cells["Nombre"].Value))));

							#region Actualiza Liquidado
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaEnviadoPagosCartera {0}, {1}, 3", Convert.ToInt32(dr.Cells["idPagoCuotaCredito"].Value), Convert.ToInt32(dr.Cells["idAnticipo"].Value)));					
							#endregion Actualiza Liquidado

							dCapital = dCapital + Convert.ToDecimal(dr.Cells["Capital"].Value);
							dInteres = dInteres + Convert.ToDecimal(dr.Cells["Interes"].Value);
							dMora = dMora + Convert.ToDecimal(dr.Cells["Mora"].Value);
						}
					}
				
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec BancoAsientoDevolucionCobranzaPichincha {0}", idAsiento));

					string sSQLCO = string.Format("Exec AsientoCuentasDeOrdenEnvioPagos 0, '{0}', '{1}', {2}, {3}, {4}, 1", 
						this.txtNumeroDoc.Text.ToString(), Convert.ToDateTime(this.dtpFechaDoc.Value).ToString("yyyyMMdd"), dCapital, dInteres, dMora);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCO, true);

					string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
				
					Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
					Reporte.MdiParent = this.MdiParent;			
					Reporte.Show();
				}
				#endregion Asiento Pagos

				#region Asiento Pagos FDG
				if ((int)this.optTipos.Value == 2)
				{
					#region Liquidacion Rescates
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridPagos.Rows.All)
						if (dr.Cells["Debitado"].Value != System.DBNull.Value) dRescates += (decimal) dr.Cells["Debitado"].Value;

					string sSQLAct = string.Format("Exec Cre_InsertaDetalleLiquidacion {0}, 1, {1}, 'RESCATES', {2}",
						idCre_LiquidacionPagosPichincha, (bool)this.optTipoPago.Value, dRescates); 
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
					#endregion Liquidacion Rescates

					int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EncabezadoAsientoDevolucionCobranzaPichincha '{0}', '{1}', 86", 
						Convert.ToDateTime(this.dtpFechaDoc.Value).ToString("yyyyMMdd"), this.txtNumeroDoc.Text.ToString()));

					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows)
					{
						if (Convert.ToBoolean(dr.Cells["Sel"].Value))
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec DetalleAsientoDevolucionCobranzaPichincha {0}, {1}, {2}, '{3}', 2",
								idAsiento, 
								Convert.ToInt32(dr.Cells["idPagoCuotaCredito"].Value),
								Convert.ToDecimal(dr.Cells["Debitado"].Value), 
								(Convert.ToString(dr.Cells["Ruc"].Value) + "-" + Convert.ToString(dr.Cells["Nombre"].Value))));												

							dCapital = dCapital + Convert.ToDecimal(dr.Cells["Capital"].Value);
							dInteres = dInteres + Convert.ToDecimal(dr.Cells["Interes"].Value);
							dMora = dMora + Convert.ToDecimal(dr.Cells["Mora"].Value);						
						}
					}
				
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec BancoAsientoDevolucionCobranzaPichincha {0}", idAsiento));

					string sSQLCO = string.Format("Exec AsientoCuentasDeOrdenEnvioPagos 0, '{0}', {1}, {2}, {3}", 
						this.txtNumeroDoc.Text.ToString(), dCapital, dInteres, dMora);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCO, true);

					string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
				
					Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
					Reporte.MdiParent = this.MdiParent;			
					Reporte.Show();
				}
				#endregion Asiento Pagos FDG

				#region Asiento Ajuste Centavos
				if ((int)this.optTipos.Value == 3)
				{
					#region Liquidacion Saldos y Precancelaciones
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridPagos.Rows.All)
						if (dr.Cells["Debitado"].Value != System.DBNull.Value) dRescates += (decimal) dr.Cells["Debitado"].Value;

					string sSQLAct = string.Format("Exec Cre_InsertaDetalleLiquidacion {0}, 1, {1}, 'SALDOS Y PRECANCELACIONES', {2}",
						idCre_LiquidacionPagosPichincha, (bool)this.optTipoPago.Value, dRescates); 
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
					#endregion Liquidacion Saldos y Precancelaciones

					int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EncabezadoAsientoDevolucionCobranzaPichincha '{0}', '{1}', 86", 
						Convert.ToDateTime(this.dtpFechaDoc.Value).ToString("yyyyMMdd"), this.txtNumeroDoc.Text.ToString()));

					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows)
					{
						if (Convert.ToBoolean(dr.Cells["Sel"].Value))
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec DetalleAsientoDevolucionCobranzaPichincha {0}, {1}, {2}, '{3}', 3",
								idAsiento, 
								Convert.ToInt32(dr.Cells["idPagoCuotaCredito"].Value),
								Convert.ToDecimal(dr.Cells["Debitado"].Value), 
								(Convert.ToString(dr.Cells["Ruc"].Value) + "-" + Convert.ToString(dr.Cells["Nombre"].Value))));												

							dCapital = dCapital + Convert.ToDecimal(dr.Cells["Capital"].Value);
							dInteres = dInteres + Convert.ToDecimal(dr.Cells["Interes"].Value);
							dMora = dMora + Convert.ToDecimal(dr.Cells["Mora"].Value);						
						}
					}
				
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec BancoAsientoDevolucionCobranzaPichincha {0}", idAsiento));
					//@idAsiento Int, @Referencia Varchar(20), @Fecha Date, @Capital Decimal(18, 2), @Interes Decimal(18, 2), @Mora Decimal(18, 2), @Tipo Int

					string sSQLCO = string.Format("Exec AsientoCuentasDeOrdenEnvioPagos 0, '{0}', '{1}', {2}, {3}, {4}, 3", 
						this.txtNumeroDoc.Text.ToString(), Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), dCapital, dInteres, dMora);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCO, true);

					string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
				
					Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
					Reporte.MdiParent = this.MdiParent;			
					Reporte.Show();
				}
				#endregion Asiento Ajuste Centavos

				#region Asiento Retiros De Mercaderia
				if ((int)this.optTipos.Value == 4)
				{
					int idAsiento = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec EncabezadoAsientoDevolucionCobranzaPichincha '{0}', '{1}', 86", 
						Convert.ToDateTime(this.dtpFechaDoc.Value).ToString("yyyyMMdd"), this.txtNumeroDoc.Text.ToString()));

					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows)
					{
						if (Convert.ToBoolean(dr.Cells["Sel"].Value))
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec DetalleAsientoDevolucionCobranzaPichincha {0}, {1}, {2}, '{3}', 2",
								idAsiento, 
								Convert.ToInt32(dr.Cells["idPagoCuotaCredito"].Value),
								Convert.ToDecimal(dr.Cells["Debitado"].Value), 
								(Convert.ToString(dr.Cells["Ruc"].Value) + "-" + Convert.ToString(dr.Cells["Nombre"].Value))));												

							dCapital = dCapital + Convert.ToDecimal(dr.Cells["Capital"].Value);
							dInteres = dInteres + Convert.ToDecimal(dr.Cells["Interes"].Value);
							dMora = dMora + Convert.ToDecimal(dr.Cells["Mora"].Value);
							
				      string sSQL = string.Format("Exec CreaNotaDeDebitoRetiroMercaderia {0}, '{1}', {2}", Convert.ToInt32(dr.Cells["idAnticipo"].Value), DateTime.Today.ToString("yyyyMMdd"),
								Convert.ToDecimal(dr.Cells["Debitado"].Value));
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);			
						}
					}
				
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec BancoAsientoDevolucionCobranzaPichincha {0}", idAsiento));

					string sSQLCO = string.Format("Exec AsientoCuentasDeOrdenEnvioPagos 0, '{0}', {1}, {2}, {3}", 
						this.txtNumeroDoc.Text.ToString(), dCapital, dInteres, dMora);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLCO, true);

					string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
				
					Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);
					Reporte.MdiParent = this.MdiParent;			
					Reporte.Show();
				}
				#endregion Asiento Retiros De Mercaderia
				#endregion Crea Asiento

				Cursor = Cursors.Default;

				MessageBox.Show("Liquidación Generada Correctamente", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
				ContPagos();

				this.txtNumeroDoc.Text = "";			

				this.btnVer_Click(sender, e);				
			}
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;                                  

					//Funcion.EjecutaSQL(cdsSeteoF, "Update PagosCuotasCredito Set Sel = 0 Where EstadoBanco = 2 And idAsiento = 0");

					int iCont = 0;

					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [PAGOS$]", "PAGOS").Tables[0].Rows)
					{
						#region Variables Archivo
						string sReferencia = "";
						string sOperacion = "";
						string sTipoPago = "";
						decimal dDebitado = 0.00m;
						string sTipo = "";
						string sError = "";						

						if (row["REFERENCIA"] != System.DBNull.Value) sReferencia = row["REFERENCIA"].ToString();
						if (row["OPERACION"] != System.DBNull.Value) sOperacion = row["OPERACION"].ToString();
						if (row["TIPO_PAGO"] != System.DBNull.Value) sTipoPago = row["TIPO_PAGO"].ToString();
						if (row["VALOR_APLICADO"] != System.DBNull.Value) dDebitado = decimal.Parse(row["VALOR_APLICADO"].ToString());
						if (row["TIPO"] != System.DBNull.Value) sTipo = row["TIPO"].ToString();
						if (row["ERROR"] != System.DBNull.Value) sError = row["ERROR"].ToString();						
						#endregion Variables Archivo	

						if ((int)this.optTipos.Value == 1)
						{
							string sSQLAct = string.Format("Exec Cre_ActualizaPagosPichincha '{0}', '{1}', '{2}', {3}, '{4}', '{5}'", sReferencia, sOperacion, sTipoPago, dDebitado, sTipo, sError);						
							Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						}

						if ((int)this.optTipos.Value == 2)
						{
							DateTime dtFecha = DateTime.Today;
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("update TEMPORALPICVENCIDAS set sel = 1 where operacion = '{0}' and convert(date, fecha) = '{1}'", sOperacion, dtFecha.ToString("yyyyMMdd")));
						}						

						if ((int)this.optTipos.Value == 3)
						{
							DateTime dtFecha = DateTime.Today;
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("update TEMPORALPICSALDOS set sel = 1 where operacion = '{0}' and convert(date, fechasistema) = '20200831'", sOperacion));//, dtFecha.ToString("yyyyMMdd")));
						}	

						iCont++;
					}				
					#endregion Sube archivo

					MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
		
					this.btnImportar.Enabled = false;
					this.btnVer_Click(sender, e);
					this.Cursor = Cursors.Default;	
				}
			}
			catch (Exception ex)
			{
				try
				{
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
					this.Cursor = Cursors.Default;	
					return;
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					this.Cursor = Cursors.Default;	
					return;
				}						
			}
		}

		private void uGridPagos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	
			if (e.Cell.Column.ToString() == "Sel")
			{
				string sSQL = string.Format("Exec Cre_ActualizaValoresDevolucionCobranza 1, {0}, {1}", 
					(bool)e.Cell.Row.Cells["Sel"].Value, (int)e.Cell.Row.Cells["idPagoCuotaCredito"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			if (e.Cell.Column.ToString() == "EstadoBanco")
			{
				string sSQL = string.Format("Update PagosCuotasCredito Set EstadoBanco = {0}, Enviado = 0 Where idPagoCuotaCredito = {1}", 
					(int)e.Cell.Row.Cells["EstadoBanco"].Value, (int)e.Cell.Row.Cells["idPagoCuotaCredito"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}

			if (e.Cell.Column.ToString() == "Debitado")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Cre_ActualizaDebitadoPagos {0}, {1}", 
					(int)e.Cell.Row.Cells["idPagoCuotaCredito"].Value, Convert.ToDecimal(e.Cell.Row.Cells["Debitado"].Value)));
			}
		}

		private void uGridPagos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Debitado", 2);
		}

		private void uGridPagos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridPagos);
		}

		private void optTipoPago_ValueChanged(object sender, System.EventArgs e)
		{			
			this.btnImportar.Enabled = false;
			this.dtpFechaDoc.Enabled = false;
			this.dtpFechaDoc.Value = System.DBNull.Value;			
			this.txtNumeroDoc.Enabled = false;
			this.txtNumeroDoc.Text = "";
			this.btnContabilizar.Enabled = false;

			this.btnVer_Click(sender, e);

			if ((int)this.optEstado.Value == 2)
			{
//				if (!(bool)this.optTipoPago.Value)
//				{
//					this.btnImportar.Enabled = true;				
//				}
//				else
//				{
					this.dtpFechaDoc.Enabled = true;
					this.dtpFechaDoc.Value = DateTime.Today;
					this.dtpFechaDoc.CalendarInfo.MaxDate = DateTime.Today;
					this.txtNumeroDoc.Enabled = true;
					this.btnContabilizar.Enabled = true;
//				}
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridPagos);
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridPagos.Rows.All)
			{	
				dr.Cells["Sel"].Value = !(bool)dr.Cells["Sel"].Value;
			}
		}

		private void uGridPagos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.uGridPagos.UpdateData();
				
				string sSQL = string.Format("Exec Cre_ActualizaValoresDevolucionCobranza 1, {0}, {1}", 
					(bool)e.Cell.Row.Cells["Sel"].Value, (int)e.Cell.Row.Cells["idPagoCuotaCredito"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{			
			FuncionesContabilidad.ImprimeAsiento((int)this.uGridPagos.ActiveRow.Cells["idAsiento"].Value);
		}

		private void dtpFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void uGridPagos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString().Equals("btnVer"))
			{	
				if (e.Cell.Row.Cells["Ruc"].Value != System.DBNull.Value)
				{
					frmCobranzas CB = new  frmCobranzas(e.Cell.Row.Cells["Ruc"].Value.ToString(), true);
					CB.MdiParent = MdiParent;
					CB.Show();	
				}
			}		
		}

		private void btnImportarCV_Click(object sender, System.EventArgs e)
		{
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;
				
					int iCont = 0;

					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [DATOS$]", "DATOS").Tables[0].Rows)
					{
						#region Variables
						string sCEDULA = "";
						string sOPERACION = "";
						string sCLIENTE= "";
						string sFECHA_VCTO_CUOTA = "";
						decimal dMONTO_FINANCIADO =  0.00m;
						decimal dCAPITAL_AMORTIZACION =  0.00m;
						decimal dINTERES =  0.00m;
						decimal dMORA =  0.00m;
						decimal dVALOR_A_PAGAR =  0.00m;
						string sESTADO = "";
						int iNUMERO_CUOTA = 0;
						int iDIAS_MORA = 0;
						int iNUMERO_CUOTAS_VENCIDAS = 0;												
						DateTime dtFechaVcto = DateTime.Today;
						string sNumCuota = "";
						#endregion Variables				            												
						//OPERACION	IDENTIFICACION	CLIENTE	MONTO_FINANCIADO	SALDO_CAPITAL_VENCIDO	INTERES_X_PAGAR	MORA_X_PAGAR	DIAS_MORA	NUMERO_CUOTAS_VENCIDAS	V_CUOTAS_VENCIDAS	CUOTAS_CANCELADAS	SEGMENTO	CUOTA

						if (row["CEDULA"] != System.DBNull.Value) sCEDULA = row["CEDULA"].ToString();												
						if (row["OPERACION"] != System.DBNull.Value) sOPERACION = row["OPERACION"].ToString();
						if (row["CLIENTE"] != System.DBNull.Value) sCLIENTE = row["CLIENTE"].ToString();
						if (row["FECHA_VCTO_CUOTA"] != System.DBNull.Value) sFECHA_VCTO_CUOTA = row["FECHA_VCTO_CUOTA"].ToString(); //DateTime.ParseExact(row["FECHA_VCTO_CUOTA"].ToString(),"yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy-MM-dd");																	
						if (row["MONTO_FINANCIADO"] != System.DBNull.Value) dMONTO_FINANCIADO = decimal.Parse(row["MONTO_FINANCIADO"].ToString());
						if (row["CAPITAL_AMORTIZACION"] != System.DBNull.Value) dCAPITAL_AMORTIZACION = decimal.Parse(row["CAPITAL_AMORTIZACION"].ToString());
						if (row["INTERES"] != System.DBNull.Value) dINTERES = decimal.Parse(row["INTERES"].ToString());
						if (row["MORA"] != System.DBNull.Value) dMORA = decimal.Parse(row["MORA"].ToString());
						if (row["VALOR_A_PAGAR"] != System.DBNull.Value) dVALOR_A_PAGAR = decimal.Parse(row["VALOR_A_PAGAR"].ToString());
						if (row["ESTADO"] != System.DBNull.Value) sESTADO = row["ESTADO"].ToString();
						if (row["NUMERO_CUOTA"] != System.DBNull.Value) sNumCuota = row["NUMERO_CUOTA"].ToString();
						if (row["DIAS_MORA"] != System.DBNull.Value) iDIAS_MORA = int.Parse(row["DIAS_MORA"].ToString());
						if (row["NUMERO_CUOTAS_VENCIDAS"] != System.DBNull.Value) iNUMERO_CUOTAS_VENCIDAS = int.Parse(row["NUMERO_CUOTAS_VENCIDAS"].ToString());

						if (sCEDULA.Length < 10) sCEDULA = sCEDULA.PadLeft(10, '0');

						dtFechaVcto = DateTime.Parse(sFECHA_VCTO_CUOTA.Substring(0, 4) + "-" + sFECHA_VCTO_CUOTA.Substring(4, 2) + "-" + sFECHA_VCTO_CUOTA.Substring(6, 2));
						//dtFechaVcto = DateTime.Today.AddDays(-iDIAS_MORA);

						iNUMERO_CUOTA = int.Parse(sNumCuota.Replace(",00", ""));

						#region Fondo de garantia
						if ((int)this.optTipos.Value == 2)
						{							
							int iCuoC_CuoV = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select NumCuotas - Cuotas_Vendidas From Compra Where NumeroDeOperacion = '{0}' And idEntidadFinanciera = 4", sOPERACION));

							if (iCuoC_CuoV > 0) iNUMERO_CUOTA = iNUMERO_CUOTA - iCuoC_CuoV;

							int iEstado = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaCuotaPagadaRescate '{0}', {1}", sOPERACION, iNUMERO_CUOTA));

							int iDiasVencidos = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT DATEDIFF(DAY, '{0}', GETDATE())", dtFechaVcto.ToString("yyyyMMdd")));
						
							if (iEstado != 2)
							{
//								if (iDiasVencidos > 98)
//								{
									if (Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSubeCarteraRescatada '{0}'", sOPERACION)) == 0)
									{
										string sSQLAct = string.Format("Exec GrabaRegFondo '{0}', '{1}', '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}", 
											sCEDULA, sOPERACION, sCLIENTE, dtFechaVcto.ToString("yyyyMMdd"), dMONTO_FINANCIADO, dCAPITAL_AMORTIZACION, dINTERES, dMORA, dVALOR_A_PAGAR, iNUMERO_CUOTA, iDIAS_MORA, iNUMERO_CUOTAS_VENCIDAS);						
										Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
						
										iCont++;
									}
//								}
							}
						}
						#endregion Fondo de garantia
					}				
					#endregion Sube archivo

					MessageBox.Show(string.Format("Importación Finalizada Correctamente, {0} Lineas Afectadas", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);		
					this.btnVer_Click(sender, e);
					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{				
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				this.Cursor = Cursors.Default;
				return;										
			}

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if (iEstado == 0 || iEstado == 2)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, "Exec Cre_ValidaPagosPichinchaPorLiquidar") > 0)
				{
					MessageBox.Show("Existen pagos por liquidar de dias anteriores a Hoy ('01-11-2018')\n\nLiquide los pagos antes de generar un nuevo envio", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				idCre_LiquidacionPagosPichincha = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Insert Into Cre_LiquidacionPagosPichincha (Fecha, Estado) Values (GETDATE(), 1) Declare @id Int Set @id = SCOPE_IDENTITY() Select @id"));

				this.btnImportarCV.Enabled = true;
				this.btnGenerar.Enabled = true;
				this.btnImportar.Enabled = true;

				MessageBox.Show("Liquidación Generada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				iEstado = 1;

				this.btnNuevo.Text = "Cerrar";
			}
			else if (iEstado == 1)
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Cre_LiquidacionPagosPichincha Set CerradoPor = SUSER_SNAME(), FechaCierre = GETDATE(), Estado = 2 Where idCre_LiquidacionPagosPichincha = {0}", idCre_LiquidacionPagosPichincha));

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfidAsiento = new ParameterField ();
				ParameterDiscreteValue ValIdAsiento = new ParameterDiscreteValue ();
				pfidAsiento.ParameterFieldName = "@idCre_LiquidacionPagosPichincha";
				ValIdAsiento.Value = idCre_LiquidacionPagosPichincha;
				pfidAsiento.CurrentValues.Add (ValIdAsiento);
				paramFields.Add (pfidAsiento);
							
				Reporte Reporte = new Reporte("Cre_LiquidacionPagosPichincha.rpt", "");
				Reporte.MdiParent = this.MdiParent;
				Reporte.Titulo("Liquidacion");
				Reporte.crVista.ParameterFieldInfo = paramFields;
				Reporte.Show();					

				MessageBox.Show("Liquidación Cerrada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnImportarCV.Enabled = false;
				this.btnGenerar.Enabled = false;
				this.btnImportar.Enabled = false;				
				this.btnNuevo.Text = "Nuevo";

				iEstado = 2;
			}			
		}

		private void optTipos_ValueChanged(object sender, System.EventArgs e)
		{
			this.optEstado.Value = 1;
		}

		private void btnRescates_Click(object sender, System.EventArgs e)
		{
			using (frmAnalisisFondoGarantia FondoGarantia = new frmAnalisisFondoGarantia())
			{
				if (DialogResult.OK == FondoGarantia.ShowDialog())
				{
				}
			}
	
		}
	}
}
