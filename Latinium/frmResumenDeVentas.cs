using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmResumenDeVentas.
	/// </summary>
	public class frmResumenDeVentas : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridMatriz;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMatriz;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridHorizontal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPorLocal;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPorDias;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPorVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPorlocalVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
														//	private Infragistics.Win.UltraWinGrid.UltraGrid uGridRemates;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridAgrupadoPorDias;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgReporte;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridHorizontalGex;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridGExPorLocal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDYD;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkturno1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkturno2;
		private System.ComponentModel.IContainer components;


		public frmResumenDeVentas()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DOC");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PayJoy");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 9, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 10, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 12, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 13, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 14, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 15, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 16, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 17, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 18, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 19, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 20, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 21, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 22, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 23, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 24, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 25, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 26, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 26, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PayJoy", 27, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PayJoy", 27, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DOC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalExcento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PayJoy");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmResumenDeVentas));
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalExcento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PayJoy");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PayJoy");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 2, true);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 3, true);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 4, true);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 5, true);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 6, true);
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 7, true);
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 8, true);
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 10, true);
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 9, true);
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 11, true);
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 12, true);
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 13, true);
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 14, true);
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 16, true);
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 17, true);
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 18, true);
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 19, true);
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 20, true);
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PayJoy", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PayJoy", 21, true);
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PayJoy");
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 2, true);
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 3, true);
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 4, true);
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 5, true);
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 6, true);
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 7, true);
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 8, true);
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 10, true);
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 9, true);
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 11, true);
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 12, true);
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 13, true);
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings54 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 14, true);
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings55 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings56 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 16, true);
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings57 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 17, true);
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings58 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 19, true);
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings59 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 18, true);
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings60 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 20, true);
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings61 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PayJoy", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PayJoy", 21, true);
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalExcento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PayJoy");
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubtotalExcento");
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PayJoy");
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings62 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings63 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 2, true);
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings64 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 3, true);
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings65 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 4, true);
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings66 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 5, true);
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings67 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 6, true);
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings68 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 7, true);
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings69 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 8, true);
			Infragistics.Win.Appearance appearance204 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings70 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 10, true);
			Infragistics.Win.Appearance appearance205 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings71 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 9, true);
			Infragistics.Win.Appearance appearance206 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings72 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 11, true);
			Infragistics.Win.Appearance appearance207 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings73 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 12, true);
			Infragistics.Win.Appearance appearance208 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings74 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 13, true);
			Infragistics.Win.Appearance appearance209 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings75 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 14, true);
			Infragistics.Win.Appearance appearance210 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings76 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance211 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings77 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 16, true);
			Infragistics.Win.Appearance appearance212 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings78 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 17, true);
			Infragistics.Win.Appearance appearance213 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings79 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 18, true);
			Infragistics.Win.Appearance appearance214 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings80 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 19, true);
			Infragistics.Win.Appearance appearance215 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings81 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SubtotalExcento", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SubtotalExcento", 20, true);
			Infragistics.Win.Appearance appearance216 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings82 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PayJoy", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PayJoy", 21, true);
			Infragistics.Win.Appearance appearance217 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance218 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance219 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance220 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance221 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance222 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance223 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance224 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance225 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubtotalExcento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PayJoy");
			Infragistics.Win.Appearance appearance226 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance227 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance228 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance229 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance230 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance231 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance232 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance233 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance234 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance235 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance236 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance237 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance238 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance239 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance240 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance241 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance242 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance243 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance244 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance245 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings83 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 2, true);
			Infragistics.Win.Appearance appearance246 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings84 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 3, true);
			Infragistics.Win.Appearance appearance247 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings85 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 4, true);
			Infragistics.Win.Appearance appearance248 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings86 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 5, true);
			Infragistics.Win.Appearance appearance249 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings87 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 6, true);
			Infragistics.Win.Appearance appearance250 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings88 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 7, true);
			Infragistics.Win.Appearance appearance251 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings89 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 8, true);
			Infragistics.Win.Appearance appearance252 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings90 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 9, true);
			Infragistics.Win.Appearance appearance253 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings91 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 11, true);
			Infragistics.Win.Appearance appearance254 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings92 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 10, true);
			Infragistics.Win.Appearance appearance255 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings93 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 12, true);
			Infragistics.Win.Appearance appearance256 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings94 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 13, true);
			Infragistics.Win.Appearance appearance257 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings95 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 14, true);
			Infragistics.Win.Appearance appearance258 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings96 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 15, true);
			Infragistics.Win.Appearance appearance259 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings97 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 16, true);
			Infragistics.Win.Appearance appearance260 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings98 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 17, true);
			Infragistics.Win.Appearance appearance261 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings99 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 18, true);
			Infragistics.Win.Appearance appearance262 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings100 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 19, true);
			Infragistics.Win.Appearance appearance263 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings101 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 20, true);
			Infragistics.Win.Appearance appearance264 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance265 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance266 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance267 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance268 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance269 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance270 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance271 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance272 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance273 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance274 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance275 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance276 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance277 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance278 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance279 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance280 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance281 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance282 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance283 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance284 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance285 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance286 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance287 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance288 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance289 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance290 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance291 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance292 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings102 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Subtotal", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 2, true);
			Infragistics.Win.Appearance appearance293 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings103 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "IVA", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 3, true);
			Infragistics.Win.Appearance appearance294 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings104 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Total", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 4, true);
			Infragistics.Win.Appearance appearance295 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings105 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Efectivo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 5, true);
			Infragistics.Win.Appearance appearance296 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings106 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Cheque", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 6, true);
			Infragistics.Win.Appearance appearance297 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings107 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Tarjeta", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 7, true);
			Infragistics.Win.Appearance appearance298 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings108 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Nota", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 8, true);
			Infragistics.Win.Appearance appearance299 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings109 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Credito", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 9, true);
			Infragistics.Win.Appearance appearance300 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings110 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Anticipo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 11, true);
			Infragistics.Win.Appearance appearance301 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings111 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Transferencia", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 10, true);
			Infragistics.Win.Appearance appearance302 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings112 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "CrediPoint", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 12, true);
			Infragistics.Win.Appearance appearance303 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings113 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Retencion", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 13, true);
			Infragistics.Win.Appearance appearance304 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings114 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Debito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 14, true);
			Infragistics.Win.Appearance appearance305 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings115 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Deposito", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 15, true);
			Infragistics.Win.Appearance appearance306 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings116 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Bono", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 16, true);
			Infragistics.Win.Appearance appearance307 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings117 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "C_Entrega", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 17, true);
			Infragistics.Win.Appearance appearance308 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings118 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Cupon", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 18, true);
			Infragistics.Win.Appearance appearance309 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings119 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Locales", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 19, true);
			Infragistics.Win.Appearance appearance310 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings120 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Carpa", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 20, true);
			Infragistics.Win.Appearance appearance311 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance312 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance313 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance314 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance315 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance316 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance317 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance318 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance319 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn126 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn127 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn128 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn129 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn130 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn131 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn132 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn133 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn134 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn135 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn136 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn137 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn138 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn139 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn140 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn141 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn142 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance320 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance321 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn143 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn144 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DOC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn145 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn146 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn147 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn148 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn149 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn150 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn151 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn152 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn153 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn154 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn155 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance322 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance323 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DOC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance324 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance325 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance326 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings121 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 10, true);
			Infragistics.Win.Appearance appearance327 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings122 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 11, true);
			Infragistics.Win.Appearance appearance328 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings123 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 12, true);
			Infragistics.Win.Appearance appearance329 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance330 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance331 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance332 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance333 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance334 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance335 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance336 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance337 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance338 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings124 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance339 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance340 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance341 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance342 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance343 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance344 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance345 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance346 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn156 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn157 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.Appearance appearance347 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Seguro");
			Infragistics.Win.Appearance appearance348 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA_Seguro");
			Infragistics.Win.Appearance appearance349 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Seguro_Con_IVA");
			Infragistics.Win.Appearance appearance350 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFactura");
			Infragistics.Win.Appearance appearance351 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto_Credito");
			Infragistics.Win.Appearance appearance352 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings125 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Seguro", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Seguro", 4, true);
			Infragistics.Win.Appearance appearance353 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings126 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA_Seguro", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA_Seguro", 5, true);
			Infragistics.Win.Appearance appearance354 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings127 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Seguro_Con_IVA", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Seguro_Con_IVA", 6, true);
			Infragistics.Win.Appearance appearance355 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings128 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalFactura", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalFactura", 7, true);
			Infragistics.Win.Appearance appearance356 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings129 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto_Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto_Credito", 8, true);
			Infragistics.Win.Appearance appearance357 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance358 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance359 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance360 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance361 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance362 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance363 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance364 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance365 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn158 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn159 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn160 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn161 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn162 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn163 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn164 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro_Con_IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn165 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn166 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Credito");
			Infragistics.Win.Appearance appearance366 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance367 = new Infragistics.Win.Appearance();
			this.uGridHorizontal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridMatriz = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkMatriz = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPorLocal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridPorDias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPorVendedor = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridPorlocalVendedor = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.uGridAgrupadoPorDias = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.rdgReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridHorizontalGex = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridGExPorLocal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridDYD = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkturno1 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkturno2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.uGridHorizontal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMatriz)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorlocalVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAgrupadoPorDias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHorizontalGex)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGExPorLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDYD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridHorizontal
			// 
			this.uGridHorizontal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridHorizontal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridHorizontal.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridHorizontal.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Header.Caption = "Doc";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 100;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Header.VisiblePosition = 4;
			ultraGridColumn9.Width = 40;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance12;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "T/C";
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance14;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "N/C";
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance15;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance16;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance17;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Transf";
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance18;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "C/P";
			ultraGridColumn20.Header.VisiblePosition = 21;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance19;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Ret";
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance20;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "T/D";
			ultraGridColumn22.Header.VisiblePosition = 23;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance21;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 24;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance22;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 25;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance23;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 26;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance24;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 27;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance25;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "Subtotal Excento";
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance26;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 18;
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
																										 ultraGridColumn28});
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance27;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance28;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance29;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance30;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance31;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance32;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance33;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance34;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance35;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance36;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance37;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance38;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance39;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance40;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance41;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance42;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance43;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance44;
			summarySettings18.DisplayFormat = "{0:#,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance45;
			summarySettings19.DisplayFormat = "{0:#,##0.00}";
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
			ultraGridBand1.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridHorizontal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridHorizontal.DisplayLayout.GroupByBox.Hidden = true;
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontal.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.uGridHorizontal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridHorizontal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHorizontal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridHorizontal.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontal.DisplayLayout.Override.CellAppearance = appearance48;
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontal.DisplayLayout.Override.GroupByRowAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 8.5F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHorizontal.DisplayLayout.Override.HeaderAppearance = appearance50;
			this.uGridHorizontal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance51.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance51.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontal.DisplayLayout.Override.RowAlternateAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontal.DisplayLayout.Override.RowSelectorAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontal.DisplayLayout.Override.SelectedRowAppearance = appearance53;
			this.uGridHorizontal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridHorizontal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHorizontal.Location = new System.Drawing.Point(264, 8);
			this.uGridHorizontal.Name = "uGridHorizontal";
			this.uGridHorizontal.Size = new System.Drawing.Size(976, 704);
			this.uGridHorizontal.TabIndex = 182;
			this.uGridHorizontal.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridHorizontal_InitializeRow);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(System.DateTime);
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
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn28});
			// 
			// dtHasta
			// 
			appearance54.FontData.Name = "Tahoma";
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance54;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(56, 168);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 639;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.Validated += new System.EventHandler(this.dtHasta_Validated);
			// 
			// dtDesde
			// 
			appearance55.FontData.Name = "Tahoma";
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance55;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 144);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 638;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.Validated += new System.EventHandler(this.dtDesde_Validated);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(8, 168);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(33, 16);
			this.label17.TabIndex = 641;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 144);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(35, 16);
			this.label18.TabIndex = 640;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridLocales
			// 
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource1;
			appearance56.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance56;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 100;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 199;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.Caption = "...";
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn31.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance57.ForeColor = System.Drawing.Color.Black;
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance57;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance58.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance58;
			appearance59.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance59.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance59.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance59.FontData.BoldAsString = "True";
			appearance59.FontData.Name = "Arial";
			appearance59.FontData.SizeInPoints = 8.5F;
			appearance59.ForeColor = System.Drawing.Color.White;
			appearance59.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance59;
			appearance60.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance60.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance60.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance60;
			appearance61.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance61.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance61.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance61;
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance62.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance62.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance62;
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 200);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(240, 264);
			this.uGridLocales.TabIndex = 646;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(bool);
			ultraDataColumn31.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// uGridMatriz
			// 
			this.uGridMatriz.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMatriz.DataSource = this.ultraDataSource2;
			appearance63.BackColor = System.Drawing.Color.White;
			this.uGridMatriz.DisplayLayout.Appearance = appearance63;
			this.uGridMatriz.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 199;
			ultraGridColumn34.Header.Caption = "...";
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridColumn34.Width = 20;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			ultraGridBand3.Header.Caption = "MATRIZ";
			ultraGridBand3.HeaderVisible = true;
			this.uGridMatriz.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance64.ForeColor = System.Drawing.Color.Black;
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMatriz.DisplayLayout.Override.ActiveRowAppearance = appearance64;
			this.uGridMatriz.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridMatriz.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMatriz.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance65.BackColor = System.Drawing.Color.Transparent;
			this.uGridMatriz.DisplayLayout.Override.CardAreaAppearance = appearance65;
			appearance66.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance66.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance66.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance66.FontData.BoldAsString = "True";
			appearance66.FontData.Name = "Arial";
			appearance66.FontData.SizeInPoints = 8.5F;
			appearance66.ForeColor = System.Drawing.Color.White;
			appearance66.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMatriz.DisplayLayout.Override.HeaderAppearance = appearance66;
			appearance67.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance67.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMatriz.DisplayLayout.Override.RowAlternateAppearance = appearance67;
			appearance68.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance68.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMatriz.DisplayLayout.Override.RowSelectorAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMatriz.DisplayLayout.Override.SelectedRowAppearance = appearance69;
			this.uGridMatriz.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMatriz.Location = new System.Drawing.Point(8, 472);
			this.uGridMatriz.Name = "uGridMatriz";
			this.uGridMatriz.Size = new System.Drawing.Size(240, 176);
			this.uGridMatriz.TabIndex = 647;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(bool);
			ultraDataColumn34.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance70;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 335;
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(8, 656);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(240, 21);
			this.cmbVendedor.TabIndex = 651;
			this.toolTip1.SetToolTip(this.cmbVendedor, "Filtrar por Vendedor");
			this.cmbVendedor.ValueMember = "idPersonal";
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
			// chkTodos
			// 
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance71;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(232, 224);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 13);
			this.chkTodos.TabIndex = 653;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// chkMatriz
			// 
			appearance72.ForeColorDisabled = System.Drawing.Color.Black;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkMatriz.Appearance = appearance72;
			this.chkMatriz.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkMatriz.Enabled = false;
			this.chkMatriz.Location = new System.Drawing.Point(232, 480);
			this.chkMatriz.Name = "chkMatriz";
			this.chkMatriz.Size = new System.Drawing.Size(13, 13);
			this.chkMatriz.TabIndex = 654;
			this.chkMatriz.CheckedChanged += new System.EventHandler(this.chkMatriz_CheckedChanged);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn36.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn45.DataType = typeof(System.Decimal);
			ultraDataColumn45.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn46.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn47.DataType = typeof(System.Decimal);
			ultraDataColumn47.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn48.DataType = typeof(System.Decimal);
			ultraDataColumn48.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn49.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn50.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn51.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn54.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn56});
			// 
			// uGridPorLocal
			// 
			this.uGridPorLocal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorLocal.DataSource = this.ultraDataSource4;
			appearance73.BackColor = System.Drawing.Color.White;
			this.uGridPorLocal.DisplayLayout.Appearance = appearance73;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.Caption = "Almacen";
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Width = 160;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance74;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn39.Width = 80;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance75;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 3;
			ultraGridColumn40.Width = 80;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance76;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 4;
			ultraGridColumn41.Width = 80;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance77;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 5;
			ultraGridColumn42.Width = 80;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance78;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 6;
			ultraGridColumn43.Width = 80;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance79;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "T/C";
			ultraGridColumn44.Header.VisiblePosition = 7;
			ultraGridColumn44.Width = 80;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance80;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "N/C";
			ultraGridColumn45.Header.VisiblePosition = 8;
			ultraGridColumn45.Width = 80;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance81;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 9;
			ultraGridColumn46.Width = 80;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance82;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "Transf";
			ultraGridColumn47.Header.VisiblePosition = 12;
			ultraGridColumn47.Width = 80;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance83;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.VisiblePosition = 11;
			ultraGridColumn48.Width = 80;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance84;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "C/P";
			ultraGridColumn49.Header.VisiblePosition = 13;
			ultraGridColumn49.Width = 80;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance85;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "Ret";
			ultraGridColumn50.Header.VisiblePosition = 14;
			ultraGridColumn50.Width = 80;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance86;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "T/D";
			ultraGridColumn51.Header.VisiblePosition = 15;
			ultraGridColumn51.Width = 80;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance87;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 16;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance88;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 17;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance89;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 18;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance90;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 19;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance91;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 20;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance92;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 21;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance93;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "Subtotal Excento";
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance94;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 10;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn59});
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance95;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance96;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance97;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance98;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance99;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance100;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance101;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance102;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance103;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance104;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance105;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance106;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance107;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance108;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance109;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance110;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance111;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance112;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance113;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance114;
			summarySettings39.DisplayFormat = "{0:#,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance115;
			summarySettings40.DisplayFormat = "{0:#,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23,
																																															summarySettings24,
																																															summarySettings25,
																																															summarySettings26,
																																															summarySettings27,
																																															summarySettings28,
																																															summarySettings29,
																																															summarySettings30,
																																															summarySettings31,
																																															summarySettings32,
																																															summarySettings33,
																																															summarySettings34,
																																															summarySettings35,
																																															summarySettings36,
																																															summarySettings37,
																																															summarySettings38,
																																															summarySettings39,
																																															summarySettings40});
			ultraGridBand5.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance116.ForeColor = System.Drawing.Color.Black;
			appearance116.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorLocal.DisplayLayout.Override.ActiveRowAppearance = appearance116;
			this.uGridPorLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance117.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorLocal.DisplayLayout.Override.CardAreaAppearance = appearance117;
			appearance118.ForeColor = System.Drawing.Color.Black;
			appearance118.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorLocal.DisplayLayout.Override.CellAppearance = appearance118;
			appearance119.ForeColor = System.Drawing.Color.Black;
			appearance119.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorLocal.DisplayLayout.Override.GroupByRowAppearance = appearance119;
			appearance120.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance120.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance120.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance120.FontData.BoldAsString = "True";
			appearance120.FontData.Name = "Arial";
			appearance120.FontData.SizeInPoints = 8.5F;
			appearance120.ForeColor = System.Drawing.Color.White;
			appearance120.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorLocal.DisplayLayout.Override.HeaderAppearance = appearance120;
			this.uGridPorLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance121.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance121.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorLocal.DisplayLayout.Override.RowAlternateAppearance = appearance121;
			appearance122.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance122.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance122.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorLocal.DisplayLayout.Override.RowSelectorAppearance = appearance122;
			appearance123.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance123.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance123.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorLocal.DisplayLayout.Override.SelectedRowAppearance = appearance123;
			this.uGridPorLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorLocal.Location = new System.Drawing.Point(264, 8);
			this.uGridPorLocal.Name = "uGridPorLocal";
			this.uGridPorLocal.Size = new System.Drawing.Size(976, 704);
			this.uGridPorLocal.TabIndex = 655;
			this.uGridPorLocal.Visible = false;
			// 
			// uGridPorDias
			// 
			this.uGridPorDias.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorDias.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorDias.DataSource = this.ultraDataSource5;
			appearance124.BackColor = System.Drawing.Color.White;
			this.uGridPorDias.DisplayLayout.Appearance = appearance124;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn60.Header.VisiblePosition = 0;
			ultraGridColumn60.Width = 74;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance125.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance125;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 2;
			ultraGridColumn61.Width = 80;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance126;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.VisiblePosition = 3;
			ultraGridColumn62.Width = 80;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance127;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 4;
			ultraGridColumn63.Width = 80;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance128.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance128;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 5;
			ultraGridColumn64.Width = 80;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance129.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance129;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.VisiblePosition = 6;
			ultraGridColumn65.Width = 80;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance130;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.Caption = "T/C";
			ultraGridColumn66.Header.VisiblePosition = 7;
			ultraGridColumn66.Width = 80;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance131.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance131;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.Caption = "N/C";
			ultraGridColumn67.Header.VisiblePosition = 8;
			ultraGridColumn67.Width = 80;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance132.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance132;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 9;
			ultraGridColumn68.Width = 80;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance133.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance133;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.Caption = "Transf";
			ultraGridColumn69.Header.VisiblePosition = 12;
			ultraGridColumn69.Width = 80;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance134;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 11;
			ultraGridColumn70.Width = 80;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance135;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.Caption = "C/P";
			ultraGridColumn71.Header.VisiblePosition = 13;
			ultraGridColumn71.Width = 80;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance136.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance136;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.Caption = "Ret";
			ultraGridColumn72.Header.VisiblePosition = 14;
			ultraGridColumn72.Width = 80;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance137.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance137;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.Caption = "T/D";
			ultraGridColumn73.Header.VisiblePosition = 15;
			ultraGridColumn73.Width = 80;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance138;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 16;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance139;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 17;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance140;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.VisiblePosition = 18;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance141;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 19;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance142;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 20;
			ultraGridColumn78.Width = 80;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance143;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 21;
			ultraGridColumn79.Width = 80;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance144;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.Caption = "Subtotal Excento";
			ultraGridColumn80.Header.VisiblePosition = 1;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance145;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 10;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn81});
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance146;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance147;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance148;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance149.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance149;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance150.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance150;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance151.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance151;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance152.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance152;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance153.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance153;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance154.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance154;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance155.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance155;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance156.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance156;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance157.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance157;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance158.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance158;
			summarySettings53.DisplayFormat = "{0: #,##0.00}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance159.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings54.Appearance = appearance159;
			summarySettings54.DisplayFormat = "{0: #,##0.00}";
			summarySettings54.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance160.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings55.Appearance = appearance160;
			summarySettings55.DisplayFormat = "{0: #,##0.00}";
			summarySettings55.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance161.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings56.Appearance = appearance161;
			summarySettings56.DisplayFormat = "{0: #,##0.00}";
			summarySettings56.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance162.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings57.Appearance = appearance162;
			summarySettings57.DisplayFormat = "{0: #,##0.00}";
			summarySettings57.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance163.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings58.Appearance = appearance163;
			summarySettings58.DisplayFormat = "{0: #,##0.00}";
			summarySettings58.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance164.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings59.Appearance = appearance164;
			summarySettings59.DisplayFormat = "{0: #,##0.00}";
			summarySettings59.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance165.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings60.Appearance = appearance165;
			summarySettings60.DisplayFormat = "{0: #,##0.00}";
			summarySettings60.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance166.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings61.Appearance = appearance166;
			summarySettings61.DisplayFormat = "{0: #,##0.00}";
			summarySettings61.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings41,
																																															summarySettings42,
																																															summarySettings43,
																																															summarySettings44,
																																															summarySettings45,
																																															summarySettings46,
																																															summarySettings47,
																																															summarySettings48,
																																															summarySettings49,
																																															summarySettings50,
																																															summarySettings51,
																																															summarySettings52,
																																															summarySettings53,
																																															summarySettings54,
																																															summarySettings55,
																																															summarySettings56,
																																															summarySettings57,
																																															summarySettings58,
																																															summarySettings59,
																																															summarySettings60,
																																															summarySettings61});
			ultraGridBand6.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorDias.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance167.ForeColor = System.Drawing.Color.Black;
			appearance167.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorDias.DisplayLayout.Override.ActiveRowAppearance = appearance167;
			this.uGridPorDias.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorDias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorDias.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance168.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorDias.DisplayLayout.Override.CardAreaAppearance = appearance168;
			appearance169.ForeColor = System.Drawing.Color.Black;
			appearance169.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorDias.DisplayLayout.Override.CellAppearance = appearance169;
			appearance170.ForeColor = System.Drawing.Color.Black;
			appearance170.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorDias.DisplayLayout.Override.GroupByRowAppearance = appearance170;
			appearance171.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance171.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance171.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance171.FontData.BoldAsString = "True";
			appearance171.FontData.Name = "Arial";
			appearance171.FontData.SizeInPoints = 8.5F;
			appearance171.ForeColor = System.Drawing.Color.White;
			appearance171.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorDias.DisplayLayout.Override.HeaderAppearance = appearance171;
			this.uGridPorDias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance172.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance172.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance172.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorDias.DisplayLayout.Override.RowAlternateAppearance = appearance172;
			appearance173.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance173.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance173.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorDias.DisplayLayout.Override.RowSelectorAppearance = appearance173;
			appearance174.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance174.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance174.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorDias.DisplayLayout.Override.SelectedRowAppearance = appearance174;
			this.uGridPorDias.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorDias.Location = new System.Drawing.Point(264, 8);
			this.uGridPorDias.Name = "uGridPorDias";
			this.uGridPorDias.Size = new System.Drawing.Size(976, 704);
			this.uGridPorDias.TabIndex = 656;
			this.uGridPorDias.Visible = false;
			this.uGridPorDias.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPorDias_InitializeLayout);
			// 
			// ultraDataSource5
			// 
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
			ultraDataColumn68.DataType = typeof(System.Decimal);
			ultraDataColumn68.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn69.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn70.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn71.DataType = typeof(System.Decimal);
			ultraDataColumn71.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn72.DataType = typeof(System.Decimal);
			ultraDataColumn72.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn73.DataType = typeof(System.Decimal);
			ultraDataColumn73.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn74.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn75.DataType = typeof(System.Decimal);
			ultraDataColumn75.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn76.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn78});
			// 
			// uGridPorVendedor
			// 
			this.uGridPorVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorVendedor.DataSource = this.ultraDataSource6;
			appearance175.BackColor = System.Drawing.Color.White;
			this.uGridPorVendedor.DisplayLayout.Appearance = appearance175;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn82.Header.Caption = "Vendedor";
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Width = 240;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance176.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance176;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.VisiblePosition = 2;
			ultraGridColumn83.Width = 80;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance177.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance177;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.VisiblePosition = 3;
			ultraGridColumn84.Width = 80;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance178.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance178;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.VisiblePosition = 4;
			ultraGridColumn85.Width = 80;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance179.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn86.CellAppearance = appearance179;
			ultraGridColumn86.Format = "#,##0.00";
			ultraGridColumn86.Header.VisiblePosition = 5;
			ultraGridColumn86.Width = 80;
			ultraGridColumn87.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance180.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance180;
			ultraGridColumn87.Format = "#,##0.00";
			ultraGridColumn87.Header.VisiblePosition = 6;
			ultraGridColumn87.Width = 80;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance181.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance181;
			ultraGridColumn88.Format = "#,##0.00";
			ultraGridColumn88.Header.Caption = "T/C";
			ultraGridColumn88.Header.VisiblePosition = 7;
			ultraGridColumn88.Width = 80;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance182.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance182;
			ultraGridColumn89.Format = "#,##0.00";
			ultraGridColumn89.Header.Caption = "N/C";
			ultraGridColumn89.Header.VisiblePosition = 8;
			ultraGridColumn89.Width = 80;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance183.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn90.CellAppearance = appearance183;
			ultraGridColumn90.Format = "#,##0.00";
			ultraGridColumn90.Header.VisiblePosition = 9;
			ultraGridColumn90.Width = 80;
			ultraGridColumn91.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance184.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn91.CellAppearance = appearance184;
			ultraGridColumn91.Format = "#,##0.00";
			ultraGridColumn91.Header.Caption = "Transf";
			ultraGridColumn91.Header.VisiblePosition = 12;
			ultraGridColumn91.Width = 80;
			ultraGridColumn92.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance185.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn92.CellAppearance = appearance185;
			ultraGridColumn92.Format = "#,##0.00";
			ultraGridColumn92.Header.VisiblePosition = 11;
			ultraGridColumn92.Width = 80;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance186.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance186;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.Caption = "C/P";
			ultraGridColumn93.Header.VisiblePosition = 13;
			ultraGridColumn93.Width = 80;
			ultraGridColumn94.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance187.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance187;
			ultraGridColumn94.Format = "#,##0.00";
			ultraGridColumn94.Header.Caption = "Ret";
			ultraGridColumn94.Header.VisiblePosition = 14;
			ultraGridColumn94.Width = 80;
			ultraGridColumn95.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance188.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance188;
			ultraGridColumn95.Format = "#,##0.00";
			ultraGridColumn95.Header.Caption = "T/D";
			ultraGridColumn95.Header.VisiblePosition = 15;
			ultraGridColumn95.Width = 80;
			ultraGridColumn96.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance189.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn96.CellAppearance = appearance189;
			ultraGridColumn96.Format = "#,##0.00";
			ultraGridColumn96.Header.VisiblePosition = 16;
			ultraGridColumn97.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance190.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn97.CellAppearance = appearance190;
			ultraGridColumn97.Format = "#,##0.00";
			ultraGridColumn97.Header.VisiblePosition = 17;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance191.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance191;
			ultraGridColumn98.Format = "#,##0.00";
			ultraGridColumn98.Header.VisiblePosition = 18;
			ultraGridColumn99.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance192.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn99.CellAppearance = appearance192;
			ultraGridColumn99.Format = "#,##0.00";
			ultraGridColumn99.Header.VisiblePosition = 19;
			ultraGridColumn100.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance193.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance193;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 20;
			ultraGridColumn100.Width = 80;
			ultraGridColumn101.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance194.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance194;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 21;
			ultraGridColumn101.Width = 80;
			ultraGridColumn102.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance195.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance195;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridColumn103.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance196.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn103.CellAppearance = appearance196;
			ultraGridColumn103.Format = "#,##0.00";
			ultraGridColumn103.Header.VisiblePosition = 10;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103});
			appearance197.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings62.Appearance = appearance197;
			summarySettings62.DisplayFormat = "{0: #,##0.00}";
			summarySettings62.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance198.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings63.Appearance = appearance198;
			summarySettings63.DisplayFormat = "{0: #,##0.00}";
			summarySettings63.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance199.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings64.Appearance = appearance199;
			summarySettings64.DisplayFormat = "{0: #,##0.00}";
			summarySettings64.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance200.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings65.Appearance = appearance200;
			summarySettings65.DisplayFormat = "{0: #,##0.00}";
			summarySettings65.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance201.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings66.Appearance = appearance201;
			summarySettings66.DisplayFormat = "{0: #,##0.00}";
			summarySettings66.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance202.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings67.Appearance = appearance202;
			summarySettings67.DisplayFormat = "{0: #,##0.00}";
			summarySettings67.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance203.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings68.Appearance = appearance203;
			summarySettings68.DisplayFormat = "{0: #,##0.00}";
			summarySettings68.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance204.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings69.Appearance = appearance204;
			summarySettings69.DisplayFormat = "{0: #,##0.00}";
			summarySettings69.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance205.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings70.Appearance = appearance205;
			summarySettings70.DisplayFormat = "{0: #,##0.00}";
			summarySettings70.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance206.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings71.Appearance = appearance206;
			summarySettings71.DisplayFormat = "{0: #,##0.00}";
			summarySettings71.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance207.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings72.Appearance = appearance207;
			summarySettings72.DisplayFormat = "{0: #,##0.00}";
			summarySettings72.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance208.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings73.Appearance = appearance208;
			summarySettings73.DisplayFormat = "{0: #,##0.00}";
			summarySettings73.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance209.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings74.Appearance = appearance209;
			summarySettings74.DisplayFormat = "{0: #,##0.00}";
			summarySettings74.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance210.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings75.Appearance = appearance210;
			summarySettings75.DisplayFormat = "{0: #,##0.00}";
			summarySettings75.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance211.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings76.Appearance = appearance211;
			summarySettings76.DisplayFormat = "{0: #,##0.00}";
			summarySettings76.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance212.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings77.Appearance = appearance212;
			summarySettings77.DisplayFormat = "{0: #,##0.00}";
			summarySettings77.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance213.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings78.Appearance = appearance213;
			summarySettings78.DisplayFormat = "{0: #,##0.00}";
			summarySettings78.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance214.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings79.Appearance = appearance214;
			summarySettings79.DisplayFormat = "{0: #,##0.00}";
			summarySettings79.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance215.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings80.Appearance = appearance215;
			summarySettings80.DisplayFormat = "{0: #,##0.00}";
			summarySettings80.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance216.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings81.Appearance = appearance216;
			summarySettings81.DisplayFormat = "{0: #,##0.00}";
			summarySettings81.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance217.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings82.Appearance = appearance217;
			summarySettings82.DisplayFormat = "{0: #,##0.00}";
			summarySettings82.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand7.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings62,
																																															summarySettings63,
																																															summarySettings64,
																																															summarySettings65,
																																															summarySettings66,
																																															summarySettings67,
																																															summarySettings68,
																																															summarySettings69,
																																															summarySettings70,
																																															summarySettings71,
																																															summarySettings72,
																																															summarySettings73,
																																															summarySettings74,
																																															summarySettings75,
																																															summarySettings76,
																																															summarySettings77,
																																															summarySettings78,
																																															summarySettings79,
																																															summarySettings80,
																																															summarySettings81,
																																															summarySettings82});
			ultraGridBand7.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance218.ForeColor = System.Drawing.Color.Black;
			appearance218.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorVendedor.DisplayLayout.Override.ActiveRowAppearance = appearance218;
			this.uGridPorVendedor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorVendedor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorVendedor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance219.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorVendedor.DisplayLayout.Override.CardAreaAppearance = appearance219;
			appearance220.ForeColor = System.Drawing.Color.Black;
			appearance220.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorVendedor.DisplayLayout.Override.CellAppearance = appearance220;
			appearance221.ForeColor = System.Drawing.Color.Black;
			appearance221.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorVendedor.DisplayLayout.Override.GroupByRowAppearance = appearance221;
			appearance222.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance222.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance222.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance222.FontData.BoldAsString = "True";
			appearance222.FontData.Name = "Arial";
			appearance222.FontData.SizeInPoints = 8.5F;
			appearance222.ForeColor = System.Drawing.Color.White;
			appearance222.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorVendedor.DisplayLayout.Override.HeaderAppearance = appearance222;
			this.uGridPorVendedor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance223.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance223.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance223.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorVendedor.DisplayLayout.Override.RowAlternateAppearance = appearance223;
			appearance224.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance224.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance224.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorVendedor.DisplayLayout.Override.RowSelectorAppearance = appearance224;
			appearance225.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance225.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance225.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorVendedor.DisplayLayout.Override.SelectedRowAppearance = appearance225;
			this.uGridPorVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorVendedor.Location = new System.Drawing.Point(264, 8);
			this.uGridPorVendedor.Name = "uGridPorVendedor";
			this.uGridPorVendedor.Size = new System.Drawing.Size(976, 704);
			this.uGridPorVendedor.TabIndex = 657;
			this.uGridPorVendedor.Visible = false;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn80.DataType = typeof(System.Decimal);
			ultraDataColumn80.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn81.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn82.DataType = typeof(System.Decimal);
			ultraDataColumn82.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn83.DataType = typeof(System.Decimal);
			ultraDataColumn83.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn84.DataType = typeof(System.Decimal);
			ultraDataColumn84.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn85.DataType = typeof(System.Decimal);
			ultraDataColumn85.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn86.DataType = typeof(System.Decimal);
			ultraDataColumn86.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn87.DataType = typeof(System.Decimal);
			ultraDataColumn87.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn88.DataType = typeof(System.Decimal);
			ultraDataColumn88.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn89.DataType = typeof(System.Decimal);
			ultraDataColumn89.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn90.DataType = typeof(System.Decimal);
			ultraDataColumn90.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn91.DataType = typeof(System.Decimal);
			ultraDataColumn91.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn92.DataType = typeof(System.Decimal);
			ultraDataColumn92.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn93.DataType = typeof(System.Decimal);
			ultraDataColumn93.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn94.DataType = typeof(System.Decimal);
			ultraDataColumn94.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn95.DataType = typeof(System.Decimal);
			ultraDataColumn95.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn96.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn97.DataType = typeof(System.Decimal);
			ultraDataColumn97.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn98.DataType = typeof(System.Decimal);
			ultraDataColumn98.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn99.DataType = typeof(System.Decimal);
			ultraDataColumn100.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92,
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97,
																																 ultraDataColumn98,
																																 ultraDataColumn99,
																																 ultraDataColumn100});
			// 
			// uGridPorlocalVendedor
			// 
			this.uGridPorlocalVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorlocalVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorlocalVendedor.DataSource = this.ultraDataSource7;
			appearance226.BackColor = System.Drawing.Color.White;
			this.uGridPorlocalVendedor.DisplayLayout.Appearance = appearance226;
			ultraGridColumn104.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn104.Header.VisiblePosition = 0;
			ultraGridColumn104.Width = 49;
			ultraGridColumn105.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn105.Header.VisiblePosition = 1;
			ultraGridColumn105.Width = 200;
			ultraGridColumn106.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance227.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance227;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.VisiblePosition = 2;
			ultraGridColumn106.Width = 80;
			ultraGridColumn107.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance228.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance228;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 3;
			ultraGridColumn107.Width = 80;
			ultraGridColumn108.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance229.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn108.CellAppearance = appearance229;
			ultraGridColumn108.Format = "#,##0.00";
			ultraGridColumn108.Header.VisiblePosition = 4;
			ultraGridColumn108.Width = 80;
			ultraGridColumn109.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance230.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn109.CellAppearance = appearance230;
			ultraGridColumn109.Format = "#,##0.00";
			ultraGridColumn109.Header.VisiblePosition = 5;
			ultraGridColumn109.Width = 80;
			ultraGridColumn110.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance231.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn110.CellAppearance = appearance231;
			ultraGridColumn110.Format = "#,##0.00";
			ultraGridColumn110.Header.VisiblePosition = 6;
			ultraGridColumn110.Width = 80;
			ultraGridColumn111.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance232.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn111.CellAppearance = appearance232;
			ultraGridColumn111.Format = "#,##0.00";
			ultraGridColumn111.Header.Caption = "T/C";
			ultraGridColumn111.Header.VisiblePosition = 7;
			ultraGridColumn111.Width = 80;
			ultraGridColumn112.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance233.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn112.CellAppearance = appearance233;
			ultraGridColumn112.Format = "#,##0.00";
			ultraGridColumn112.Header.Caption = "N/C";
			ultraGridColumn112.Header.VisiblePosition = 8;
			ultraGridColumn112.Width = 80;
			ultraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance234.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn113.CellAppearance = appearance234;
			ultraGridColumn113.Format = "#,##0.00";
			ultraGridColumn113.Header.VisiblePosition = 9;
			ultraGridColumn113.Width = 80;
			ultraGridColumn114.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance235.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn114.CellAppearance = appearance235;
			ultraGridColumn114.Format = "#,##0.00";
			ultraGridColumn114.Header.Caption = "Transf";
			ultraGridColumn114.Header.VisiblePosition = 11;
			ultraGridColumn114.Width = 80;
			ultraGridColumn115.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance236.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn115.CellAppearance = appearance236;
			ultraGridColumn115.Format = "#,##0.00";
			ultraGridColumn115.Header.VisiblePosition = 10;
			ultraGridColumn115.Width = 80;
			ultraGridColumn116.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance237.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance237;
			ultraGridColumn116.Format = "#,##0.00";
			ultraGridColumn116.Header.Caption = "C/P";
			ultraGridColumn116.Header.VisiblePosition = 12;
			ultraGridColumn116.Width = 80;
			ultraGridColumn117.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance238.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn117.CellAppearance = appearance238;
			ultraGridColumn117.Format = "#,##0.00";
			ultraGridColumn117.Header.Caption = "Ret";
			ultraGridColumn117.Header.VisiblePosition = 13;
			ultraGridColumn117.Width = 80;
			ultraGridColumn118.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance239.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn118.CellAppearance = appearance239;
			ultraGridColumn118.Format = "#,##0.00";
			ultraGridColumn118.Header.Caption = "T/D";
			ultraGridColumn118.Header.VisiblePosition = 14;
			ultraGridColumn118.Width = 80;
			ultraGridColumn119.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance240.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn119.CellAppearance = appearance240;
			ultraGridColumn119.Format = "#,##0.00";
			ultraGridColumn119.Header.VisiblePosition = 15;
			ultraGridColumn120.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance241.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn120.CellAppearance = appearance241;
			ultraGridColumn120.Format = "#,##0.00";
			ultraGridColumn120.Header.VisiblePosition = 16;
			ultraGridColumn121.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance242.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn121.CellAppearance = appearance242;
			ultraGridColumn121.Format = "#,##0.00";
			ultraGridColumn121.Header.VisiblePosition = 17;
			ultraGridColumn122.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance243.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn122.CellAppearance = appearance243;
			ultraGridColumn122.Format = "#,##0.00";
			ultraGridColumn122.Header.VisiblePosition = 18;
			ultraGridColumn123.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance244.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn123.CellAppearance = appearance244;
			ultraGridColumn123.Format = "#,##0.00";
			ultraGridColumn123.Header.VisiblePosition = 19;
			ultraGridColumn123.Width = 80;
			ultraGridColumn124.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance245.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn124.CellAppearance = appearance245;
			ultraGridColumn124.Format = "#,##0.00";
			ultraGridColumn124.Header.VisiblePosition = 20;
			ultraGridColumn124.Width = 80;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn104,
																										 ultraGridColumn105,
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108,
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124});
			appearance246.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings83.Appearance = appearance246;
			summarySettings83.DisplayFormat = "{0: #,##0.00}";
			summarySettings83.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance247.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings84.Appearance = appearance247;
			summarySettings84.DisplayFormat = "{0: #,##0.00}";
			summarySettings84.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance248.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings85.Appearance = appearance248;
			summarySettings85.DisplayFormat = "{0: #,##0.00}";
			summarySettings85.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance249.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings86.Appearance = appearance249;
			summarySettings86.DisplayFormat = "{0: #,##0.00}";
			summarySettings86.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance250.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings87.Appearance = appearance250;
			summarySettings87.DisplayFormat = "{0: #,##0.00}";
			summarySettings87.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance251.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings88.Appearance = appearance251;
			summarySettings88.DisplayFormat = "{0: #,##0.00}";
			summarySettings88.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance252.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings89.Appearance = appearance252;
			summarySettings89.DisplayFormat = "{0: #,##0.00}";
			summarySettings89.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance253.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings90.Appearance = appearance253;
			summarySettings90.DisplayFormat = "{0: #,##0.00}";
			summarySettings90.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance254.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings91.Appearance = appearance254;
			summarySettings91.DisplayFormat = "{0: #,##0.00}";
			summarySettings91.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance255.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings92.Appearance = appearance255;
			summarySettings92.DisplayFormat = "{0: #,##0.00}";
			summarySettings92.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance256.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings93.Appearance = appearance256;
			summarySettings93.DisplayFormat = "{0: #,##0.00}";
			summarySettings93.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance257.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings94.Appearance = appearance257;
			summarySettings94.DisplayFormat = "{0: #,##0.00}";
			summarySettings94.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance258.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings95.Appearance = appearance258;
			summarySettings95.DisplayFormat = "{0: #,##0.00}";
			summarySettings95.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance259.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings96.Appearance = appearance259;
			summarySettings96.DisplayFormat = "{0: #,##0.00}";
			summarySettings96.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance260.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings97.Appearance = appearance260;
			summarySettings97.DisplayFormat = "{0: #,##0.00}";
			summarySettings97.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance261.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings98.Appearance = appearance261;
			summarySettings98.DisplayFormat = "{0: #,##0.00}";
			summarySettings98.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance262.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings99.Appearance = appearance262;
			summarySettings99.DisplayFormat = "{0: #,##0.00}";
			summarySettings99.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance263.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings100.Appearance = appearance263;
			summarySettings100.DisplayFormat = "{0: #,##0.00}";
			summarySettings100.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance264.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings101.Appearance = appearance264;
			summarySettings101.DisplayFormat = "{0: #,##0.00}";
			summarySettings101.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand8.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings83,
																																															summarySettings84,
																																															summarySettings85,
																																															summarySettings86,
																																															summarySettings87,
																																															summarySettings88,
																																															summarySettings89,
																																															summarySettings90,
																																															summarySettings91,
																																															summarySettings92,
																																															summarySettings93,
																																															summarySettings94,
																																															summarySettings95,
																																															summarySettings96,
																																															summarySettings97,
																																															summarySettings98,
																																															summarySettings99,
																																															summarySettings100,
																																															summarySettings101});
			ultraGridBand8.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorlocalVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.uGridPorlocalVendedor.DisplayLayout.GroupByBox.Hidden = true;
			appearance265.ForeColor = System.Drawing.Color.Black;
			appearance265.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorlocalVendedor.DisplayLayout.Override.ActiveRowAppearance = appearance265;
			this.uGridPorlocalVendedor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorlocalVendedor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorlocalVendedor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance266.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorlocalVendedor.DisplayLayout.Override.CardAreaAppearance = appearance266;
			appearance267.ForeColor = System.Drawing.Color.Black;
			appearance267.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorlocalVendedor.DisplayLayout.Override.CellAppearance = appearance267;
			appearance268.ForeColor = System.Drawing.Color.Black;
			appearance268.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorlocalVendedor.DisplayLayout.Override.GroupByRowAppearance = appearance268;
			appearance269.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance269.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance269.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance269.FontData.BoldAsString = "True";
			appearance269.FontData.Name = "Arial";
			appearance269.FontData.SizeInPoints = 8.5F;
			appearance269.ForeColor = System.Drawing.Color.White;
			appearance269.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorlocalVendedor.DisplayLayout.Override.HeaderAppearance = appearance269;
			this.uGridPorlocalVendedor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance270.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance270.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance270.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorlocalVendedor.DisplayLayout.Override.RowAlternateAppearance = appearance270;
			appearance271.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance271.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance271.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorlocalVendedor.DisplayLayout.Override.RowSelectorAppearance = appearance271;
			appearance272.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance272.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance272.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorlocalVendedor.DisplayLayout.Override.SelectedRowAppearance = appearance272;
			this.uGridPorlocalVendedor.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridPorlocalVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorlocalVendedor.Location = new System.Drawing.Point(264, 8);
			this.uGridPorlocalVendedor.Name = "uGridPorlocalVendedor";
			this.uGridPorlocalVendedor.Size = new System.Drawing.Size(976, 704);
			this.uGridPorlocalVendedor.TabIndex = 658;
			this.uGridPorlocalVendedor.Visible = false;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn103.DataType = typeof(System.Decimal);
			ultraDataColumn103.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn104.DataType = typeof(System.Decimal);
			ultraDataColumn104.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn105.DataType = typeof(System.Decimal);
			ultraDataColumn105.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn106.DataType = typeof(System.Decimal);
			ultraDataColumn106.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn107.DataType = typeof(System.Decimal);
			ultraDataColumn107.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn108.DataType = typeof(System.Decimal);
			ultraDataColumn108.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn109.DataType = typeof(System.Decimal);
			ultraDataColumn109.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn110.DataType = typeof(System.Decimal);
			ultraDataColumn110.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn111.DataType = typeof(System.Decimal);
			ultraDataColumn111.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn112.DataType = typeof(System.Decimal);
			ultraDataColumn112.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn113.DataType = typeof(System.Decimal);
			ultraDataColumn113.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn114.DataType = typeof(System.Decimal);
			ultraDataColumn114.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn115.DataType = typeof(System.Decimal);
			ultraDataColumn115.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn116.DataType = typeof(System.Decimal);
			ultraDataColumn116.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn117.DataType = typeof(System.Decimal);
			ultraDataColumn117.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn118.DataType = typeof(System.Decimal);
			ultraDataColumn118.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn119.DataType = typeof(System.Decimal);
			ultraDataColumn119.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn120.DataType = typeof(System.Decimal);
			ultraDataColumn120.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn121.DataType = typeof(System.Decimal);
			ultraDataColumn121.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn101,
																																 ultraDataColumn102,
																																 ultraDataColumn103,
																																 ultraDataColumn104,
																																 ultraDataColumn105,
																																 ultraDataColumn106,
																																 ultraDataColumn107,
																																 ultraDataColumn108,
																																 ultraDataColumn109,
																																 ultraDataColumn110,
																																 ultraDataColumn111,
																																 ultraDataColumn112,
																																 ultraDataColumn113,
																																 ultraDataColumn114,
																																 ultraDataColumn115,
																																 ultraDataColumn116,
																																 ultraDataColumn117,
																																 ultraDataColumn118,
																																 ultraDataColumn119,
																																 ultraDataColumn120,
																																 ultraDataColumn121});
			// 
			// uGridAgrupadoPorDias
			// 
			this.uGridAgrupadoPorDias.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAgrupadoPorDias.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAgrupadoPorDias.DataSource = this.ultraDataSource8;
			appearance273.BackColor = System.Drawing.Color.White;
			this.uGridAgrupadoPorDias.DisplayLayout.Appearance = appearance273;
			ultraGridColumn125.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn125.Header.VisiblePosition = 0;
			ultraGridColumn126.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn126.Header.VisiblePosition = 1;
			ultraGridColumn126.Width = 74;
			ultraGridColumn127.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance274.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn127.CellAppearance = appearance274;
			ultraGridColumn127.Format = "#,##0.00";
			ultraGridColumn127.Header.VisiblePosition = 2;
			ultraGridColumn127.Width = 80;
			ultraGridColumn128.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance275.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn128.CellAppearance = appearance275;
			ultraGridColumn128.Format = "#,##0.00";
			ultraGridColumn128.Header.VisiblePosition = 3;
			ultraGridColumn128.Width = 80;
			ultraGridColumn129.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance276.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn129.CellAppearance = appearance276;
			ultraGridColumn129.Format = "#,##0.00";
			ultraGridColumn129.Header.VisiblePosition = 4;
			ultraGridColumn129.Width = 80;
			ultraGridColumn130.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance277.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn130.CellAppearance = appearance277;
			ultraGridColumn130.Format = "#,##0.00";
			ultraGridColumn130.Header.VisiblePosition = 5;
			ultraGridColumn130.Width = 80;
			ultraGridColumn131.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance278.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn131.CellAppearance = appearance278;
			ultraGridColumn131.Format = "#,##0.00";
			ultraGridColumn131.Header.VisiblePosition = 6;
			ultraGridColumn131.Width = 80;
			ultraGridColumn132.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance279.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn132.CellAppearance = appearance279;
			ultraGridColumn132.Format = "#,##0.00";
			ultraGridColumn132.Header.Caption = "T/C";
			ultraGridColumn132.Header.VisiblePosition = 7;
			ultraGridColumn132.Width = 80;
			ultraGridColumn133.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance280.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn133.CellAppearance = appearance280;
			ultraGridColumn133.Format = "#,##0.00";
			ultraGridColumn133.Header.Caption = "N/C";
			ultraGridColumn133.Header.VisiblePosition = 8;
			ultraGridColumn133.Width = 80;
			ultraGridColumn134.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance281.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn134.CellAppearance = appearance281;
			ultraGridColumn134.Format = "#,##0.00";
			ultraGridColumn134.Header.VisiblePosition = 9;
			ultraGridColumn134.Width = 80;
			ultraGridColumn135.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance282.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn135.CellAppearance = appearance282;
			ultraGridColumn135.Format = "#,##0.00";
			ultraGridColumn135.Header.Caption = "Transf";
			ultraGridColumn135.Header.VisiblePosition = 11;
			ultraGridColumn135.Width = 80;
			ultraGridColumn136.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance283.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn136.CellAppearance = appearance283;
			ultraGridColumn136.Format = "#,##0.00";
			ultraGridColumn136.Header.VisiblePosition = 10;
			ultraGridColumn136.Width = 80;
			ultraGridColumn137.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance284.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn137.CellAppearance = appearance284;
			ultraGridColumn137.Format = "#,##0.00";
			ultraGridColumn137.Header.Caption = "C/P";
			ultraGridColumn137.Header.VisiblePosition = 12;
			ultraGridColumn137.Width = 80;
			ultraGridColumn138.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance285.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn138.CellAppearance = appearance285;
			ultraGridColumn138.Format = "#,##0.00";
			ultraGridColumn138.Header.Caption = "Ret";
			ultraGridColumn138.Header.VisiblePosition = 13;
			ultraGridColumn138.Width = 80;
			ultraGridColumn139.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance286.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn139.CellAppearance = appearance286;
			ultraGridColumn139.Format = "#,##0.00";
			ultraGridColumn139.Header.Caption = "T/D";
			ultraGridColumn139.Header.VisiblePosition = 14;
			ultraGridColumn139.Width = 80;
			ultraGridColumn140.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance287.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn140.CellAppearance = appearance287;
			ultraGridColumn140.Format = "#,##0.00";
			ultraGridColumn140.Header.VisiblePosition = 15;
			ultraGridColumn141.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance288.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn141.CellAppearance = appearance288;
			ultraGridColumn141.Format = "#,##0.00";
			ultraGridColumn141.Header.VisiblePosition = 16;
			ultraGridColumn142.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance289.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn142.CellAppearance = appearance289;
			ultraGridColumn142.Format = "#,##0.00";
			ultraGridColumn142.Header.VisiblePosition = 17;
			ultraGridColumn143.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance290.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn143.CellAppearance = appearance290;
			ultraGridColumn143.Format = "#,##0.00";
			ultraGridColumn143.Header.VisiblePosition = 18;
			ultraGridColumn144.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance291.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn144.CellAppearance = appearance291;
			ultraGridColumn144.Format = "#,##0.00";
			ultraGridColumn144.Header.VisiblePosition = 19;
			ultraGridColumn144.Width = 80;
			ultraGridColumn145.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance292.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn145.CellAppearance = appearance292;
			ultraGridColumn145.Format = "#,##0.00";
			ultraGridColumn145.Header.VisiblePosition = 20;
			ultraGridColumn145.Width = 80;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145});
			appearance293.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings102.Appearance = appearance293;
			summarySettings102.DisplayFormat = "{0: #,##0.00}";
			summarySettings102.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance294.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings103.Appearance = appearance294;
			summarySettings103.DisplayFormat = "{0: #,##0.00}";
			summarySettings103.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance295.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings104.Appearance = appearance295;
			summarySettings104.DisplayFormat = "{0: #,##0.00}";
			summarySettings104.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance296.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings105.Appearance = appearance296;
			summarySettings105.DisplayFormat = "{0: #,##0.00}";
			summarySettings105.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance297.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings106.Appearance = appearance297;
			summarySettings106.DisplayFormat = "{0: #,##0.00}";
			summarySettings106.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance298.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings107.Appearance = appearance298;
			summarySettings107.DisplayFormat = "{0: #,##0.00}";
			summarySettings107.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance299.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings108.Appearance = appearance299;
			summarySettings108.DisplayFormat = "{0: #,##0.00}";
			summarySettings108.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance300.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings109.Appearance = appearance300;
			summarySettings109.DisplayFormat = "{0: #,##0.00}";
			summarySettings109.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance301.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings110.Appearance = appearance301;
			summarySettings110.DisplayFormat = "{0: #,##0.00}";
			summarySettings110.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance302.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings111.Appearance = appearance302;
			summarySettings111.DisplayFormat = "{0: #,##0.00}";
			summarySettings111.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance303.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings112.Appearance = appearance303;
			summarySettings112.DisplayFormat = "{0: #,##0.00}";
			summarySettings112.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance304.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings113.Appearance = appearance304;
			summarySettings113.DisplayFormat = "{0: #,##0.00}";
			summarySettings113.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance305.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings114.Appearance = appearance305;
			summarySettings114.DisplayFormat = "{0: #,##0.00}";
			summarySettings114.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance306.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings115.Appearance = appearance306;
			summarySettings115.DisplayFormat = "{0: #,##0.00}";
			summarySettings115.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance307.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings116.Appearance = appearance307;
			summarySettings116.DisplayFormat = "{0: #,##0.00}";
			summarySettings116.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance308.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings117.Appearance = appearance308;
			summarySettings117.DisplayFormat = "{0: #,##0.00}";
			summarySettings117.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance309.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings118.Appearance = appearance309;
			summarySettings118.DisplayFormat = "{0: #,##0.00}";
			summarySettings118.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance310.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings119.Appearance = appearance310;
			summarySettings119.DisplayFormat = "{0: #,##0.00}";
			summarySettings119.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance311.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings120.Appearance = appearance311;
			summarySettings120.DisplayFormat = "{0: #,##0.00}";
			summarySettings120.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand9.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings102,
																																															summarySettings103,
																																															summarySettings104,
																																															summarySettings105,
																																															summarySettings106,
																																															summarySettings107,
																																															summarySettings108,
																																															summarySettings109,
																																															summarySettings110,
																																															summarySettings111,
																																															summarySettings112,
																																															summarySettings113,
																																															summarySettings114,
																																															summarySettings115,
																																															summarySettings116,
																																															summarySettings117,
																																															summarySettings118,
																																															summarySettings119,
																																															summarySettings120});
			ultraGridBand9.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridAgrupadoPorDias.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.uGridAgrupadoPorDias.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridAgrupadoPorDias.DisplayLayout.GroupByBox.Prompt = " ";
			appearance312.ForeColor = System.Drawing.Color.Black;
			appearance312.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.ActiveRowAppearance = appearance312;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance313.BackColor = System.Drawing.Color.Transparent;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.CardAreaAppearance = appearance313;
			appearance314.ForeColor = System.Drawing.Color.Black;
			appearance314.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.CellAppearance = appearance314;
			appearance315.ForeColor = System.Drawing.Color.Black;
			appearance315.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.GroupByRowAppearance = appearance315;
			appearance316.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance316.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance316.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance316.FontData.BoldAsString = "True";
			appearance316.FontData.Name = "Arial";
			appearance316.FontData.SizeInPoints = 8.5F;
			appearance316.ForeColor = System.Drawing.Color.White;
			appearance316.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.HeaderAppearance = appearance316;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance317.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance317.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance317.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.RowAlternateAppearance = appearance317;
			appearance318.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance318.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance318.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.RowSelectorAppearance = appearance318;
			appearance319.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance319.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance319.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.SelectedRowAppearance = appearance319;
			this.uGridAgrupadoPorDias.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridAgrupadoPorDias.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAgrupadoPorDias.Location = new System.Drawing.Point(264, 8);
			this.uGridAgrupadoPorDias.Name = "uGridAgrupadoPorDias";
			this.uGridAgrupadoPorDias.Size = new System.Drawing.Size(976, 704);
			this.uGridAgrupadoPorDias.TabIndex = 661;
			this.uGridAgrupadoPorDias.Visible = false;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn123.DataType = typeof(System.DateTime);
			ultraDataColumn123.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn124.DataType = typeof(System.Decimal);
			ultraDataColumn124.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn125.DataType = typeof(System.Decimal);
			ultraDataColumn125.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn126.DataType = typeof(System.Decimal);
			ultraDataColumn126.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn127.DataType = typeof(System.Decimal);
			ultraDataColumn127.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn128.DataType = typeof(System.Decimal);
			ultraDataColumn128.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn129.DataType = typeof(System.Decimal);
			ultraDataColumn129.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn130.DataType = typeof(System.Decimal);
			ultraDataColumn130.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn131.DataType = typeof(System.Decimal);
			ultraDataColumn131.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn132.DataType = typeof(System.Decimal);
			ultraDataColumn132.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn133.DataType = typeof(System.Decimal);
			ultraDataColumn133.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn134.DataType = typeof(System.Decimal);
			ultraDataColumn134.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn135.DataType = typeof(System.Decimal);
			ultraDataColumn135.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn136.DataType = typeof(System.Decimal);
			ultraDataColumn136.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn137.DataType = typeof(System.Decimal);
			ultraDataColumn137.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn138.DataType = typeof(System.Decimal);
			ultraDataColumn138.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn139.DataType = typeof(System.Decimal);
			ultraDataColumn139.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn140.DataType = typeof(System.Decimal);
			ultraDataColumn140.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn141.DataType = typeof(System.Decimal);
			ultraDataColumn141.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn142.DataType = typeof(System.Decimal);
			ultraDataColumn142.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn122,
																																 ultraDataColumn123,
																																 ultraDataColumn124,
																																 ultraDataColumn125,
																																 ultraDataColumn126,
																																 ultraDataColumn127,
																																 ultraDataColumn128,
																																 ultraDataColumn129,
																																 ultraDataColumn130,
																																 ultraDataColumn131,
																																 ultraDataColumn132,
																																 ultraDataColumn133,
																																 ultraDataColumn134,
																																 ultraDataColumn135,
																																 ultraDataColumn136,
																																 ultraDataColumn137,
																																 ultraDataColumn138,
																																 ultraDataColumn139,
																																 ultraDataColumn140,
																																 ultraDataColumn141,
																																 ultraDataColumn142});
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(168, 688);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(75, 24);
			this.btnExportar.TabIndex = 678;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.SystemColors.Control;
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(88, 688);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(75, 24);
			this.btnImprimir.TabIndex = 679;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// rdgReporte
			// 
			appearance320.FontData.Name = "Tahoma";
			appearance320.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgReporte.Appearance = appearance320;
			this.rdgReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rdgReporte.CheckedIndex = 0;
			this.rdgReporte.ItemAppearance = appearance321;
			this.rdgReporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Horizontal";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Por Local";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Por Dias";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "Por Vendedor";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "Por Local y Vendedor";
			valueListItem6.DataValue = 6;
			valueListItem6.DisplayText = "Agrupado por Dias";
			valueListItem7.DataValue = 7;
			valueListItem7.DisplayText = "Remates";
			valueListItem8.DataValue = 8;
			valueListItem8.DisplayText = "GEx Horizontal";
			valueListItem9.DataValue = 9;
			valueListItem9.DisplayText = "GEx Por Local";
			valueListItem10.DataValue = 10;
			valueListItem10.DisplayText = "D & D";
			this.rdgReporte.Items.Add(valueListItem1);
			this.rdgReporte.Items.Add(valueListItem2);
			this.rdgReporte.Items.Add(valueListItem3);
			this.rdgReporte.Items.Add(valueListItem4);
			this.rdgReporte.Items.Add(valueListItem5);
			this.rdgReporte.Items.Add(valueListItem6);
			this.rdgReporte.Items.Add(valueListItem7);
			this.rdgReporte.Items.Add(valueListItem8);
			this.rdgReporte.Items.Add(valueListItem9);
			this.rdgReporte.Items.Add(valueListItem10);
			this.rdgReporte.ItemSpacingVertical = 10;
			this.rdgReporte.Location = new System.Drawing.Point(2, 8);
			this.rdgReporte.Name = "rdgReporte";
			this.rdgReporte.Size = new System.Drawing.Size(248, 128);
			this.rdgReporte.TabIndex = 680;
			this.rdgReporte.Text = "Horizontal";
			this.rdgReporte.ValueChanged += new System.EventHandler(this.rdgReporte_ValueChanged);
			// 
			// btnVer
			// 
			this.btnVer.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(8, 688);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 681;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn143.DataType = typeof(bool);
			ultraDataColumn143.DefaultValue = false;
			ultraDataColumn146.DataType = typeof(System.DateTime);
			ultraDataColumn146.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn153.DataType = typeof(System.Double);
			ultraDataColumn153.DefaultValue = 0;
			ultraDataColumn154.DataType = typeof(System.Double);
			ultraDataColumn154.DefaultValue = 0;
			ultraDataColumn155.DataType = typeof(System.Double);
			ultraDataColumn155.DefaultValue = 0;
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn143,
																																 ultraDataColumn144,
																																 ultraDataColumn145,
																																 ultraDataColumn146,
																																 ultraDataColumn147,
																																 ultraDataColumn148,
																																 ultraDataColumn149,
																																 ultraDataColumn150,
																																 ultraDataColumn151,
																																 ultraDataColumn152,
																																 ultraDataColumn153,
																																 ultraDataColumn154,
																																 ultraDataColumn155});
			// 
			// uGridHorizontalGex
			// 
			this.uGridHorizontalGex.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridHorizontalGex.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridHorizontalGex.DataSource = this.ultraDataSource9;
			appearance322.BackColor = System.Drawing.Color.White;
			this.uGridHorizontalGex.DisplayLayout.Appearance = appearance322;
			ultraGridColumn146.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance323.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn146.CellAppearance = appearance323;
			ultraGridColumn146.Format = "#,##0.00";
			ultraGridColumn146.Header.VisiblePosition = 0;
			ultraGridColumn146.Width = 20;
			ultraGridColumn147.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn147.Header.VisiblePosition = 1;
			ultraGridColumn147.Width = 40;
			ultraGridColumn148.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn148.Header.VisiblePosition = 2;
			ultraGridColumn148.Width = 90;
			ultraGridColumn149.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn149.Header.VisiblePosition = 3;
			ultraGridColumn149.Width = 75;
			ultraGridColumn150.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn150.Header.VisiblePosition = 4;
			ultraGridColumn150.Width = 60;
			ultraGridColumn151.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn151.Header.VisiblePosition = 5;
			ultraGridColumn151.Width = 70;
			ultraGridColumn152.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn152.Header.VisiblePosition = 6;
			ultraGridColumn152.Width = 100;
			ultraGridColumn153.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn153.Header.Caption = "Cliente";
			ultraGridColumn153.Header.VisiblePosition = 7;
			ultraGridColumn153.Width = 240;
			ultraGridColumn154.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn154.Header.VisiblePosition = 8;
			ultraGridColumn154.Width = 50;
			ultraGridColumn155.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn155.Header.VisiblePosition = 9;
			ultraGridColumn155.Width = 240;
			ultraGridColumn156.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance324.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn156.CellAppearance = appearance324;
			ultraGridColumn156.Format = "#,##0.00";
			ultraGridColumn156.Header.VisiblePosition = 10;
			ultraGridColumn156.Width = 80;
			ultraGridColumn157.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance325.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn157.CellAppearance = appearance325;
			ultraGridColumn157.Format = "#,##0.00";
			ultraGridColumn157.Header.VisiblePosition = 11;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn157.Width = 80;
			ultraGridColumn158.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance326.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn158.CellAppearance = appearance326;
			ultraGridColumn158.Format = "#,##0.00";
			ultraGridColumn158.Header.VisiblePosition = 12;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn158.Width = 80;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn146,
																											ultraGridColumn147,
																											ultraGridColumn148,
																											ultraGridColumn149,
																											ultraGridColumn150,
																											ultraGridColumn151,
																											ultraGridColumn152,
																											ultraGridColumn153,
																											ultraGridColumn154,
																											ultraGridColumn155,
																											ultraGridColumn156,
																											ultraGridColumn157,
																											ultraGridColumn158});
			appearance327.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings121.Appearance = appearance327;
			summarySettings121.DisplayFormat = "{0: #,##0.00}";
			summarySettings121.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance328.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings122.Appearance = appearance328;
			summarySettings122.DisplayFormat = "{0: #,##0.00}";
			summarySettings122.Hidden = true;
			summarySettings122.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance329.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings123.Appearance = appearance329;
			summarySettings123.DisplayFormat = "{0: #,##0.00}";
			summarySettings123.Hidden = true;
			summarySettings123.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand10.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings121,
																																															 summarySettings122,
																																															 summarySettings123});
			ultraGridBand10.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridHorizontalGex.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			appearance330.ForeColor = System.Drawing.Color.Black;
			appearance330.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontalGex.DisplayLayout.Override.ActiveRowAppearance = appearance330;
			this.uGridHorizontalGex.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridHorizontalGex.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHorizontalGex.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance331.BackColor = System.Drawing.Color.Transparent;
			this.uGridHorizontalGex.DisplayLayout.Override.CardAreaAppearance = appearance331;
			appearance332.ForeColor = System.Drawing.Color.Black;
			appearance332.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontalGex.DisplayLayout.Override.GroupByRowAppearance = appearance332;
			appearance333.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance333.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance333.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance333.FontData.BoldAsString = "True";
			appearance333.FontData.Name = "Arial";
			appearance333.FontData.SizeInPoints = 8.5F;
			appearance333.ForeColor = System.Drawing.Color.White;
			appearance333.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHorizontalGex.DisplayLayout.Override.HeaderAppearance = appearance333;
			this.uGridHorizontalGex.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance334.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance334.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance334.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontalGex.DisplayLayout.Override.RowAlternateAppearance = appearance334;
			appearance335.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance335.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance335.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontalGex.DisplayLayout.Override.RowSelectorAppearance = appearance335;
			appearance336.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance336.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance336.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontalGex.DisplayLayout.Override.SelectedRowAppearance = appearance336;
			this.uGridHorizontalGex.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHorizontalGex.Location = new System.Drawing.Point(264, 8);
			this.uGridHorizontalGex.Name = "uGridHorizontalGex";
			this.uGridHorizontalGex.Size = new System.Drawing.Size(976, 704);
			this.uGridHorizontalGex.TabIndex = 682;
			this.uGridHorizontalGex.Visible = false;
			// 
			// uGridGExPorLocal
			// 
			this.uGridGExPorLocal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridGExPorLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridGExPorLocal.DataSource = this.ultraDataSource10;
			appearance337.BackColor = System.Drawing.Color.White;
			this.uGridGExPorLocal.DisplayLayout.Appearance = appearance337;
			this.uGridGExPorLocal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn159.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn159.Header.VisiblePosition = 0;
			ultraGridColumn159.Width = 287;
			ultraGridColumn160.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance338.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn160.CellAppearance = appearance338;
			ultraGridColumn160.Format = "#,##0.00";
			ultraGridColumn160.Header.VisiblePosition = 1;
			ultraGridColumn160.Width = 148;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn159,
																											ultraGridColumn160});
			appearance339.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings124.Appearance = appearance339;
			summarySettings124.DisplayFormat = "{0: #,##0.00}";
			summarySettings124.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand11.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings124});
			ultraGridBand11.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridGExPorLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			appearance340.ForeColor = System.Drawing.Color.Black;
			appearance340.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGExPorLocal.DisplayLayout.Override.ActiveRowAppearance = appearance340;
			this.uGridGExPorLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGExPorLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGExPorLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance341.BackColor = System.Drawing.Color.Transparent;
			this.uGridGExPorLocal.DisplayLayout.Override.CardAreaAppearance = appearance341;
			appearance342.ForeColor = System.Drawing.Color.Black;
			appearance342.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGExPorLocal.DisplayLayout.Override.GroupByRowAppearance = appearance342;
			appearance343.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance343.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance343.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance343.FontData.BoldAsString = "True";
			appearance343.FontData.Name = "Arial";
			appearance343.FontData.SizeInPoints = 8.5F;
			appearance343.ForeColor = System.Drawing.Color.White;
			appearance343.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGExPorLocal.DisplayLayout.Override.HeaderAppearance = appearance343;
			this.uGridGExPorLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance344.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance344.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance344.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGExPorLocal.DisplayLayout.Override.RowAlternateAppearance = appearance344;
			appearance345.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance345.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance345.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGExPorLocal.DisplayLayout.Override.RowSelectorAppearance = appearance345;
			appearance346.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance346.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance346.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGExPorLocal.DisplayLayout.Override.SelectedRowAppearance = appearance346;
			this.uGridGExPorLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGExPorLocal.Location = new System.Drawing.Point(264, 8);
			this.uGridGExPorLocal.Name = "uGridGExPorLocal";
			this.uGridGExPorLocal.Size = new System.Drawing.Size(456, 704);
			this.uGridGExPorLocal.TabIndex = 683;
			this.uGridGExPorLocal.Visible = false;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn157.DataType = typeof(System.Decimal);
			ultraDataColumn157.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn156,
																																	ultraDataColumn157});
			// 
			// uGridDYD
			// 
			this.uGridDYD.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDYD.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDYD.DataSource = this.ultraDataSource11;
			appearance347.BackColor = System.Drawing.Color.White;
			this.uGridDYD.DisplayLayout.Appearance = appearance347;
			this.uGridDYD.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn161.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn161.Header.VisiblePosition = 0;
			ultraGridColumn161.Width = 118;
			ultraGridColumn162.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn162.Header.VisiblePosition = 1;
			ultraGridColumn162.Width = 192;
			ultraGridColumn163.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn163.Header.VisiblePosition = 2;
			ultraGridColumn163.Width = 78;
			ultraGridColumn164.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn164.Header.VisiblePosition = 3;
			ultraGridColumn164.Width = 78;
			ultraGridColumn165.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance348.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn165.CellAppearance = appearance348;
			ultraGridColumn165.Format = "#,##0.00";
			ultraGridColumn165.Header.Caption = "Valor D & D";
			ultraGridColumn165.Header.VisiblePosition = 4;
			ultraGridColumn165.Width = 102;
			ultraGridColumn166.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance349.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn166.CellAppearance = appearance349;
			ultraGridColumn166.Format = "#,##0.00";
			ultraGridColumn166.Header.Caption = "I.V.A. D & D";
			ultraGridColumn166.Header.VisiblePosition = 5;
			ultraGridColumn166.Width = 103;
			ultraGridColumn167.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance350.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn167.CellAppearance = appearance350;
			ultraGridColumn167.Format = "#,##0.00";
			ultraGridColumn167.Header.Caption = "D & D Con I.V.A.";
			ultraGridColumn167.Header.VisiblePosition = 6;
			ultraGridColumn167.Width = 100;
			ultraGridColumn168.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance351.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn168.CellAppearance = appearance351;
			ultraGridColumn168.Format = "#,##0.00";
			ultraGridColumn168.Header.Caption = "Total Factura";
			ultraGridColumn168.Header.VisiblePosition = 7;
			ultraGridColumn168.Width = 104;
			ultraGridColumn169.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance352.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn169.CellAppearance = appearance352;
			ultraGridColumn169.Format = "#,##0.00";
			ultraGridColumn169.Header.Caption = "Monto Credito";
			ultraGridColumn169.Header.VisiblePosition = 8;
			ultraGridColumn169.Width = 128;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn161,
																											ultraGridColumn162,
																											ultraGridColumn163,
																											ultraGridColumn164,
																											ultraGridColumn165,
																											ultraGridColumn166,
																											ultraGridColumn167,
																											ultraGridColumn168,
																											ultraGridColumn169});
			appearance353.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings125.Appearance = appearance353;
			summarySettings125.DisplayFormat = "{0: #,##0.00}";
			summarySettings125.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance354.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings126.Appearance = appearance354;
			summarySettings126.DisplayFormat = "{0: #,##0.00}";
			summarySettings126.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance355.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings127.Appearance = appearance355;
			summarySettings127.DisplayFormat = "{0: #,##0.00}";
			summarySettings127.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance356.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings128.Appearance = appearance356;
			summarySettings128.DisplayFormat = "{0: #,##0.00}";
			summarySettings128.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance357.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings129.Appearance = appearance357;
			summarySettings129.DisplayFormat = "{0: #,##0.00}";
			summarySettings129.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand12.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings125,
																																															 summarySettings126,
																																															 summarySettings127,
																																															 summarySettings128,
																																															 summarySettings129});
			ultraGridBand12.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridDYD.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			appearance358.ForeColor = System.Drawing.Color.Black;
			appearance358.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.ActiveRowAppearance = appearance358;
			this.uGridDYD.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDYD.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDYD.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance359.BackColor = System.Drawing.Color.Transparent;
			this.uGridDYD.DisplayLayout.Override.CardAreaAppearance = appearance359;
			appearance360.ForeColor = System.Drawing.Color.Black;
			appearance360.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.CellAppearance = appearance360;
			appearance361.ForeColor = System.Drawing.Color.Black;
			appearance361.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.GroupByRowAppearance = appearance361;
			appearance362.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance362.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance362.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance362.FontData.BoldAsString = "True";
			appearance362.FontData.Name = "Arial";
			appearance362.FontData.SizeInPoints = 8.5F;
			appearance362.ForeColor = System.Drawing.Color.White;
			appearance362.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDYD.DisplayLayout.Override.HeaderAppearance = appearance362;
			this.uGridDYD.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance363.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance363.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance363.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowAlternateAppearance = appearance363;
			appearance364.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance364.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance364.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowSelectorAppearance = appearance364;
			appearance365.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance365.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance365.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.SelectedRowAppearance = appearance365;
			this.uGridDYD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDYD.Location = new System.Drawing.Point(264, 7);
			this.uGridDYD.Name = "uGridDYD";
			this.uGridDYD.Size = new System.Drawing.Size(1024, 704);
			this.uGridDYD.TabIndex = 684;
			this.uGridDYD.Visible = false;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn160.DataType = typeof(System.DateTime);
			ultraDataColumn162.DataType = typeof(System.Decimal);
			ultraDataColumn163.DataType = typeof(System.Decimal);
			ultraDataColumn164.DataType = typeof(System.Decimal);
			ultraDataColumn165.DataType = typeof(System.Decimal);
			ultraDataColumn166.DataType = typeof(System.Decimal);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn158,
																																	ultraDataColumn159,
																																	ultraDataColumn160,
																																	ultraDataColumn161,
																																	ultraDataColumn162,
																																	ultraDataColumn163,
																																	ultraDataColumn164,
																																	ultraDataColumn165,
																																	ultraDataColumn166});
			// 
			// chkturno1
			// 
			appearance366.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkturno1.Appearance = appearance366;
			this.chkturno1.BackColor = System.Drawing.Color.Transparent;
			this.chkturno1.Enabled = false;
			this.chkturno1.Location = new System.Drawing.Point(176, 144);
			this.chkturno1.Name = "chkturno1";
			this.chkturno1.Size = new System.Drawing.Size(72, 21);
			this.chkturno1.TabIndex = 685;
			this.chkturno1.Text = "Turno 1";
			this.chkturno1.Visible = false;
			this.chkturno1.CheckedChanged += new System.EventHandler(this.chkturno1_CheckedChanged);
			// 
			// chkturno2
			// 
			appearance367.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkturno2.Appearance = appearance367;
			this.chkturno2.BackColor = System.Drawing.Color.Transparent;
			this.chkturno2.Enabled = false;
			this.chkturno2.Location = new System.Drawing.Point(176, 168);
			this.chkturno2.Name = "chkturno2";
			this.chkturno2.Size = new System.Drawing.Size(72, 21);
			this.chkturno2.TabIndex = 686;
			this.chkturno2.Text = "Turno 2";
			this.chkturno2.Visible = false;
			this.chkturno2.CheckedChanged += new System.EventHandler(this.chkturno2_CheckedChanged);
			// 
			// frmResumenDeVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1248, 718);
			this.Controls.Add(this.chkturno2);
			this.Controls.Add(this.chkturno1);
			this.Controls.Add(this.uGridDYD);
			this.Controls.Add(this.uGridGExPorLocal);
			this.Controls.Add(this.uGridAgrupadoPorDias);
			this.Controls.Add(this.uGridPorlocalVendedor);
			this.Controls.Add(this.uGridPorVendedor);
			this.Controls.Add(this.uGridPorDias);
			this.Controls.Add(this.uGridPorLocal);
			this.Controls.Add(this.uGridHorizontal);
			this.Controls.Add(this.uGridHorizontalGex);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.rdgReporte);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.chkMatriz);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.uGridMatriz);
			this.Controls.Add(this.uGridLocales);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmResumenDeVentas";
			this.Text = "Resumen De Ventas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmResumenDeVentas_KeyDown);
			this.Load += new System.EventHandler(this.frmResumenDeVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridHorizontal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridMatriz)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridPorlocalVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridAgrupadoPorDias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridHorizontalGex)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridGExPorLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDYD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void frmResumenDeVentas_Load(object sender, System.EventArgs e)
		{		
			Funcion.Resolucion(this, this.Height, this.Width);

			miAcceso = new Acceso(cdsSeteoF, "0712");

			if (miAcceso.BLimiteBusquedaFechas)
			{
				DateTime dtInicio = DateTime.Today.AddMonths(-1);

				this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());				
			}
			else
			{
				this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");				
			}
			
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ResumenDeVentasLocales {0}", MenuLatinium.IdUsuario));
				//Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Convert(Bit, 0) As Sel From Bodega Where Activo = 1 And Factura = 1 And Sel = 0 Order By Nombre");
			
			//

			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From Personal Where Vendedor = 1 Order By Nombre");

			if (miAcceso.BVerMatriz)
			{
				this.uGridMatriz.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Convert(Bit, 0) As Sel From Bodega Where Activo = 1 And Factura = 1 And Sel = 1 Order By Nombre");
				this.chkMatriz.Enabled = true;
			}

			if (this.uGridLocales.Rows.Count > 1) 
			{
				this.chkTodos.Enabled = true;
				this.uGridLocales.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			
			this.chkTodos.Checked = true;			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void chkMatriz_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridMatriz.Rows.All)
			{
				if (this.chkMatriz.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void uGridPorDias_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtHasta_Validated(object sender, System.EventArgs e)
		{
			if (this.dtHasta.Value == System.DBNull.Value) this.dtHasta.Value = DateTime.Today;
		}

		private void dtDesde_Validated(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value == System.DBNull.Value) if (this.dtHasta.Value != System.DBNull.Value) this.dtDesde.Value = (DateTime)this.dtHasta.Value; else this.dtDesde.Value = DateTime.Today;
		}

		private void frmResumenDeVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridHorizontal_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["Estado"].Value.ToString() == "ANULADO")
			{			
				e.Row.Appearance.BackColor = Color.Red;
				e.Row.Appearance.BackColor2 = Color.Red;				
			}
			if (e.Row.Cells["Estado"].Value.ToString() == "MANUAL")
			{			
				e.Row.Appearance.BackColor = Color.Yellow;
				e.Row.Appearance.BackColor2 = Color.Yellow;				
			}
		}

		private void rdgReporte_ValueChanged(object sender, System.EventArgs e)
		{	
		  this.uGridHorizontal.Visible = false;
			this.uGridPorDias.Visible = false;
			this.uGridPorLocal.Visible = false;
			this.uGridPorlocalVendedor.Visible = false;
			this.uGridPorVendedor.Visible = false;
			this.uGridAgrupadoPorDias.Visible = false;
			this.uGridHorizontalGex.Visible = false;
			this.uGridGExPorLocal.Visible = false;
			this.uGridDYD.Visible = false;

			this.btnVer.Enabled = false;

			if (miAcceso.BVerRemates || (int)this.rdgReporte.Value != 7) 
				this.btnVer.Enabled = true;

			if ((int)this.rdgReporte.Value == 1 || (int)this.rdgReporte.Value == 7)				
				this.uGridHorizontal.Visible = true;						
			
			if ((int)this.rdgReporte.Value == 2)								
				this.uGridPorLocal.Visible = true;				
			
			if ((int)this.rdgReporte.Value == 3)								
				this.uGridPorDias.Visible = true;				
			
			if ((int)this.rdgReporte.Value == 4)								
				this.uGridPorVendedor.Visible = true;				
			
			if ((int)this.rdgReporte.Value == 5)								
				this.uGridPorlocalVendedor.Visible = true;				
			
			if ((int)this.rdgReporte.Value == 6)
				this.uGridAgrupadoPorDias.Visible = true;			
			
			if ((int)this.rdgReporte.Value == 8)
				this.uGridHorizontalGex.Visible = true;		

			if ((int)this.rdgReporte.Value == 9)
				this.uGridGExPorLocal.Visible = true;

			if ((int)this.rdgReporte.Value == 10)
				this.uGridDYD.Visible = true;
		
			if (MenuLatinium.stIdDB == 34 )
			{
				if ((int) this.rdgReporte.Value == 6)
				{

					this.chkturno1.Visible = true;
					this.chkturno2.Visible = true;
					this.chkturno1.Enabled = true;
					this.chkturno2.Enabled = true;
				}
				else
				{

					this.chkturno1.Visible = false;
					this.chkturno2.Visible = false;
					this.chkturno1.Enabled = false;
					this.chkturno2.Enabled = false;
				}
			}
			this.Consulta();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridHorizontal.Rows.All)
//				if (dr.Cells["DOC"].Value.ToString() == "FAC")
//					if (dr.Cells["Estado"].Value.ToString() != "ANULADO")
//						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CreaAsientoDeVenta {0}", (int)dr.Cells["idCompra"].Value));
//
//			MessageBox.Show("OKEYES");

			if ((int)this.rdgReporte.Value == 1 || (int)this.rdgReporte.Value == 7)
				FuncionesProcedimientos.ExportaExcel(this.uGridHorizontal);			
			
			if ((int)this.rdgReporte.Value == 2)
				FuncionesProcedimientos.ExportaExcel(this.uGridPorLocal);
		
			if ((int)this.rdgReporte.Value == 3)
				FuncionesProcedimientos.ExportaExcel(this.uGridPorDias);
		
			if ((int)this.rdgReporte.Value == 4)
				FuncionesProcedimientos.ExportaExcel(this.uGridPorVendedor);
		
			if ((int)this.rdgReporte.Value == 5)
				FuncionesProcedimientos.ExportaExcel(this.uGridPorlocalVendedor);
		
			if ((int)this.rdgReporte.Value == 6)
				FuncionesProcedimientos.ExportaExcel(this.uGridAgrupadoPorDias);
	
			if ((int)this.rdgReporte.Value == 8)
				FuncionesProcedimientos.ExportaExcel(this.uGridHorizontalGex);

			if ((int)this.rdgReporte.Value == 9)
				FuncionesProcedimientos.ExportaExcel(this.uGridGExPorLocal);

			if ((int)this.rdgReporte.Value == 10)
				FuncionesProcedimientos.ExportaExcel(this.uGridDYD);			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{			
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;

			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@Desde";
			discreteValDesde.Value = dtDesde;
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);

			ParameterField paramFieldHasta = new ParameterField ();
			ParameterDiscreteValue discreteValHasta = new ParameterDiscreteValue ();
			paramFieldHasta.ParameterFieldName = "@Hasta";
			discreteValHasta.Value = dtHasta;
			paramFieldHasta.CurrentValues.Add (discreteValHasta);
			paramFields.Add (paramFieldHasta);

			string sLocales = "";
			
			int idVendedor = 0;
			if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridMatriz.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			ParameterField paramFieldInicial = new ParameterField ();
			ParameterDiscreteValue discreteValInicial = new ParameterDiscreteValue ();
			paramFieldInicial.ParameterFieldName = "@Locales";
			discreteValInicial.Value = sLocales; 
			paramFieldInicial.CurrentValues.Add (discreteValInicial);
			paramFields.Add (paramFieldInicial);

			ParameterField paramFieldFinal = new ParameterField ();
			ParameterDiscreteValue discreteValFinal = new ParameterDiscreteValue ();
			paramFieldFinal.ParameterFieldName = "@idVendedor";
			discreteValFinal.Value = idVendedor; 
			paramFieldFinal.CurrentValues.Add (discreteValFinal);
			paramFields.Add (paramFieldFinal);

			string sReporte = "";

			if ((int)this.rdgReporte.Value == 1)
			{				
				ParameterField paramFieldOrden = new ParameterField ();
				ParameterDiscreteValue discreteValOrden = new ParameterDiscreteValue ();
				paramFieldOrden.ParameterFieldName = "@Orden";
				discreteValOrden.Value = 0; 
				paramFieldOrden.CurrentValues.Add (discreteValOrden);
				paramFields.Add (paramFieldOrden);

				sReporte = "ReporteDeVentasHorizontal.rpt";							
			}
			else if ((int)this.rdgReporte.Value == 2)	sReporte = "ReporteDeVentasPorLocal.rpt";
			else if ((int)this.rdgReporte.Value == 3)	sReporte = "ReporteDeVentasPorDias.rpt";
			else if ((int)this.rdgReporte.Value == 4) sReporte = "ReporteDeVentasPorVendedor.rpt";	
			else if ((int)this.rdgReporte.Value == 6) sReporte = "ReporteDeVentasAgrupadoPorDia.rpt";		
			else if ((int)this.rdgReporte.Value == 10) sReporte = "ReporteDeVentasDYD.rpt";	
			else if ((int)this.rdgReporte.Value == 7) sReporte = "ReporteDeVentasRemates.rpt";	
			
			if ((int)this.rdgReporte.Value == 5)
			{
				ParameterField paramFieldOrden = new ParameterField ();
				ParameterDiscreteValue discreteValOrden = new ParameterDiscreteValue ();
				paramFieldOrden.ParameterFieldName = "@Carpa";
				discreteValOrden.Value = 0; 
				paramFieldOrden.CurrentValues.Add (discreteValOrden);
				paramFields.Add (paramFieldOrden);
				sReporte = "ReporteDeVentasLocalPorVendedor.rpt";
			}
			if ((int)this.rdgReporte.Value == 9)
			{
				ParameterField paramFieldOrden = new ParameterField ();
				ParameterDiscreteValue discreteValOrden = new ParameterDiscreteValue ();
				paramFieldOrden.ParameterFieldName = "@Orden";
				discreteValOrden.Value = 0; 
				paramFieldOrden.CurrentValues.Add (discreteValOrden);
				paramFields.Add (paramFieldOrden);
				sReporte = "ReporteDeVentasGEX.rpt";
			}
			if ((int)this.rdgReporte.Value == 8)
			{
				ParameterField paramFieldOrden = new ParameterField ();
				ParameterDiscreteValue discreteValOrden = new ParameterDiscreteValue ();
				paramFieldOrden.ParameterFieldName = "@Orden";
				discreteValOrden.Value = 0; 
				paramFieldOrden.CurrentValues.Add (discreteValOrden);
				paramFields.Add (paramFieldOrden);
				sReporte = "ReporteDeVentasGexHoriz.rpt";
			}
			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();

		}

		private void Consulta()
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false, "", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;

			this.Text = "Resumen De Ventas";

			string sLocales = "";			
			int idVendedor = 0;
			string sSQL = "";

			if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridMatriz.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString();
					else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if ((int)this.rdgReporte.Value == 1)
			{				
				this.uGridHorizontal.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, false);			
				
				sSQL = string.Format("Exec ReporteDeVentasHorizontal '{0}', '{1}', '{2}', {3}, {4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor, 0);
				
				this.uGridHorizontal.DataSource = this.ultraDataSource3;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridHorizontal);				

				this.Text = this.Text + " - " + this.uGridHorizontal.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 2)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorLocal '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridPorLocal.DataSource = this.ultraDataSource4;

				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridPorLocal);

				this.Text = this.Text + " - " + this.uGridPorLocal.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 3)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorDias '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridPorDias.DataSource = this.ultraDataSource5;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridPorDias);

				this.Text = this.Text + " - " + this.uGridPorDias.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 4)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorVendedor '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridPorVendedor.DataSource = this.ultraDataSource6;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridPorVendedor);

				this.Text = this.Text + " - " + this.uGridPorVendedor.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 5)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorLocalVendedor '{0}', '{1}', '{2}', {3}, {4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor, 0);
				
				this.uGridPorlocalVendedor.DataSource = this.ultraDataSource7;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridPorlocalVendedor);

				this.uGridPorlocalVendedor.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, true);			
				this.uGridPorlocalVendedor.Rows.CollapseAll(true);

				this.Text = this.Text + " - " + this.uGridPorlocalVendedor.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 6)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorDiasAgrupados '{0}', '{1}', '{2}', {3}, {4}, {5}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor,
					(bool) this.chkturno1.Checked, (bool) this.chkturno2.Checked 
					);
			
				this.uGridAgrupadoPorDias.DataSource = this.ultraDataSource8;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridAgrupadoPorDias);

				this.uGridAgrupadoPorDias.DisplayLayout.Bands[0].SortedColumns.Add("Dia", false, true);			
				this.uGridAgrupadoPorDias.Rows.ExpandAll(true);

				this.Text = this.Text + " - " + this.uGridPorDias.Rows.Count + " Registros Encontrados";

				return;
			}
						
			if ((int)this.rdgReporte.Value == 7)
			{				
				if (miAcceso.BVerRemates) sSQL = string.Format("Exec ReporteDeVentasRemates '{0}', '{1}', '{2}', {3}", 
																		Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);							
				else sSQL = "Exec ReporteDeVentasRemates '20000101', '20000101', '', 0";

				this.uGridHorizontal.DataSource = this.ultraDataSource3;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridHorizontal);
				
				this.uGridHorizontal.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, true);			
				this.uGridHorizontal.Rows.ExpandAll(true);

				this.Text = this.Text + " - " + this.uGridHorizontal.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 8)
			{				
				sSQL = string.Format("Exec ReporteDeVentasHorizontalGEx '{0}', '{1}', '{2}', {3}, {4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor, 0);
				
				this.uGridHorizontalGex.DataSource = this.ultraDataSource9;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridHorizontalGex);				

				this.Text = this.Text + " - " + this.uGridHorizontalGex.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 9)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorLocalGEx '{0}', '{1}', '{2}', {3}, {4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor, 0);
				
				this.uGridGExPorLocal.DataSource = this.ultraDataSource10;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridGExPorLocal);				

				this.Text = this.Text + " - " + this.uGridGExPorLocal.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 10)
			{				
				sSQL = string.Format("Exec ReporteDeVentasDyD '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridDYD.DataSource = this.ultraDataSource11;
				FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridDYD);				

				this.Text = this.Text + " - " + this.uGridDYD.Rows.Count + " Registros Encontrados";

				return;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void chkturno1_CheckedChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click(sender , e);
		}

		private void chkturno2_CheckedChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click(sender , e);
		}
	}
}

