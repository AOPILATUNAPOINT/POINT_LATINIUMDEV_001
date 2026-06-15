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
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 9, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 10, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 12, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 13, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 14, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 15, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 16, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 17, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 18, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 19, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 20, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 21, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 21, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 22, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 22, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 23, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 23, true);
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 24, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 24, true);
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 25, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 25, true);
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmResumenDeVentas));
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 2, true);
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 3, true);
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 4, true);
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 5, true);
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 6, true);
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 7, true);
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 8, true);
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 10, true);
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings27 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 9, true);
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings28 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 11, true);
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings29 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 12, true);
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings30 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 13, true);
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings31 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 14, true);
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings32 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings33 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 16, true);
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings34 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 17, true);
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings35 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 18, true);
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings36 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 19, true);
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings37 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings38 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 2, true);
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings39 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 3, true);
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings40 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 4, true);
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings41 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 5, true);
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings42 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 6, true);
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings43 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 7, true);
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings44 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 8, true);
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings45 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 10, true);
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings46 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 9, true);
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings47 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 11, true);
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings48 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 12, true);
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings49 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 13, true);
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings50 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 14, true);
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings51 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings52 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 16, true);
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings53 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 17, true);
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings54 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 19, true);
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings55 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 18, true);
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance179 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance180 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance181 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance182 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings56 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance183 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings57 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 2, true);
			Infragistics.Win.Appearance appearance184 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings58 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 3, true);
			Infragistics.Win.Appearance appearance185 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings59 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 4, true);
			Infragistics.Win.Appearance appearance186 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings60 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 5, true);
			Infragistics.Win.Appearance appearance187 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings61 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 6, true);
			Infragistics.Win.Appearance appearance188 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings62 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 7, true);
			Infragistics.Win.Appearance appearance189 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings63 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 8, true);
			Infragistics.Win.Appearance appearance190 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings64 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 10, true);
			Infragistics.Win.Appearance appearance191 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings65 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 9, true);
			Infragistics.Win.Appearance appearance192 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings66 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 11, true);
			Infragistics.Win.Appearance appearance193 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings67 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 12, true);
			Infragistics.Win.Appearance appearance194 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings68 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 13, true);
			Infragistics.Win.Appearance appearance195 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings69 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 14, true);
			Infragistics.Win.Appearance appearance196 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings70 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 15, true);
			Infragistics.Win.Appearance appearance197 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings71 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 16, true);
			Infragistics.Win.Appearance appearance198 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings72 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 17, true);
			Infragistics.Win.Appearance appearance199 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings73 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 18, true);
			Infragistics.Win.Appearance appearance200 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings74 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 19, true);
			Infragistics.Win.Appearance appearance201 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance202 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance203 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance204 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance205 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance206 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance207 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance208 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance209 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance210 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance211 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance212 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance213 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance214 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance215 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance216 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance217 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance218 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance219 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance220 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance221 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance222 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance223 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance224 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance225 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance226 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance227 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance228 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance229 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings75 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 2, true);
			Infragistics.Win.Appearance appearance230 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings76 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 3, true);
			Infragistics.Win.Appearance appearance231 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings77 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 4, true);
			Infragistics.Win.Appearance appearance232 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings78 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Efectivo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 5, true);
			Infragistics.Win.Appearance appearance233 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings79 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cheque", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 6, true);
			Infragistics.Win.Appearance appearance234 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings80 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Tarjeta", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 7, true);
			Infragistics.Win.Appearance appearance235 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings81 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Nota", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 8, true);
			Infragistics.Win.Appearance appearance236 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings82 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Credito", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 9, true);
			Infragistics.Win.Appearance appearance237 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings83 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Anticipo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 11, true);
			Infragistics.Win.Appearance appearance238 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings84 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Transferencia", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 10, true);
			Infragistics.Win.Appearance appearance239 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings85 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CrediPoint", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 12, true);
			Infragistics.Win.Appearance appearance240 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings86 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retencion", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 13, true);
			Infragistics.Win.Appearance appearance241 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings87 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Debito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 14, true);
			Infragistics.Win.Appearance appearance242 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings88 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Deposito", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 15, true);
			Infragistics.Win.Appearance appearance243 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings89 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Bono", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 16, true);
			Infragistics.Win.Appearance appearance244 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings90 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "C_Entrega", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 17, true);
			Infragistics.Win.Appearance appearance245 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings91 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cupon", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 18, true);
			Infragistics.Win.Appearance appearance246 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings92 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Locales", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 19, true);
			Infragistics.Win.Appearance appearance247 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings93 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Carpa", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 20, true);
			Infragistics.Win.Appearance appearance248 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance249 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance250 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance251 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance252 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance253 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance254 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance255 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance256 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance257 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance258 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance259 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance260 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance261 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.Appearance appearance262 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance263 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nota");
			Infragistics.Win.Appearance appearance264 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Credito");
			Infragistics.Win.Appearance appearance265 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.Appearance appearance266 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anticipo");
			Infragistics.Win.Appearance appearance267 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance268 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retencion");
			Infragistics.Win.Appearance appearance269 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Debito");
			Infragistics.Win.Appearance appearance270 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Deposito");
			Infragistics.Win.Appearance appearance271 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.Appearance appearance272 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Entrega");
			Infragistics.Win.Appearance appearance273 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupon");
			Infragistics.Win.Appearance appearance274 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Locales");
			Infragistics.Win.Appearance appearance275 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance276 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings94 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Subtotal", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 2, true);
			Infragistics.Win.Appearance appearance277 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings95 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "IVA", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 3, true);
			Infragistics.Win.Appearance appearance278 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings96 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Total", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 4, true);
			Infragistics.Win.Appearance appearance279 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings97 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Efectivo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Efectivo", 5, true);
			Infragistics.Win.Appearance appearance280 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings98 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Cheque", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cheque", 6, true);
			Infragistics.Win.Appearance appearance281 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings99 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Tarjeta", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Tarjeta", 7, true);
			Infragistics.Win.Appearance appearance282 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings100 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Nota", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Nota", 8, true);
			Infragistics.Win.Appearance appearance283 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings101 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Credito", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Credito", 9, true);
			Infragistics.Win.Appearance appearance284 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings102 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Anticipo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Anticipo", 11, true);
			Infragistics.Win.Appearance appearance285 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings103 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Transferencia", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Transferencia", 10, true);
			Infragistics.Win.Appearance appearance286 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings104 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "CrediPoint", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CrediPoint", 12, true);
			Infragistics.Win.Appearance appearance287 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings105 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Retencion", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retencion", 13, true);
			Infragistics.Win.Appearance appearance288 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings106 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Debito", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Debito", 14, true);
			Infragistics.Win.Appearance appearance289 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings107 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Deposito", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Deposito", 15, true);
			Infragistics.Win.Appearance appearance290 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings108 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Bono", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Bono", 16, true);
			Infragistics.Win.Appearance appearance291 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings109 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "C_Entrega", 17, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "C_Entrega", 17, true);
			Infragistics.Win.Appearance appearance292 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings110 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Cupon", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cupon", 18, true);
			Infragistics.Win.Appearance appearance293 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings111 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Locales", 19, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Locales", 19, true);
			Infragistics.Win.Appearance appearance294 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings112 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Average, null, "Carpa", 20, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Carpa", 20, true);
			Infragistics.Win.Appearance appearance295 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance296 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance297 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance298 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance299 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance300 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance301 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance302 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance303 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Transferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Anticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn126 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn127 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retencion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn128 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Debito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn129 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Deposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn130 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn131 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Entrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn132 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupon");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn133 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Locales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn134 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.Appearance appearance304 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance305 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn135 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn136 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DOC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn137 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn138 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn139 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn140 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn141 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn142 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn143 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn144 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn145 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn146 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn147 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.Appearance appearance306 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Carpa");
			Infragistics.Win.Appearance appearance307 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DOC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance308 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA");
			Infragistics.Win.Appearance appearance309 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance310 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings113 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 10, true);
			Infragistics.Win.Appearance appearance311 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings114 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA", 11, true);
			Infragistics.Win.Appearance appearance312 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings115 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 12, true);
			Infragistics.Win.Appearance appearance313 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance314 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance315 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance316 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance317 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance318 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance319 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance320 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance321 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Subtotal");
			Infragistics.Win.Appearance appearance322 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings116 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Subtotal", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Subtotal", 1, true);
			Infragistics.Win.Appearance appearance323 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance324 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance325 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance326 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance327 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance328 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance329 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance330 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn148 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn149 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Subtotal");
			Infragistics.Win.Appearance appearance331 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Seguro");
			Infragistics.Win.Appearance appearance332 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IVA_Seguro");
			Infragistics.Win.Appearance appearance333 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor_Seguro_Con_IVA");
			Infragistics.Win.Appearance appearance334 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalFactura");
			Infragistics.Win.Appearance appearance335 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto_Credito");
			Infragistics.Win.Appearance appearance336 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings117 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Seguro", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Seguro", 4, true);
			Infragistics.Win.Appearance appearance337 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings118 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "IVA_Seguro", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "IVA_Seguro", 5, true);
			Infragistics.Win.Appearance appearance338 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings119 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor_Seguro_Con_IVA", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor_Seguro_Con_IVA", 6, true);
			Infragistics.Win.Appearance appearance339 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings120 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalFactura", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalFactura", 7, true);
			Infragistics.Win.Appearance appearance340 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings121 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto_Credito", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto_Credito", 8, true);
			Infragistics.Win.Appearance appearance341 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance342 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance343 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance344 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance345 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance346 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance347 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance348 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance349 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn150 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn151 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn152 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn153 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn154 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn155 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IVA_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn156 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro_Con_IVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn157 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn158 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Credito");
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
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance9;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance10;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance11;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance12;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance13;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "T/C";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance14;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "N/C";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance15;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance16;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance17;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Transf";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance18;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "C/P";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance19;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Ret";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance20;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "T/D";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance21;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance22;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance23;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance24;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 25;
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
																										 ultraGridColumn26});
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance25;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance26;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance27;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance28;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance29;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance30;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance31;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance32;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance33;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance34;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance35;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance36;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance37;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance38;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings15.Appearance = appearance39;
			summarySettings15.DisplayFormat = "{0: #,##0.00}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings16.Appearance = appearance40;
			summarySettings16.DisplayFormat = "{0: #,##0.00}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings17.Appearance = appearance41;
			summarySettings17.DisplayFormat = "{0: #,##0.00}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
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
																																															summarySettings17});
			ultraGridBand1.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridHorizontal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridHorizontal.DisplayLayout.GroupByBox.Hidden = true;
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontal.DisplayLayout.Override.ActiveRowAppearance = appearance42;
			this.uGridHorizontal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridHorizontal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHorizontal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance43.BackColor = System.Drawing.Color.Transparent;
			this.uGridHorizontal.DisplayLayout.Override.CardAreaAppearance = appearance43;
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontal.DisplayLayout.Override.CellAppearance = appearance44;
			appearance45.ForeColor = System.Drawing.Color.Black;
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontal.DisplayLayout.Override.GroupByRowAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance46.FontData.BoldAsString = "True";
			appearance46.FontData.Name = "Arial";
			appearance46.FontData.SizeInPoints = 8.5F;
			appearance46.ForeColor = System.Drawing.Color.White;
			appearance46.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHorizontal.DisplayLayout.Override.HeaderAppearance = appearance46;
			this.uGridHorizontal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance47.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance47.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontal.DisplayLayout.Override.RowAlternateAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontal.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontal.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.uGridHorizontal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridHorizontal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHorizontal.Location = new System.Drawing.Point(256, 8);
			this.uGridHorizontal.Name = "uGridHorizontal";
			this.uGridHorizontal.Size = new System.Drawing.Size(984, 704);
			this.uGridHorizontal.TabIndex = 182;
			this.uGridHorizontal.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridHorizontal_InitializeRow);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn26});
			// 
			// dtHasta
			// 
			appearance50.FontData.Name = "Tahoma";
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance50;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(80, 168);
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
			appearance51.FontData.Name = "Tahoma";
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance51;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 144);
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
			appearance52.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance52;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 199;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.Caption = "...";
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn29.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			ultraGridBand2.Header.Caption = "LOCALES";
			ultraGridBand2.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance53.ForeColor = System.Drawing.Color.Black;
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance53;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance54.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance55.FontData.BoldAsString = "True";
			appearance55.FontData.Name = "Arial";
			appearance55.FontData.SizeInPoints = 8.5F;
			appearance55.ForeColor = System.Drawing.Color.White;
			appearance55.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance55;
			appearance56.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance56.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance57;
			appearance58.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance58.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance58.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance58;
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 192);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(240, 272);
			this.uGridLocales.TabIndex = 646;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(bool);
			ultraDataColumn29.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// uGridMatriz
			// 
			this.uGridMatriz.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridMatriz.DataSource = this.ultraDataSource2;
			appearance59.BackColor = System.Drawing.Color.White;
			this.uGridMatriz.DisplayLayout.Appearance = appearance59;
			this.uGridMatriz.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 199;
			ultraGridColumn32.Header.Caption = "...";
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Width = 20;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			ultraGridBand3.Header.Caption = "MATRIZ";
			ultraGridBand3.HeaderVisible = true;
			this.uGridMatriz.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance60.ForeColor = System.Drawing.Color.Black;
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridMatriz.DisplayLayout.Override.ActiveRowAppearance = appearance60;
			this.uGridMatriz.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridMatriz.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridMatriz.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance61.BackColor = System.Drawing.Color.Transparent;
			this.uGridMatriz.DisplayLayout.Override.CardAreaAppearance = appearance61;
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance62.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance62.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance62.FontData.BoldAsString = "True";
			appearance62.FontData.Name = "Arial";
			appearance62.FontData.SizeInPoints = 8.5F;
			appearance62.ForeColor = System.Drawing.Color.White;
			appearance62.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridMatriz.DisplayLayout.Override.HeaderAppearance = appearance62;
			appearance63.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance63.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance63.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMatriz.DisplayLayout.Override.RowAlternateAppearance = appearance63;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMatriz.DisplayLayout.Override.RowSelectorAppearance = appearance64;
			appearance65.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance65.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridMatriz.DisplayLayout.Override.SelectedRowAppearance = appearance65;
			this.uGridMatriz.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridMatriz.Location = new System.Drawing.Point(8, 472);
			this.uGridMatriz.Name = "uGridMatriz";
			this.uGridMatriz.Size = new System.Drawing.Size(240, 176);
			this.uGridMatriz.TabIndex = 647;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(bool);
			ultraDataColumn32.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance66;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 335;
			ultraGridColumn35.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
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
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance67;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(232, 200);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 13);
			this.chkTodos.TabIndex = 653;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// chkMatriz
			// 
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkMatriz.Appearance = appearance68;
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
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn52});
			// 
			// uGridPorLocal
			// 
			this.uGridPorLocal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorLocal.DataSource = this.ultraDataSource4;
			appearance69.BackColor = System.Drawing.Color.White;
			this.uGridPorLocal.DisplayLayout.Appearance = appearance69;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.Caption = "Almacen";
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Width = 160;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance70;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance71;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridColumn38.Width = 80;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance72;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 3;
			ultraGridColumn39.Width = 80;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance73;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.VisiblePosition = 4;
			ultraGridColumn40.Width = 80;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance74;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 5;
			ultraGridColumn41.Width = 80;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance75;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "T/C";
			ultraGridColumn42.Header.VisiblePosition = 6;
			ultraGridColumn42.Width = 80;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance76;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "N/C";
			ultraGridColumn43.Header.VisiblePosition = 7;
			ultraGridColumn43.Width = 80;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance77;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.VisiblePosition = 8;
			ultraGridColumn44.Width = 80;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance78;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "Transf";
			ultraGridColumn45.Header.VisiblePosition = 10;
			ultraGridColumn45.Width = 80;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance79;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 9;
			ultraGridColumn46.Width = 80;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance80;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "C/P";
			ultraGridColumn47.Header.VisiblePosition = 11;
			ultraGridColumn47.Width = 80;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance81;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "Ret";
			ultraGridColumn48.Header.VisiblePosition = 12;
			ultraGridColumn48.Width = 80;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance82;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "T/D";
			ultraGridColumn49.Header.VisiblePosition = 13;
			ultraGridColumn49.Width = 80;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn50.CellAppearance = appearance83;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.VisiblePosition = 14;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance84;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 15;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance85;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 16;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance86;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 17;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn54.CellAppearance = appearance87;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.VisiblePosition = 18;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance88;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 19;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn55});
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings18.Appearance = appearance89;
			summarySettings18.DisplayFormat = "{0: #,##0.00}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance90.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings19.Appearance = appearance90;
			summarySettings19.DisplayFormat = "{0: #,##0.00}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance91.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings20.Appearance = appearance91;
			summarySettings20.DisplayFormat = "{0: #,##0.00}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance92.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings21.Appearance = appearance92;
			summarySettings21.DisplayFormat = "{0: #,##0.00}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance93.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings22.Appearance = appearance93;
			summarySettings22.DisplayFormat = "{0: #,##0.00}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance94.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings23.Appearance = appearance94;
			summarySettings23.DisplayFormat = "{0: #,##0.00}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings24.Appearance = appearance95;
			summarySettings24.DisplayFormat = "{0: #,##0.00}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance96.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings25.Appearance = appearance96;
			summarySettings25.DisplayFormat = "{0: #,##0.00}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance97.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings26.Appearance = appearance97;
			summarySettings26.DisplayFormat = "{0: #,##0.00}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance98.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings27.Appearance = appearance98;
			summarySettings27.DisplayFormat = "{0: #,##0.00}";
			summarySettings27.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance99.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings28.Appearance = appearance99;
			summarySettings28.DisplayFormat = "{0: #,##0.00}";
			summarySettings28.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance100.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings29.Appearance = appearance100;
			summarySettings29.DisplayFormat = "{0: #,##0.00}";
			summarySettings29.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings30.Appearance = appearance101;
			summarySettings30.DisplayFormat = "{0: #,##0.00}";
			summarySettings30.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings31.Appearance = appearance102;
			summarySettings31.DisplayFormat = "{0: #,##0.00}";
			summarySettings31.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings32.Appearance = appearance103;
			summarySettings32.DisplayFormat = "{0: #,##0.00}";
			summarySettings32.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings33.Appearance = appearance104;
			summarySettings33.DisplayFormat = "{0: #,##0.00}";
			summarySettings33.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings34.Appearance = appearance105;
			summarySettings34.DisplayFormat = "{0: #,##0.00}";
			summarySettings34.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings35.Appearance = appearance106;
			summarySettings35.DisplayFormat = "{0: #,##0.00}";
			summarySettings35.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings36.Appearance = appearance107;
			summarySettings36.DisplayFormat = "{0: #,##0.00}";
			summarySettings36.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings18,
																																															summarySettings19,
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
																																															summarySettings36});
			ultraGridBand5.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance108.ForeColor = System.Drawing.Color.Black;
			appearance108.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorLocal.DisplayLayout.Override.ActiveRowAppearance = appearance108;
			this.uGridPorLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance109.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorLocal.DisplayLayout.Override.CardAreaAppearance = appearance109;
			appearance110.ForeColor = System.Drawing.Color.Black;
			appearance110.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorLocal.DisplayLayout.Override.CellAppearance = appearance110;
			appearance111.ForeColor = System.Drawing.Color.Black;
			appearance111.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorLocal.DisplayLayout.Override.GroupByRowAppearance = appearance111;
			appearance112.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance112.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance112.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance112.FontData.BoldAsString = "True";
			appearance112.FontData.Name = "Arial";
			appearance112.FontData.SizeInPoints = 8.5F;
			appearance112.ForeColor = System.Drawing.Color.White;
			appearance112.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorLocal.DisplayLayout.Override.HeaderAppearance = appearance112;
			this.uGridPorLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance113.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance113.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance113.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorLocal.DisplayLayout.Override.RowAlternateAppearance = appearance113;
			appearance114.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance114.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance114.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorLocal.DisplayLayout.Override.RowSelectorAppearance = appearance114;
			appearance115.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance115.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance115.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorLocal.DisplayLayout.Override.SelectedRowAppearance = appearance115;
			this.uGridPorLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorLocal.Location = new System.Drawing.Point(256, 8);
			this.uGridPorLocal.Name = "uGridPorLocal";
			this.uGridPorLocal.Size = new System.Drawing.Size(984, 704);
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
			appearance116.BackColor = System.Drawing.Color.White;
			this.uGridPorDias.DisplayLayout.Appearance = appearance116;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Width = 74;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance117;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Width = 80;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance118;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Width = 80;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance119;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.VisiblePosition = 3;
			ultraGridColumn59.Width = 80;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance120.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance120;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.VisiblePosition = 4;
			ultraGridColumn60.Width = 80;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance121.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance121;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 5;
			ultraGridColumn61.Width = 80;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance122.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance122;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "T/C";
			ultraGridColumn62.Header.VisiblePosition = 6;
			ultraGridColumn62.Width = 80;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance123.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance123;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.Caption = "N/C";
			ultraGridColumn63.Header.VisiblePosition = 7;
			ultraGridColumn63.Width = 80;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance124.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance124;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.VisiblePosition = 8;
			ultraGridColumn64.Width = 80;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance125.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance125;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.Caption = "Transf";
			ultraGridColumn65.Header.VisiblePosition = 10;
			ultraGridColumn65.Width = 80;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance126.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance126;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 9;
			ultraGridColumn66.Width = 80;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance127.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance127;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.Caption = "C/P";
			ultraGridColumn67.Header.VisiblePosition = 11;
			ultraGridColumn67.Width = 80;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance128.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance128;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.Caption = "Ret";
			ultraGridColumn68.Header.VisiblePosition = 12;
			ultraGridColumn68.Width = 80;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance129.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance129;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.Caption = "T/D";
			ultraGridColumn69.Header.VisiblePosition = 13;
			ultraGridColumn69.Width = 80;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance130.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance130;
			ultraGridColumn70.Format = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 14;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance131.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance131;
			ultraGridColumn71.Format = "#,##0.00";
			ultraGridColumn71.Header.VisiblePosition = 15;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance132.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance132;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.VisiblePosition = 16;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance133.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance133;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 17;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance134.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance134;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 18;
			ultraGridColumn74.Width = 80;
			ultraGridColumn75.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance135.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance135;
			ultraGridColumn75.Format = "#,##0.00";
			ultraGridColumn75.Header.VisiblePosition = 19;
			ultraGridColumn75.Width = 80;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn75});
			appearance136.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings37.Appearance = appearance136;
			summarySettings37.DisplayFormat = "{0: #,##0.00}";
			summarySettings37.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance137.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings38.Appearance = appearance137;
			summarySettings38.DisplayFormat = "{0: #,##0.00}";
			summarySettings38.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings39.Appearance = appearance138;
			summarySettings39.DisplayFormat = "{0: #,##0.00}";
			summarySettings39.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance139.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings40.Appearance = appearance139;
			summarySettings40.DisplayFormat = "{0: #,##0.00}";
			summarySettings40.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance140.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings41.Appearance = appearance140;
			summarySettings41.DisplayFormat = "{0: #,##0.00}";
			summarySettings41.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance141.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings42.Appearance = appearance141;
			summarySettings42.DisplayFormat = "{0: #,##0.00}";
			summarySettings42.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance142.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings43.Appearance = appearance142;
			summarySettings43.DisplayFormat = "{0: #,##0.00}";
			summarySettings43.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings44.Appearance = appearance143;
			summarySettings44.DisplayFormat = "{0: #,##0.00}";
			summarySettings44.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings45.Appearance = appearance144;
			summarySettings45.DisplayFormat = "{0: #,##0.00}";
			summarySettings45.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings46.Appearance = appearance145;
			summarySettings46.DisplayFormat = "{0: #,##0.00}";
			summarySettings46.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance146.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings47.Appearance = appearance146;
			summarySettings47.DisplayFormat = "{0: #,##0.00}";
			summarySettings47.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance147.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings48.Appearance = appearance147;
			summarySettings48.DisplayFormat = "{0: #,##0.00}";
			summarySettings48.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance148.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings49.Appearance = appearance148;
			summarySettings49.DisplayFormat = "{0: #,##0.00}";
			summarySettings49.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance149.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings50.Appearance = appearance149;
			summarySettings50.DisplayFormat = "{0: #,##0.00}";
			summarySettings50.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance150.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings51.Appearance = appearance150;
			summarySettings51.DisplayFormat = "{0: #,##0.00}";
			summarySettings51.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance151.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings52.Appearance = appearance151;
			summarySettings52.DisplayFormat = "{0: #,##0.00}";
			summarySettings52.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance152.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings53.Appearance = appearance152;
			summarySettings53.DisplayFormat = "{0: #,##0.00}";
			summarySettings53.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance153.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings54.Appearance = appearance153;
			summarySettings54.DisplayFormat = "{0: #,##0.00}";
			summarySettings54.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance154.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings55.Appearance = appearance154;
			summarySettings55.DisplayFormat = "{0: #,##0.00}";
			summarySettings55.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand6.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings37,
																																															summarySettings38,
																																															summarySettings39,
																																															summarySettings40,
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
																																															summarySettings55});
			ultraGridBand6.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorDias.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance155.ForeColor = System.Drawing.Color.Black;
			appearance155.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorDias.DisplayLayout.Override.ActiveRowAppearance = appearance155;
			this.uGridPorDias.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorDias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorDias.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance156.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorDias.DisplayLayout.Override.CardAreaAppearance = appearance156;
			appearance157.ForeColor = System.Drawing.Color.Black;
			appearance157.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorDias.DisplayLayout.Override.CellAppearance = appearance157;
			appearance158.ForeColor = System.Drawing.Color.Black;
			appearance158.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorDias.DisplayLayout.Override.GroupByRowAppearance = appearance158;
			appearance159.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance159.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance159.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance159.FontData.BoldAsString = "True";
			appearance159.FontData.Name = "Arial";
			appearance159.FontData.SizeInPoints = 8.5F;
			appearance159.ForeColor = System.Drawing.Color.White;
			appearance159.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorDias.DisplayLayout.Override.HeaderAppearance = appearance159;
			this.uGridPorDias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance160.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance160.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance160.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorDias.DisplayLayout.Override.RowAlternateAppearance = appearance160;
			appearance161.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance161.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance161.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorDias.DisplayLayout.Override.RowSelectorAppearance = appearance161;
			appearance162.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance162.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance162.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorDias.DisplayLayout.Override.SelectedRowAppearance = appearance162;
			this.uGridPorDias.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorDias.Location = new System.Drawing.Point(256, 8);
			this.uGridPorDias.Name = "uGridPorDias";
			this.uGridPorDias.Size = new System.Drawing.Size(984, 704);
			this.uGridPorDias.TabIndex = 656;
			this.uGridPorDias.Visible = false;
			this.uGridPorDias.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridPorDias_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn53.DataType = typeof(System.DateTime);
			ultraDataColumn53.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn54.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn55.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn56.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn57.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn72});
			// 
			// uGridPorVendedor
			// 
			this.uGridPorVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorVendedor.DataSource = this.ultraDataSource6;
			appearance163.BackColor = System.Drawing.Color.White;
			this.uGridPorVendedor.DisplayLayout.Appearance = appearance163;
			ultraGridColumn76.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn76.Header.Caption = "Vendedor";
			ultraGridColumn76.Header.VisiblePosition = 0;
			ultraGridColumn76.Width = 240;
			ultraGridColumn77.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance164.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn77.CellAppearance = appearance164;
			ultraGridColumn77.Format = "#,##0.00";
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridColumn77.Width = 80;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance165.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn78.CellAppearance = appearance165;
			ultraGridColumn78.Format = "#,##0.00";
			ultraGridColumn78.Header.VisiblePosition = 2;
			ultraGridColumn78.Width = 80;
			ultraGridColumn79.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance166.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance166;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.VisiblePosition = 3;
			ultraGridColumn79.Width = 80;
			ultraGridColumn80.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance167.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn80.CellAppearance = appearance167;
			ultraGridColumn80.Format = "#,##0.00";
			ultraGridColumn80.Header.VisiblePosition = 4;
			ultraGridColumn80.Width = 80;
			ultraGridColumn81.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance168.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn81.CellAppearance = appearance168;
			ultraGridColumn81.Format = "#,##0.00";
			ultraGridColumn81.Header.VisiblePosition = 5;
			ultraGridColumn81.Width = 80;
			ultraGridColumn82.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance169.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance169;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.Caption = "T/C";
			ultraGridColumn82.Header.VisiblePosition = 6;
			ultraGridColumn82.Width = 80;
			ultraGridColumn83.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance170.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn83.CellAppearance = appearance170;
			ultraGridColumn83.Format = "#,##0.00";
			ultraGridColumn83.Header.Caption = "N/C";
			ultraGridColumn83.Header.VisiblePosition = 7;
			ultraGridColumn83.Width = 80;
			ultraGridColumn84.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance171.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance171;
			ultraGridColumn84.Format = "#,##0.00";
			ultraGridColumn84.Header.VisiblePosition = 8;
			ultraGridColumn84.Width = 80;
			ultraGridColumn85.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance172.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance172;
			ultraGridColumn85.Format = "#,##0.00";
			ultraGridColumn85.Header.Caption = "Transf";
			ultraGridColumn85.Header.VisiblePosition = 10;
			ultraGridColumn85.Width = 80;
			ultraGridColumn86.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance173.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn86.CellAppearance = appearance173;
			ultraGridColumn86.Format = "#,##0.00";
			ultraGridColumn86.Header.VisiblePosition = 9;
			ultraGridColumn86.Width = 80;
			ultraGridColumn87.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance174.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn87.CellAppearance = appearance174;
			ultraGridColumn87.Format = "#,##0.00";
			ultraGridColumn87.Header.Caption = "C/P";
			ultraGridColumn87.Header.VisiblePosition = 11;
			ultraGridColumn87.Width = 80;
			ultraGridColumn88.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance175.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance175;
			ultraGridColumn88.Format = "#,##0.00";
			ultraGridColumn88.Header.Caption = "Ret";
			ultraGridColumn88.Header.VisiblePosition = 12;
			ultraGridColumn88.Width = 80;
			ultraGridColumn89.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance176.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance176;
			ultraGridColumn89.Format = "#,##0.00";
			ultraGridColumn89.Header.Caption = "T/D";
			ultraGridColumn89.Header.VisiblePosition = 13;
			ultraGridColumn89.Width = 80;
			ultraGridColumn90.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance177.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn90.CellAppearance = appearance177;
			ultraGridColumn90.Format = "#,##0.00";
			ultraGridColumn90.Header.VisiblePosition = 14;
			ultraGridColumn91.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance178.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn91.CellAppearance = appearance178;
			ultraGridColumn91.Format = "#,##0.00";
			ultraGridColumn91.Header.VisiblePosition = 15;
			ultraGridColumn92.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance179.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn92.CellAppearance = appearance179;
			ultraGridColumn92.Format = "#,##0.00";
			ultraGridColumn92.Header.VisiblePosition = 16;
			ultraGridColumn93.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance180.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn93.CellAppearance = appearance180;
			ultraGridColumn93.Format = "#,##0.00";
			ultraGridColumn93.Header.VisiblePosition = 17;
			ultraGridColumn94.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance181.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn94.CellAppearance = appearance181;
			ultraGridColumn94.Format = "#,##0.00";
			ultraGridColumn94.Header.VisiblePosition = 18;
			ultraGridColumn94.Width = 80;
			ultraGridColumn95.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance182.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance182;
			ultraGridColumn95.Format = "#,##0.00";
			ultraGridColumn95.Header.VisiblePosition = 19;
			ultraGridColumn95.Width = 80;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95});
			appearance183.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings56.Appearance = appearance183;
			summarySettings56.DisplayFormat = "{0: #,##0.00}";
			summarySettings56.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance184.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings57.Appearance = appearance184;
			summarySettings57.DisplayFormat = "{0: #,##0.00}";
			summarySettings57.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance185.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings58.Appearance = appearance185;
			summarySettings58.DisplayFormat = "{0: #,##0.00}";
			summarySettings58.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance186.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings59.Appearance = appearance186;
			summarySettings59.DisplayFormat = "{0: #,##0.00}";
			summarySettings59.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance187.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings60.Appearance = appearance187;
			summarySettings60.DisplayFormat = "{0: #,##0.00}";
			summarySettings60.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance188.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings61.Appearance = appearance188;
			summarySettings61.DisplayFormat = "{0: #,##0.00}";
			summarySettings61.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance189.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings62.Appearance = appearance189;
			summarySettings62.DisplayFormat = "{0: #,##0.00}";
			summarySettings62.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance190.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings63.Appearance = appearance190;
			summarySettings63.DisplayFormat = "{0: #,##0.00}";
			summarySettings63.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance191.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings64.Appearance = appearance191;
			summarySettings64.DisplayFormat = "{0: #,##0.00}";
			summarySettings64.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance192.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings65.Appearance = appearance192;
			summarySettings65.DisplayFormat = "{0: #,##0.00}";
			summarySettings65.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance193.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings66.Appearance = appearance193;
			summarySettings66.DisplayFormat = "{0: #,##0.00}";
			summarySettings66.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance194.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings67.Appearance = appearance194;
			summarySettings67.DisplayFormat = "{0: #,##0.00}";
			summarySettings67.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance195.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings68.Appearance = appearance195;
			summarySettings68.DisplayFormat = "{0: #,##0.00}";
			summarySettings68.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance196.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings69.Appearance = appearance196;
			summarySettings69.DisplayFormat = "{0: #,##0.00}";
			summarySettings69.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance197.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings70.Appearance = appearance197;
			summarySettings70.DisplayFormat = "{0: #,##0.00}";
			summarySettings70.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance198.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings71.Appearance = appearance198;
			summarySettings71.DisplayFormat = "{0: #,##0.00}";
			summarySettings71.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance199.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings72.Appearance = appearance199;
			summarySettings72.DisplayFormat = "{0: #,##0.00}";
			summarySettings72.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance200.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings73.Appearance = appearance200;
			summarySettings73.DisplayFormat = "{0: #,##0.00}";
			summarySettings73.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance201.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings74.Appearance = appearance201;
			summarySettings74.DisplayFormat = "{0: #,##0.00}";
			summarySettings74.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand7.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings56,
																																															summarySettings57,
																																															summarySettings58,
																																															summarySettings59,
																																															summarySettings60,
																																															summarySettings61,
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
																																															summarySettings74});
			ultraGridBand7.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance202.ForeColor = System.Drawing.Color.Black;
			appearance202.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorVendedor.DisplayLayout.Override.ActiveRowAppearance = appearance202;
			this.uGridPorVendedor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorVendedor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorVendedor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance203.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorVendedor.DisplayLayout.Override.CardAreaAppearance = appearance203;
			appearance204.ForeColor = System.Drawing.Color.Black;
			appearance204.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorVendedor.DisplayLayout.Override.CellAppearance = appearance204;
			appearance205.ForeColor = System.Drawing.Color.Black;
			appearance205.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorVendedor.DisplayLayout.Override.GroupByRowAppearance = appearance205;
			appearance206.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance206.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance206.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance206.FontData.BoldAsString = "True";
			appearance206.FontData.Name = "Arial";
			appearance206.FontData.SizeInPoints = 8.5F;
			appearance206.ForeColor = System.Drawing.Color.White;
			appearance206.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorVendedor.DisplayLayout.Override.HeaderAppearance = appearance206;
			this.uGridPorVendedor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance207.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance207.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance207.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorVendedor.DisplayLayout.Override.RowAlternateAppearance = appearance207;
			appearance208.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance208.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance208.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorVendedor.DisplayLayout.Override.RowSelectorAppearance = appearance208;
			appearance209.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance209.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance209.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorVendedor.DisplayLayout.Override.SelectedRowAppearance = appearance209;
			this.uGridPorVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorVendedor.Location = new System.Drawing.Point(256, 8);
			this.uGridPorVendedor.Name = "uGridPorVendedor";
			this.uGridPorVendedor.Size = new System.Drawing.Size(984, 704);
			this.uGridPorVendedor.TabIndex = 657;
			this.uGridPorVendedor.Visible = false;
			// 
			// ultraDataSource6
			// 
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
			ultraDataColumn77.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn78.DataType = typeof(System.Decimal);
			ultraDataColumn78.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn79.DataType = typeof(System.Decimal);
			ultraDataColumn79.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn90,
																																 ultraDataColumn91,
																																 ultraDataColumn92});
			// 
			// uGridPorlocalVendedor
			// 
			this.uGridPorlocalVendedor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPorlocalVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPorlocalVendedor.DataSource = this.ultraDataSource7;
			appearance210.BackColor = System.Drawing.Color.White;
			this.uGridPorlocalVendedor.DisplayLayout.Appearance = appearance210;
			ultraGridColumn96.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn96.Header.VisiblePosition = 0;
			ultraGridColumn96.Width = 49;
			ultraGridColumn97.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn97.Header.VisiblePosition = 1;
			ultraGridColumn97.Width = 200;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance211.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn98.CellAppearance = appearance211;
			ultraGridColumn98.Format = "#,##0.00";
			ultraGridColumn98.Header.VisiblePosition = 2;
			ultraGridColumn98.Width = 80;
			ultraGridColumn99.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance212.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn99.CellAppearance = appearance212;
			ultraGridColumn99.Format = "#,##0.00";
			ultraGridColumn99.Header.VisiblePosition = 3;
			ultraGridColumn99.Width = 80;
			ultraGridColumn100.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance213.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance213;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 4;
			ultraGridColumn100.Width = 80;
			ultraGridColumn101.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance214.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance214;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 5;
			ultraGridColumn101.Width = 80;
			ultraGridColumn102.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance215.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance215;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 6;
			ultraGridColumn102.Width = 80;
			ultraGridColumn103.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance216.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn103.CellAppearance = appearance216;
			ultraGridColumn103.Format = "#,##0.00";
			ultraGridColumn103.Header.Caption = "T/C";
			ultraGridColumn103.Header.VisiblePosition = 7;
			ultraGridColumn103.Width = 80;
			ultraGridColumn104.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance217.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn104.CellAppearance = appearance217;
			ultraGridColumn104.Format = "#,##0.00";
			ultraGridColumn104.Header.Caption = "N/C";
			ultraGridColumn104.Header.VisiblePosition = 8;
			ultraGridColumn104.Width = 80;
			ultraGridColumn105.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance218.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn105.CellAppearance = appearance218;
			ultraGridColumn105.Format = "#,##0.00";
			ultraGridColumn105.Header.VisiblePosition = 9;
			ultraGridColumn105.Width = 80;
			ultraGridColumn106.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance219.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn106.CellAppearance = appearance219;
			ultraGridColumn106.Format = "#,##0.00";
			ultraGridColumn106.Header.Caption = "Transf";
			ultraGridColumn106.Header.VisiblePosition = 11;
			ultraGridColumn106.Width = 80;
			ultraGridColumn107.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance220.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn107.CellAppearance = appearance220;
			ultraGridColumn107.Format = "#,##0.00";
			ultraGridColumn107.Header.VisiblePosition = 10;
			ultraGridColumn107.Width = 80;
			ultraGridColumn108.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance221.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn108.CellAppearance = appearance221;
			ultraGridColumn108.Format = "#,##0.00";
			ultraGridColumn108.Header.Caption = "C/P";
			ultraGridColumn108.Header.VisiblePosition = 12;
			ultraGridColumn108.Width = 80;
			ultraGridColumn109.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance222.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn109.CellAppearance = appearance222;
			ultraGridColumn109.Format = "#,##0.00";
			ultraGridColumn109.Header.Caption = "Ret";
			ultraGridColumn109.Header.VisiblePosition = 13;
			ultraGridColumn109.Width = 80;
			ultraGridColumn110.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance223.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn110.CellAppearance = appearance223;
			ultraGridColumn110.Format = "#,##0.00";
			ultraGridColumn110.Header.Caption = "T/D";
			ultraGridColumn110.Header.VisiblePosition = 14;
			ultraGridColumn110.Width = 80;
			ultraGridColumn111.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance224.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn111.CellAppearance = appearance224;
			ultraGridColumn111.Format = "#,##0.00";
			ultraGridColumn111.Header.VisiblePosition = 15;
			ultraGridColumn112.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance225.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn112.CellAppearance = appearance225;
			ultraGridColumn112.Format = "#,##0.00";
			ultraGridColumn112.Header.VisiblePosition = 16;
			ultraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance226.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn113.CellAppearance = appearance226;
			ultraGridColumn113.Format = "#,##0.00";
			ultraGridColumn113.Header.VisiblePosition = 17;
			ultraGridColumn114.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance227.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn114.CellAppearance = appearance227;
			ultraGridColumn114.Format = "#,##0.00";
			ultraGridColumn114.Header.VisiblePosition = 18;
			ultraGridColumn115.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance228.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn115.CellAppearance = appearance228;
			ultraGridColumn115.Format = "#,##0.00";
			ultraGridColumn115.Header.VisiblePosition = 19;
			ultraGridColumn115.Width = 80;
			ultraGridColumn116.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance229.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn116.CellAppearance = appearance229;
			ultraGridColumn116.Format = "#,##0.00";
			ultraGridColumn116.Header.VisiblePosition = 20;
			ultraGridColumn116.Width = 80;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
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
																										 ultraGridColumn116});
			appearance230.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings75.Appearance = appearance230;
			summarySettings75.DisplayFormat = "{0: #,##0.00}";
			summarySettings75.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance231.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings76.Appearance = appearance231;
			summarySettings76.DisplayFormat = "{0: #,##0.00}";
			summarySettings76.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance232.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings77.Appearance = appearance232;
			summarySettings77.DisplayFormat = "{0: #,##0.00}";
			summarySettings77.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance233.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings78.Appearance = appearance233;
			summarySettings78.DisplayFormat = "{0: #,##0.00}";
			summarySettings78.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance234.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings79.Appearance = appearance234;
			summarySettings79.DisplayFormat = "{0: #,##0.00}";
			summarySettings79.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance235.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings80.Appearance = appearance235;
			summarySettings80.DisplayFormat = "{0: #,##0.00}";
			summarySettings80.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance236.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings81.Appearance = appearance236;
			summarySettings81.DisplayFormat = "{0: #,##0.00}";
			summarySettings81.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance237.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings82.Appearance = appearance237;
			summarySettings82.DisplayFormat = "{0: #,##0.00}";
			summarySettings82.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance238.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings83.Appearance = appearance238;
			summarySettings83.DisplayFormat = "{0: #,##0.00}";
			summarySettings83.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance239.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings84.Appearance = appearance239;
			summarySettings84.DisplayFormat = "{0: #,##0.00}";
			summarySettings84.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance240.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings85.Appearance = appearance240;
			summarySettings85.DisplayFormat = "{0: #,##0.00}";
			summarySettings85.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance241.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings86.Appearance = appearance241;
			summarySettings86.DisplayFormat = "{0: #,##0.00}";
			summarySettings86.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance242.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings87.Appearance = appearance242;
			summarySettings87.DisplayFormat = "{0: #,##0.00}";
			summarySettings87.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance243.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings88.Appearance = appearance243;
			summarySettings88.DisplayFormat = "{0: #,##0.00}";
			summarySettings88.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance244.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings89.Appearance = appearance244;
			summarySettings89.DisplayFormat = "{0: #,##0.00}";
			summarySettings89.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance245.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings90.Appearance = appearance245;
			summarySettings90.DisplayFormat = "{0: #,##0.00}";
			summarySettings90.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance246.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings91.Appearance = appearance246;
			summarySettings91.DisplayFormat = "{0: #,##0.00}";
			summarySettings91.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance247.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings92.Appearance = appearance247;
			summarySettings92.DisplayFormat = "{0: #,##0.00}";
			summarySettings92.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance248.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings93.Appearance = appearance248;
			summarySettings93.DisplayFormat = "{0: #,##0.00}";
			summarySettings93.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand8.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings75,
																																															summarySettings76,
																																															summarySettings77,
																																															summarySettings78,
																																															summarySettings79,
																																															summarySettings80,
																																															summarySettings81,
																																															summarySettings82,
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
																																															summarySettings93});
			ultraGridBand8.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridPorlocalVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.uGridPorlocalVendedor.DisplayLayout.GroupByBox.Hidden = true;
			appearance249.ForeColor = System.Drawing.Color.Black;
			appearance249.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorlocalVendedor.DisplayLayout.Override.ActiveRowAppearance = appearance249;
			this.uGridPorlocalVendedor.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridPorlocalVendedor.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPorlocalVendedor.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance250.BackColor = System.Drawing.Color.Transparent;
			this.uGridPorlocalVendedor.DisplayLayout.Override.CardAreaAppearance = appearance250;
			appearance251.ForeColor = System.Drawing.Color.Black;
			appearance251.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorlocalVendedor.DisplayLayout.Override.CellAppearance = appearance251;
			appearance252.ForeColor = System.Drawing.Color.Black;
			appearance252.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPorlocalVendedor.DisplayLayout.Override.GroupByRowAppearance = appearance252;
			appearance253.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance253.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance253.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance253.FontData.BoldAsString = "True";
			appearance253.FontData.Name = "Arial";
			appearance253.FontData.SizeInPoints = 8.5F;
			appearance253.ForeColor = System.Drawing.Color.White;
			appearance253.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPorlocalVendedor.DisplayLayout.Override.HeaderAppearance = appearance253;
			this.uGridPorlocalVendedor.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance254.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance254.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance254.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorlocalVendedor.DisplayLayout.Override.RowAlternateAppearance = appearance254;
			appearance255.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance255.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance255.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorlocalVendedor.DisplayLayout.Override.RowSelectorAppearance = appearance255;
			appearance256.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance256.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance256.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPorlocalVendedor.DisplayLayout.Override.SelectedRowAppearance = appearance256;
			this.uGridPorlocalVendedor.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridPorlocalVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPorlocalVendedor.Location = new System.Drawing.Point(256, 8);
			this.uGridPorlocalVendedor.Name = "uGridPorlocalVendedor";
			this.uGridPorlocalVendedor.Size = new System.Drawing.Size(984, 704);
			this.uGridPorlocalVendedor.TabIndex = 658;
			this.uGridPorlocalVendedor.Visible = false;
			// 
			// ultraDataSource7
			// 
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
			ultraDataColumn99.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn100.DataType = typeof(System.Decimal);
			ultraDataColumn100.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn101.DataType = typeof(System.Decimal);
			ultraDataColumn101.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn102.DataType = typeof(System.Decimal);
			ultraDataColumn102.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
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
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn93,
																																 ultraDataColumn94,
																																 ultraDataColumn95,
																																 ultraDataColumn96,
																																 ultraDataColumn97,
																																 ultraDataColumn98,
																																 ultraDataColumn99,
																																 ultraDataColumn100,
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
																																 ultraDataColumn113});
			// 
			// uGridAgrupadoPorDias
			// 
			this.uGridAgrupadoPorDias.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridAgrupadoPorDias.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridAgrupadoPorDias.DataSource = this.ultraDataSource8;
			appearance257.BackColor = System.Drawing.Color.White;
			this.uGridAgrupadoPorDias.DisplayLayout.Appearance = appearance257;
			ultraGridColumn117.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn117.Header.VisiblePosition = 0;
			ultraGridColumn118.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn118.Header.VisiblePosition = 1;
			ultraGridColumn118.Width = 74;
			ultraGridColumn119.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance258.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn119.CellAppearance = appearance258;
			ultraGridColumn119.Format = "#,##0.00";
			ultraGridColumn119.Header.VisiblePosition = 2;
			ultraGridColumn119.Width = 80;
			ultraGridColumn120.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance259.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn120.CellAppearance = appearance259;
			ultraGridColumn120.Format = "#,##0.00";
			ultraGridColumn120.Header.VisiblePosition = 3;
			ultraGridColumn120.Width = 80;
			ultraGridColumn121.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance260.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn121.CellAppearance = appearance260;
			ultraGridColumn121.Format = "#,##0.00";
			ultraGridColumn121.Header.VisiblePosition = 4;
			ultraGridColumn121.Width = 80;
			ultraGridColumn122.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance261.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn122.CellAppearance = appearance261;
			ultraGridColumn122.Format = "#,##0.00";
			ultraGridColumn122.Header.VisiblePosition = 5;
			ultraGridColumn122.Width = 80;
			ultraGridColumn123.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance262.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn123.CellAppearance = appearance262;
			ultraGridColumn123.Format = "#,##0.00";
			ultraGridColumn123.Header.VisiblePosition = 6;
			ultraGridColumn123.Width = 80;
			ultraGridColumn124.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance263.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn124.CellAppearance = appearance263;
			ultraGridColumn124.Format = "#,##0.00";
			ultraGridColumn124.Header.Caption = "T/C";
			ultraGridColumn124.Header.VisiblePosition = 7;
			ultraGridColumn124.Width = 80;
			ultraGridColumn125.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance264.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn125.CellAppearance = appearance264;
			ultraGridColumn125.Format = "#,##0.00";
			ultraGridColumn125.Header.Caption = "N/C";
			ultraGridColumn125.Header.VisiblePosition = 8;
			ultraGridColumn125.Width = 80;
			ultraGridColumn126.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance265.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn126.CellAppearance = appearance265;
			ultraGridColumn126.Format = "#,##0.00";
			ultraGridColumn126.Header.VisiblePosition = 9;
			ultraGridColumn126.Width = 80;
			ultraGridColumn127.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance266.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn127.CellAppearance = appearance266;
			ultraGridColumn127.Format = "#,##0.00";
			ultraGridColumn127.Header.Caption = "Transf";
			ultraGridColumn127.Header.VisiblePosition = 11;
			ultraGridColumn127.Width = 80;
			ultraGridColumn128.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance267.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn128.CellAppearance = appearance267;
			ultraGridColumn128.Format = "#,##0.00";
			ultraGridColumn128.Header.VisiblePosition = 10;
			ultraGridColumn128.Width = 80;
			ultraGridColumn129.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance268.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn129.CellAppearance = appearance268;
			ultraGridColumn129.Format = "#,##0.00";
			ultraGridColumn129.Header.Caption = "C/P";
			ultraGridColumn129.Header.VisiblePosition = 12;
			ultraGridColumn129.Width = 80;
			ultraGridColumn130.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance269.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn130.CellAppearance = appearance269;
			ultraGridColumn130.Format = "#,##0.00";
			ultraGridColumn130.Header.Caption = "Ret";
			ultraGridColumn130.Header.VisiblePosition = 13;
			ultraGridColumn130.Width = 80;
			ultraGridColumn131.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance270.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn131.CellAppearance = appearance270;
			ultraGridColumn131.Format = "#,##0.00";
			ultraGridColumn131.Header.Caption = "T/D";
			ultraGridColumn131.Header.VisiblePosition = 14;
			ultraGridColumn131.Width = 80;
			ultraGridColumn132.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance271.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn132.CellAppearance = appearance271;
			ultraGridColumn132.Format = "#,##0.00";
			ultraGridColumn132.Header.VisiblePosition = 15;
			ultraGridColumn133.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance272.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn133.CellAppearance = appearance272;
			ultraGridColumn133.Format = "#,##0.00";
			ultraGridColumn133.Header.VisiblePosition = 16;
			ultraGridColumn134.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance273.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn134.CellAppearance = appearance273;
			ultraGridColumn134.Format = "#,##0.00";
			ultraGridColumn134.Header.VisiblePosition = 17;
			ultraGridColumn135.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance274.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn135.CellAppearance = appearance274;
			ultraGridColumn135.Format = "#,##0.00";
			ultraGridColumn135.Header.VisiblePosition = 18;
			ultraGridColumn136.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance275.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn136.CellAppearance = appearance275;
			ultraGridColumn136.Format = "#,##0.00";
			ultraGridColumn136.Header.VisiblePosition = 19;
			ultraGridColumn136.Width = 80;
			ultraGridColumn137.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance276.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn137.CellAppearance = appearance276;
			ultraGridColumn137.Format = "#,##0.00";
			ultraGridColumn137.Header.VisiblePosition = 20;
			ultraGridColumn137.Width = 80;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
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
																										 ultraGridColumn137});
			appearance277.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings94.Appearance = appearance277;
			summarySettings94.DisplayFormat = "{0: #,##0.00}";
			summarySettings94.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance278.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings95.Appearance = appearance278;
			summarySettings95.DisplayFormat = "{0: #,##0.00}";
			summarySettings95.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance279.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings96.Appearance = appearance279;
			summarySettings96.DisplayFormat = "{0: #,##0.00}";
			summarySettings96.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance280.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings97.Appearance = appearance280;
			summarySettings97.DisplayFormat = "{0: #,##0.00}";
			summarySettings97.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance281.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings98.Appearance = appearance281;
			summarySettings98.DisplayFormat = "{0: #,##0.00}";
			summarySettings98.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance282.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings99.Appearance = appearance282;
			summarySettings99.DisplayFormat = "{0: #,##0.00}";
			summarySettings99.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance283.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings100.Appearance = appearance283;
			summarySettings100.DisplayFormat = "{0: #,##0.00}";
			summarySettings100.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance284.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings101.Appearance = appearance284;
			summarySettings101.DisplayFormat = "{0: #,##0.00}";
			summarySettings101.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance285.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings102.Appearance = appearance285;
			summarySettings102.DisplayFormat = "{0: #,##0.00}";
			summarySettings102.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance286.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings103.Appearance = appearance286;
			summarySettings103.DisplayFormat = "{0: #,##0.00}";
			summarySettings103.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance287.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings104.Appearance = appearance287;
			summarySettings104.DisplayFormat = "{0: #,##0.00}";
			summarySettings104.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance288.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings105.Appearance = appearance288;
			summarySettings105.DisplayFormat = "{0: #,##0.00}";
			summarySettings105.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance289.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings106.Appearance = appearance289;
			summarySettings106.DisplayFormat = "{0: #,##0.00}";
			summarySettings106.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance290.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings107.Appearance = appearance290;
			summarySettings107.DisplayFormat = "{0: #,##0.00}";
			summarySettings107.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance291.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings108.Appearance = appearance291;
			summarySettings108.DisplayFormat = "{0: #,##0.00}";
			summarySettings108.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance292.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings109.Appearance = appearance292;
			summarySettings109.DisplayFormat = "{0: #,##0.00}";
			summarySettings109.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance293.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings110.Appearance = appearance293;
			summarySettings110.DisplayFormat = "{0: #,##0.00}";
			summarySettings110.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance294.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings111.Appearance = appearance294;
			summarySettings111.DisplayFormat = "{0: #,##0.00}";
			summarySettings111.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance295.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings112.Appearance = appearance295;
			summarySettings112.DisplayFormat = "{0: #,##0.00}";
			summarySettings112.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand9.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings94,
																																															summarySettings95,
																																															summarySettings96,
																																															summarySettings97,
																																															summarySettings98,
																																															summarySettings99,
																																															summarySettings100,
																																															summarySettings101,
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
																																															summarySettings112});
			ultraGridBand9.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridAgrupadoPorDias.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.uGridAgrupadoPorDias.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridAgrupadoPorDias.DisplayLayout.GroupByBox.Prompt = " ";
			appearance296.ForeColor = System.Drawing.Color.Black;
			appearance296.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.ActiveRowAppearance = appearance296;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance297.BackColor = System.Drawing.Color.Transparent;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.CardAreaAppearance = appearance297;
			appearance298.ForeColor = System.Drawing.Color.Black;
			appearance298.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.CellAppearance = appearance298;
			appearance299.ForeColor = System.Drawing.Color.Black;
			appearance299.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.GroupByRowAppearance = appearance299;
			appearance300.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance300.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance300.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance300.FontData.BoldAsString = "True";
			appearance300.FontData.Name = "Arial";
			appearance300.FontData.SizeInPoints = 8.5F;
			appearance300.ForeColor = System.Drawing.Color.White;
			appearance300.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.HeaderAppearance = appearance300;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance301.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance301.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance301.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.RowAlternateAppearance = appearance301;
			appearance302.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance302.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance302.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.RowSelectorAppearance = appearance302;
			appearance303.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance303.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance303.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridAgrupadoPorDias.DisplayLayout.Override.SelectedRowAppearance = appearance303;
			this.uGridAgrupadoPorDias.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridAgrupadoPorDias.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridAgrupadoPorDias.Location = new System.Drawing.Point(256, 8);
			this.uGridAgrupadoPorDias.Name = "uGridAgrupadoPorDias";
			this.uGridAgrupadoPorDias.Size = new System.Drawing.Size(984, 704);
			this.uGridAgrupadoPorDias.TabIndex = 661;
			this.uGridAgrupadoPorDias.Visible = false;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn115.DataType = typeof(System.DateTime);
			ultraDataColumn115.DefaultValue = new System.DateTime(((long)(0)));
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
			ultraDataColumn122.DataType = typeof(System.Decimal);
			ultraDataColumn122.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			ultraDataColumn123.DataType = typeof(System.Decimal);
			ultraDataColumn123.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
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
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn114,
																																 ultraDataColumn115,
																																 ultraDataColumn116,
																																 ultraDataColumn117,
																																 ultraDataColumn118,
																																 ultraDataColumn119,
																																 ultraDataColumn120,
																																 ultraDataColumn121,
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
																																 ultraDataColumn134});
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
			appearance304.FontData.Name = "Tahoma";
			appearance304.ForeColorDisabled = System.Drawing.Color.Black;
			this.rdgReporte.Appearance = appearance304;
			this.rdgReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rdgReporte.CheckedIndex = 0;
			this.rdgReporte.ItemAppearance = appearance305;
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
			ultraDataColumn135.DataType = typeof(bool);
			ultraDataColumn135.DefaultValue = false;
			ultraDataColumn138.DataType = typeof(System.DateTime);
			ultraDataColumn138.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn145.DataType = typeof(System.Double);
			ultraDataColumn145.DefaultValue = 0;
			ultraDataColumn146.DataType = typeof(System.Double);
			ultraDataColumn146.DefaultValue = 0;
			ultraDataColumn147.DataType = typeof(System.Double);
			ultraDataColumn147.DefaultValue = 0;
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn135,
																																 ultraDataColumn136,
																																 ultraDataColumn137,
																																 ultraDataColumn138,
																																 ultraDataColumn139,
																																 ultraDataColumn140,
																																 ultraDataColumn141,
																																 ultraDataColumn142,
																																 ultraDataColumn143,
																																 ultraDataColumn144,
																																 ultraDataColumn145,
																																 ultraDataColumn146,
																																 ultraDataColumn147});
			// 
			// uGridHorizontalGex
			// 
			this.uGridHorizontalGex.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridHorizontalGex.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridHorizontalGex.DataSource = this.ultraDataSource9;
			appearance306.BackColor = System.Drawing.Color.White;
			this.uGridHorizontalGex.DisplayLayout.Appearance = appearance306;
			ultraGridColumn138.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance307.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn138.CellAppearance = appearance307;
			ultraGridColumn138.Format = "#,##0.00";
			ultraGridColumn138.Header.VisiblePosition = 0;
			ultraGridColumn138.Width = 20;
			ultraGridColumn139.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn139.Header.VisiblePosition = 1;
			ultraGridColumn139.Width = 40;
			ultraGridColumn140.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn140.Header.VisiblePosition = 2;
			ultraGridColumn140.Width = 90;
			ultraGridColumn141.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn141.Header.VisiblePosition = 3;
			ultraGridColumn141.Width = 75;
			ultraGridColumn142.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn142.Header.VisiblePosition = 4;
			ultraGridColumn142.Width = 60;
			ultraGridColumn143.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn143.Header.VisiblePosition = 5;
			ultraGridColumn143.Width = 70;
			ultraGridColumn144.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn144.Header.VisiblePosition = 6;
			ultraGridColumn144.Width = 100;
			ultraGridColumn145.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn145.Header.Caption = "Cliente";
			ultraGridColumn145.Header.VisiblePosition = 7;
			ultraGridColumn145.Width = 240;
			ultraGridColumn146.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn146.Header.VisiblePosition = 8;
			ultraGridColumn146.Width = 50;
			ultraGridColumn147.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn147.Header.VisiblePosition = 9;
			ultraGridColumn147.Width = 240;
			ultraGridColumn148.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance308.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn148.CellAppearance = appearance308;
			ultraGridColumn148.Format = "#,##0.00";
			ultraGridColumn148.Header.VisiblePosition = 10;
			ultraGridColumn148.Width = 80;
			ultraGridColumn149.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance309.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn149.CellAppearance = appearance309;
			ultraGridColumn149.Format = "#,##0.00";
			ultraGridColumn149.Header.VisiblePosition = 11;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn149.Width = 80;
			ultraGridColumn150.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance310.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn150.CellAppearance = appearance310;
			ultraGridColumn150.Format = "#,##0.00";
			ultraGridColumn150.Header.VisiblePosition = 12;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn150.Width = 80;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn138,
																											ultraGridColumn139,
																											ultraGridColumn140,
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143,
																											ultraGridColumn144,
																											ultraGridColumn145,
																											ultraGridColumn146,
																											ultraGridColumn147,
																											ultraGridColumn148,
																											ultraGridColumn149,
																											ultraGridColumn150});
			appearance311.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings113.Appearance = appearance311;
			summarySettings113.DisplayFormat = "{0: #,##0.00}";
			summarySettings113.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance312.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings114.Appearance = appearance312;
			summarySettings114.DisplayFormat = "{0: #,##0.00}";
			summarySettings114.Hidden = true;
			summarySettings114.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance313.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings115.Appearance = appearance313;
			summarySettings115.DisplayFormat = "{0: #,##0.00}";
			summarySettings115.Hidden = true;
			summarySettings115.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand10.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings113,
																																															 summarySettings114,
																																															 summarySettings115});
			ultraGridBand10.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridHorizontalGex.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			appearance314.ForeColor = System.Drawing.Color.Black;
			appearance314.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontalGex.DisplayLayout.Override.ActiveRowAppearance = appearance314;
			this.uGridHorizontalGex.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridHorizontalGex.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridHorizontalGex.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance315.BackColor = System.Drawing.Color.Transparent;
			this.uGridHorizontalGex.DisplayLayout.Override.CardAreaAppearance = appearance315;
			appearance316.ForeColor = System.Drawing.Color.Black;
			appearance316.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridHorizontalGex.DisplayLayout.Override.GroupByRowAppearance = appearance316;
			appearance317.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance317.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance317.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance317.FontData.BoldAsString = "True";
			appearance317.FontData.Name = "Arial";
			appearance317.FontData.SizeInPoints = 8.5F;
			appearance317.ForeColor = System.Drawing.Color.White;
			appearance317.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridHorizontalGex.DisplayLayout.Override.HeaderAppearance = appearance317;
			this.uGridHorizontalGex.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance318.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance318.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance318.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontalGex.DisplayLayout.Override.RowAlternateAppearance = appearance318;
			appearance319.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance319.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance319.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontalGex.DisplayLayout.Override.RowSelectorAppearance = appearance319;
			appearance320.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance320.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance320.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridHorizontalGex.DisplayLayout.Override.SelectedRowAppearance = appearance320;
			this.uGridHorizontalGex.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridHorizontalGex.Location = new System.Drawing.Point(256, 8);
			this.uGridHorizontalGex.Name = "uGridHorizontalGex";
			this.uGridHorizontalGex.Size = new System.Drawing.Size(984, 704);
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
			appearance321.BackColor = System.Drawing.Color.White;
			this.uGridGExPorLocal.DisplayLayout.Appearance = appearance321;
			this.uGridGExPorLocal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn151.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn151.Header.VisiblePosition = 0;
			ultraGridColumn151.Width = 300;
			ultraGridColumn152.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance322.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn152.CellAppearance = appearance322;
			ultraGridColumn152.Format = "#,##0.00";
			ultraGridColumn152.Header.VisiblePosition = 1;
			ultraGridColumn152.Width = 150;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn151,
																											ultraGridColumn152});
			appearance323.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings116.Appearance = appearance323;
			summarySettings116.DisplayFormat = "{0: #,##0.00}";
			summarySettings116.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand11.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings116});
			ultraGridBand11.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridGExPorLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			appearance324.ForeColor = System.Drawing.Color.Black;
			appearance324.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGExPorLocal.DisplayLayout.Override.ActiveRowAppearance = appearance324;
			this.uGridGExPorLocal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridGExPorLocal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridGExPorLocal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance325.BackColor = System.Drawing.Color.Transparent;
			this.uGridGExPorLocal.DisplayLayout.Override.CardAreaAppearance = appearance325;
			appearance326.ForeColor = System.Drawing.Color.Black;
			appearance326.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridGExPorLocal.DisplayLayout.Override.GroupByRowAppearance = appearance326;
			appearance327.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance327.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance327.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance327.FontData.BoldAsString = "True";
			appearance327.FontData.Name = "Arial";
			appearance327.FontData.SizeInPoints = 8.5F;
			appearance327.ForeColor = System.Drawing.Color.White;
			appearance327.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridGExPorLocal.DisplayLayout.Override.HeaderAppearance = appearance327;
			this.uGridGExPorLocal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance328.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance328.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance328.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGExPorLocal.DisplayLayout.Override.RowAlternateAppearance = appearance328;
			appearance329.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance329.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance329.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGExPorLocal.DisplayLayout.Override.RowSelectorAppearance = appearance329;
			appearance330.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance330.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance330.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridGExPorLocal.DisplayLayout.Override.SelectedRowAppearance = appearance330;
			this.uGridGExPorLocal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridGExPorLocal.Location = new System.Drawing.Point(256, 8);
			this.uGridGExPorLocal.Name = "uGridGExPorLocal";
			this.uGridGExPorLocal.Size = new System.Drawing.Size(471, 704);
			this.uGridGExPorLocal.TabIndex = 683;
			this.uGridGExPorLocal.Visible = false;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn149.DataType = typeof(System.Decimal);
			ultraDataColumn149.DefaultValue = new System.Decimal(new int[] {
																																			 0,
																																			 0,
																																			 0,
																																			 0});
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn148,
																																	ultraDataColumn149});
			// 
			// uGridDYD
			// 
			this.uGridDYD.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDYD.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDYD.DataSource = this.ultraDataSource11;
			appearance331.BackColor = System.Drawing.Color.White;
			this.uGridDYD.DisplayLayout.Appearance = appearance331;
			this.uGridDYD.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn153.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn153.Header.VisiblePosition = 0;
			ultraGridColumn153.Width = 120;
			ultraGridColumn154.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn154.Header.VisiblePosition = 1;
			ultraGridColumn154.Width = 188;
			ultraGridColumn155.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn155.Header.VisiblePosition = 2;
			ultraGridColumn155.Width = 75;
			ultraGridColumn156.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn156.Header.VisiblePosition = 3;
			ultraGridColumn156.Width = 80;
			ultraGridColumn157.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance332.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn157.CellAppearance = appearance332;
			ultraGridColumn157.Format = "#,##0.00";
			ultraGridColumn157.Header.Caption = "Valor D & D";
			ultraGridColumn157.Header.VisiblePosition = 4;
			ultraGridColumn157.Width = 100;
			ultraGridColumn158.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance333.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn158.CellAppearance = appearance333;
			ultraGridColumn158.Format = "#,##0.00";
			ultraGridColumn158.Header.Caption = "I.V.A. D & D";
			ultraGridColumn158.Header.VisiblePosition = 5;
			ultraGridColumn158.Width = 100;
			ultraGridColumn159.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance334.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn159.CellAppearance = appearance334;
			ultraGridColumn159.Format = "#,##0.00";
			ultraGridColumn159.Header.Caption = "D & D Con I.V.A.";
			ultraGridColumn159.Header.VisiblePosition = 6;
			ultraGridColumn159.Width = 100;
			ultraGridColumn160.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance335.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn160.CellAppearance = appearance335;
			ultraGridColumn160.Format = "#,##0.00";
			ultraGridColumn160.Header.Caption = "Total Factura";
			ultraGridColumn160.Header.VisiblePosition = 7;
			ultraGridColumn160.Width = 100;
			ultraGridColumn161.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance336.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn161.CellAppearance = appearance336;
			ultraGridColumn161.Format = "#,##0.00";
			ultraGridColumn161.Header.Caption = "Monto Credito";
			ultraGridColumn161.Header.VisiblePosition = 8;
			ultraGridColumn161.Width = 100;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn153,
																											ultraGridColumn154,
																											ultraGridColumn155,
																											ultraGridColumn156,
																											ultraGridColumn157,
																											ultraGridColumn158,
																											ultraGridColumn159,
																											ultraGridColumn160,
																											ultraGridColumn161});
			appearance337.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings117.Appearance = appearance337;
			summarySettings117.DisplayFormat = "{0: #,##0.00}";
			summarySettings117.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance338.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings118.Appearance = appearance338;
			summarySettings118.DisplayFormat = "{0: #,##0.00}";
			summarySettings118.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance339.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings119.Appearance = appearance339;
			summarySettings119.DisplayFormat = "{0: #,##0.00}";
			summarySettings119.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance340.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings120.Appearance = appearance340;
			summarySettings120.DisplayFormat = "{0: #,##0.00}";
			summarySettings120.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance341.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings121.Appearance = appearance341;
			summarySettings121.DisplayFormat = "{0: #,##0.00}";
			summarySettings121.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand12.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															 summarySettings117,
																																															 summarySettings118,
																																															 summarySettings119,
																																															 summarySettings120,
																																															 summarySettings121});
			ultraGridBand12.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridDYD.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			appearance342.ForeColor = System.Drawing.Color.Black;
			appearance342.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.ActiveRowAppearance = appearance342;
			this.uGridDYD.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDYD.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDYD.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance343.BackColor = System.Drawing.Color.Transparent;
			this.uGridDYD.DisplayLayout.Override.CardAreaAppearance = appearance343;
			appearance344.ForeColor = System.Drawing.Color.Black;
			appearance344.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.CellAppearance = appearance344;
			appearance345.ForeColor = System.Drawing.Color.Black;
			appearance345.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.GroupByRowAppearance = appearance345;
			appearance346.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance346.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance346.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance346.FontData.BoldAsString = "True";
			appearance346.FontData.Name = "Arial";
			appearance346.FontData.SizeInPoints = 8.5F;
			appearance346.ForeColor = System.Drawing.Color.White;
			appearance346.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDYD.DisplayLayout.Override.HeaderAppearance = appearance346;
			this.uGridDYD.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance347.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance347.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance347.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowAlternateAppearance = appearance347;
			appearance348.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance348.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance348.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowSelectorAppearance = appearance348;
			appearance349.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance349.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance349.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.SelectedRowAppearance = appearance349;
			this.uGridDYD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDYD.Location = new System.Drawing.Point(256, 7);
			this.uGridDYD.Name = "uGridDYD";
			this.uGridDYD.Size = new System.Drawing.Size(984, 704);
			this.uGridDYD.TabIndex = 684;
			this.uGridDYD.Visible = false;
			// 
			// ultraDataSource11
			// 
			ultraDataColumn152.DataType = typeof(System.DateTime);
			ultraDataColumn154.DataType = typeof(System.Decimal);
			ultraDataColumn155.DataType = typeof(System.Decimal);
			ultraDataColumn156.DataType = typeof(System.Decimal);
			ultraDataColumn157.DataType = typeof(System.Decimal);
			ultraDataColumn158.DataType = typeof(System.Decimal);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn150,
																																	ultraDataColumn151,
																																	ultraDataColumn152,
																																	ultraDataColumn153,
																																	ultraDataColumn154,
																																	ultraDataColumn155,
																																	ultraDataColumn156,
																																	ultraDataColumn157,
																																	ultraDataColumn158});
			// 
			// frmResumenDeVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1248, 718);
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
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridHorizontal);				

				this.Text = this.Text + " - " + this.uGridHorizontal.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 2)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorLocal '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridPorLocal.DataSource = this.ultraDataSource4;

				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPorLocal);

				this.Text = this.Text + " - " + this.uGridPorLocal.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 3)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorDias '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridPorDias.DataSource = this.ultraDataSource5;
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPorDias);

				this.Text = this.Text + " - " + this.uGridPorDias.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 4)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorVendedor '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridPorVendedor.DataSource = this.ultraDataSource6;
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPorVendedor);

				this.Text = this.Text + " - " + this.uGridPorVendedor.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 5)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorLocalVendedor '{0}', '{1}', '{2}', {3}, {4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor, 0);
				
				this.uGridPorlocalVendedor.DataSource = this.ultraDataSource7;
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridPorlocalVendedor);

				this.uGridPorlocalVendedor.DisplayLayout.Bands[0].SortedColumns.Add("Local", false, true);			
				this.uGridPorlocalVendedor.Rows.CollapseAll(true);

				this.Text = this.Text + " - " + this.uGridPorlocalVendedor.Rows.Count + " Registros Encontrados";

				return;
			}
			
			if ((int)this.rdgReporte.Value == 6)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorDiasAgrupados '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridAgrupadoPorDias.DataSource = this.ultraDataSource8;
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridAgrupadoPorDias);

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
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridHorizontal);
				
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
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridHorizontalGex);				

				this.Text = this.Text + " - " + this.uGridHorizontalGex.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 9)
			{				
				sSQL = string.Format("Exec ReporteDeVentasPorLocalGEx '{0}', '{1}', '{2}', {3}, {4}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor, 0);
				
				this.uGridGExPorLocal.DataSource = this.ultraDataSource10;
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridGExPorLocal);				

				this.Text = this.Text + " - " + this.uGridGExPorLocal.Rows.Count + " Registros Encontrados";

				return;
			}

			if ((int)this.rdgReporte.Value == 10)
			{				
				sSQL = string.Format("Exec ReporteDeVentasDyD '{0}', '{1}', '{2}', {3}",
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, idVendedor);
				
				this.uGridDYD.DataSource = this.ultraDataSource11;
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDYD);				

				this.Text = this.Text + " - " + this.uGridDYD.Rows.Count + " Registros Encontrados";

				return;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
		}
	}
}

