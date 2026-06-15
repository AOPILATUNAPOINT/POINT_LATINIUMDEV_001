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
	/// Descripción breve de frmReporteDataPuraCobranza.
	/// </summary>
	public class frmReporteDataPuraCobranza : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TbBasico;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDocumentos;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label lblContFacturas;
		private System.Windows.Forms.Label lblcontPagare;
		private System.Windows.Forms.Label lblContTablaAmortizacion;
		private System.Windows.Forms.Button btnExcelPagare;
		private System.Windows.Forms.Button btnConsultarPagare;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridPagare;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridTablaAmortizacion;
		private System.Windows.Forms.Button btnConsultaTablaAmortizacion;
		private System.Windows.Forms.Button btnExcelTablaAmortizacion;
		private System.Windows.Forms.Button btnconsultaFacturas;
		private System.Windows.Forms.Button btnExcelFacturas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAño;
		private System.Windows.Forms.Label lblPeriodo;
		private System.Windows.Forms.Label lblFactura;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteDataPuraCobranza()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteDataPuraCobranza));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 7, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuota");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pago");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pago", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pago", 4, true);
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 5, true);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 6, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 7, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuota");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vence");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total_Interes");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total_Capital");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mora");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Abono");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Retraso");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Docuento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total_Interes", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total_Interes", 8, true);
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total_Capital", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total_Capital", 9, true);
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Interes", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Interes", 10, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Capital", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Capital", 11, true);
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Mora", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Mora", 12, true);
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 13, true);
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Abono", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Abono", 14, true);
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Retraso", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Retraso", 15, true);
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vence");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total_Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total_Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Abono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Retraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Docuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.TbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblFactura = new System.Windows.Forms.Label();
			this.btnconsultaFacturas = new System.Windows.Forms.Button();
			this.btnExcelFacturas = new System.Windows.Forms.Button();
			this.lblContFacturas = new System.Windows.Forms.Label();
			this.uGridDocumentos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridPagare = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnConsultarPagare = new System.Windows.Forms.Button();
			this.btnExcelPagare = new System.Windows.Forms.Button();
			this.lblcontPagare = new System.Windows.Forms.Label();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnConsultaTablaAmortizacion = new System.Windows.Forms.Button();
			this.btnExcelTablaAmortizacion = new System.Windows.Forms.Button();
			this.uGridTablaAmortizacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblContTablaAmortizacion = new System.Windows.Forms.Label();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtAño = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPeriodo = new System.Windows.Forms.Label();
			this.TbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridPagare)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridTablaAmortizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).BeginInit();
			this.SuspendLayout();
			// 
			// TbBasico
			// 
			this.TbBasico.Controls.Add(this.lblFactura);
			this.TbBasico.Controls.Add(this.btnconsultaFacturas);
			this.TbBasico.Controls.Add(this.btnExcelFacturas);
			this.TbBasico.Controls.Add(this.lblContFacturas);
			this.TbBasico.Controls.Add(this.uGridDocumentos);
			this.TbBasico.Controls.Add(this.cmbCargos);
			this.TbBasico.Location = new System.Drawing.Point(-10000, -10000);
			this.TbBasico.Name = "TbBasico";
			this.TbBasico.Size = new System.Drawing.Size(972, 334);
			// 
			// lblFactura
			// 
			this.lblFactura.AutoSize = true;
			this.lblFactura.BackColor = System.Drawing.Color.Transparent;
			this.lblFactura.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFactura.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFactura.Location = new System.Drawing.Point(16, 8);
			this.lblFactura.Name = "lblFactura";
			this.lblFactura.Size = new System.Drawing.Size(0, 20);
			this.lblFactura.TabIndex = 777;
			this.lblFactura.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnconsultaFacturas
			// 
			this.btnconsultaFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnconsultaFacturas.BackColor = System.Drawing.SystemColors.Control;
			this.btnconsultaFacturas.CausesValidation = false;
			this.btnconsultaFacturas.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnconsultaFacturas.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnconsultaFacturas.Image = ((System.Drawing.Image)(resources.GetObject("btnconsultaFacturas.Image")));
			this.btnconsultaFacturas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnconsultaFacturas.Location = new System.Drawing.Point(800, 8);
			this.btnconsultaFacturas.Name = "btnconsultaFacturas";
			this.btnconsultaFacturas.Size = new System.Drawing.Size(78, 23);
			this.btnconsultaFacturas.TabIndex = 776;
			this.btnconsultaFacturas.Text = "Consultar";
			this.btnconsultaFacturas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnconsultaFacturas.Click += new System.EventHandler(this.btnconsultaFacturas_Click);
			// 
			// btnExcelFacturas
			// 
			this.btnExcelFacturas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcelFacturas.Enabled = false;
			this.btnExcelFacturas.Image = ((System.Drawing.Image)(resources.GetObject("btnExcelFacturas.Image")));
			this.btnExcelFacturas.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcelFacturas.Location = new System.Drawing.Point(888, 8);
			this.btnExcelFacturas.Name = "btnExcelFacturas";
			this.btnExcelFacturas.Size = new System.Drawing.Size(78, 23);
			this.btnExcelFacturas.TabIndex = 775;
			this.btnExcelFacturas.Text = "&Importar";
			this.btnExcelFacturas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcelFacturas.Click += new System.EventHandler(this.btnExcelFacturas_Click);
			// 
			// lblContFacturas
			// 
			this.lblContFacturas.AutoSize = true;
			this.lblContFacturas.BackColor = System.Drawing.Color.Transparent;
			this.lblContFacturas.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContFacturas.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContFacturas.Location = new System.Drawing.Point(8, 8);
			this.lblContFacturas.Name = "lblContFacturas";
			this.lblContFacturas.Size = new System.Drawing.Size(0, 20);
			this.lblContFacturas.TabIndex = 771;
			this.lblContFacturas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// uGridDocumentos
			// 
			this.uGridDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDocumentos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDocumentos.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDocumentos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 40;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 80;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 250;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance4;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridDocumentos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDocumentos.DisplayLayout.GroupByBox.Hidden = true;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.uGridDocumentos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDocumentos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDocumentos.DisplayLayout.Override.CellAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDocumentos.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.uGridDocumentos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDocumentos.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridDocumentos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDocumentos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDocumentos.Location = new System.Drawing.Point(8, 40);
			this.uGridDocumentos.Name = "uGridDocumentos";
			this.uGridDocumentos.Size = new System.Drawing.Size(960, 288);
			this.uGridDocumentos.TabIndex = 764;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 1;
			ultraGridColumn11.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(752, 448);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(24, 8);
			this.cmbCargos.TabIndex = 763;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.uGridPagare);
			this.ultraTabPageControl10.Controls.Add(this.btnConsultarPagare);
			this.ultraTabPageControl10.Controls.Add(this.btnExcelPagare);
			this.ultraTabPageControl10.Controls.Add(this.lblcontPagare);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(972, 334);
			// 
			// uGridPagare
			// 
			this.uGridPagare.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridPagare.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridPagare.DataSource = this.ultraDataSource1;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridPagare.DisplayLayout.Appearance = appearance13;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Width = 120;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 300;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance15;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance16;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance17;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance18;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			ultraGridBand3.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand3.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
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
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridPagare.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridPagare.DisplayLayout.GroupByBox.Hidden = true;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagare.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.uGridPagare.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridPagare.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.uGridPagare.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridPagare.DisplayLayout.Override.CellAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridPagare.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.uGridPagare.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagare.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagare.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridPagare.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridPagare.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridPagare.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridPagare.Location = new System.Drawing.Point(8, 40);
			this.uGridPagare.Name = "uGridPagare";
			this.uGridPagare.Size = new System.Drawing.Size(960, 288);
			this.uGridPagare.TabIndex = 775;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19});
			// 
			// btnConsultarPagare
			// 
			this.btnConsultarPagare.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultarPagare.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultarPagare.CausesValidation = false;
			this.btnConsultarPagare.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultarPagare.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultarPagare.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultarPagare.Image")));
			this.btnConsultarPagare.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultarPagare.Location = new System.Drawing.Point(792, 8);
			this.btnConsultarPagare.Name = "btnConsultarPagare";
			this.btnConsultarPagare.Size = new System.Drawing.Size(78, 23);
			this.btnConsultarPagare.TabIndex = 774;
			this.btnConsultarPagare.Text = "Consultar";
			this.btnConsultarPagare.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultarPagare.Click += new System.EventHandler(this.btnConsultarPagare_Click);
			// 
			// btnExcelPagare
			// 
			this.btnExcelPagare.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcelPagare.Enabled = false;
			this.btnExcelPagare.Image = ((System.Drawing.Image)(resources.GetObject("btnExcelPagare.Image")));
			this.btnExcelPagare.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcelPagare.Location = new System.Drawing.Point(880, 8);
			this.btnExcelPagare.Name = "btnExcelPagare";
			this.btnExcelPagare.Size = new System.Drawing.Size(78, 23);
			this.btnExcelPagare.TabIndex = 773;
			this.btnExcelPagare.Text = "&Importar";
			this.btnExcelPagare.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcelPagare.Click += new System.EventHandler(this.btnExcelPagare_Click);
			// 
			// lblcontPagare
			// 
			this.lblcontPagare.AutoSize = true;
			this.lblcontPagare.BackColor = System.Drawing.Color.Transparent;
			this.lblcontPagare.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblcontPagare.ForeColor = System.Drawing.Color.Firebrick;
			this.lblcontPagare.Location = new System.Drawing.Point(24, 8);
			this.lblcontPagare.Name = "lblcontPagare";
			this.lblcontPagare.Size = new System.Drawing.Size(0, 20);
			this.lblcontPagare.TabIndex = 770;
			this.lblcontPagare.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btnConsultaTablaAmortizacion);
			this.ultraTabPageControl1.Controls.Add(this.btnExcelTablaAmortizacion);
			this.ultraTabPageControl1.Controls.Add(this.uGridTablaAmortizacion);
			this.ultraTabPageControl1.Controls.Add(this.lblContTablaAmortizacion);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(972, 334);
			// 
			// btnConsultaTablaAmortizacion
			// 
			this.btnConsultaTablaAmortizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnConsultaTablaAmortizacion.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultaTablaAmortizacion.CausesValidation = false;
			this.btnConsultaTablaAmortizacion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultaTablaAmortizacion.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultaTablaAmortizacion.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultaTablaAmortizacion.Image")));
			this.btnConsultaTablaAmortizacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultaTablaAmortizacion.Location = new System.Drawing.Point(800, 8);
			this.btnConsultaTablaAmortizacion.Name = "btnConsultaTablaAmortizacion";
			this.btnConsultaTablaAmortizacion.Size = new System.Drawing.Size(78, 23);
			this.btnConsultaTablaAmortizacion.TabIndex = 778;
			this.btnConsultaTablaAmortizacion.Text = "Consultar";
			this.btnConsultaTablaAmortizacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultaTablaAmortizacion.Click += new System.EventHandler(this.btnConsultaTablaAmortizacion_Click);
			// 
			// btnExcelTablaAmortizacion
			// 
			this.btnExcelTablaAmortizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcelTablaAmortizacion.Enabled = false;
			this.btnExcelTablaAmortizacion.Image = ((System.Drawing.Image)(resources.GetObject("btnExcelTablaAmortizacion.Image")));
			this.btnExcelTablaAmortizacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcelTablaAmortizacion.Location = new System.Drawing.Point(888, 8);
			this.btnExcelTablaAmortizacion.Name = "btnExcelTablaAmortizacion";
			this.btnExcelTablaAmortizacion.Size = new System.Drawing.Size(78, 23);
			this.btnExcelTablaAmortizacion.TabIndex = 777;
			this.btnExcelTablaAmortizacion.Text = "&Importar";
			this.btnExcelTablaAmortizacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcelTablaAmortizacion.Click += new System.EventHandler(this.btnExcelTablaAmortizacion_Click);
			// 
			// uGridTablaAmortizacion
			// 
			this.uGridTablaAmortizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridTablaAmortizacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridTablaAmortizacion.DataSource = this.ultraDataSource2;
			appearance30.BackColor = System.Drawing.Color.White;
			this.uGridTablaAmortizacion.DisplayLayout.Appearance = appearance30;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn22.CellAppearance = appearance31;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 90;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn23.Header.VisiblePosition = 2;
			ultraGridColumn23.Width = 250;
			ultraGridColumn24.Header.VisiblePosition = 3;
			ultraGridColumn25.Header.VisiblePosition = 4;
			ultraGridColumn26.Header.VisiblePosition = 5;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellAppearance = appearance32;
			ultraGridColumn27.Header.VisiblePosition = 7;
			ultraGridColumn28.Header.VisiblePosition = 8;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance33;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 9;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance34;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 10;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance35;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.VisiblePosition = 11;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance36;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.VisiblePosition = 12;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance37;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 13;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance38;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 14;
			ultraGridColumn34.Width = 90;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance39;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.VisiblePosition = 15;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance40;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 16;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellAppearance = appearance41;
			ultraGridColumn37.Header.VisiblePosition = 17;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn39.Header.VisiblePosition = 6;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn39});
			appearance42.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance42;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance43;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance44;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance45;
			summarySettings9.DisplayFormat = "{0: #,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance46;
			summarySettings10.DisplayFormat = "{0: #,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance47.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance47;
			summarySettings11.DisplayFormat = "{0: #,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance48.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance48;
			summarySettings12.DisplayFormat = "{0: #,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance49.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance49;
			summarySettings13.DisplayFormat = "{0: #,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings14.Appearance = appearance50;
			summarySettings14.DisplayFormat = "{0: #,##0.00}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridTablaAmortizacion.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridTablaAmortizacion.DisplayLayout.GroupByBox.Hidden = true;
			appearance51.ForeColor = System.Drawing.Color.Black;
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTablaAmortizacion.DisplayLayout.Override.ActiveRowAppearance = appearance51;
			this.uGridTablaAmortizacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridTablaAmortizacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance52.BackColor = System.Drawing.Color.Transparent;
			this.uGridTablaAmortizacion.DisplayLayout.Override.CardAreaAppearance = appearance52;
			appearance53.ForeColor = System.Drawing.Color.Black;
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridTablaAmortizacion.DisplayLayout.Override.CellAppearance = appearance53;
			appearance54.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance54.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance54.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance54.FontData.BoldAsString = "True";
			appearance54.FontData.Name = "Arial";
			appearance54.FontData.SizeInPoints = 8F;
			appearance54.ForeColor = System.Drawing.Color.White;
			appearance54.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridTablaAmortizacion.DisplayLayout.Override.HeaderAppearance = appearance54;
			this.uGridTablaAmortizacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance55.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance55.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTablaAmortizacion.DisplayLayout.Override.RowAlternateAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTablaAmortizacion.DisplayLayout.Override.RowSelectorAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridTablaAmortizacion.DisplayLayout.Override.SelectedRowAppearance = appearance57;
			this.uGridTablaAmortizacion.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridTablaAmortizacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridTablaAmortizacion.Location = new System.Drawing.Point(6, 40);
			this.uGridTablaAmortizacion.Name = "uGridTablaAmortizacion";
			this.uGridTablaAmortizacion.Size = new System.Drawing.Size(960, 288);
			this.uGridTablaAmortizacion.TabIndex = 776;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn24.DataType = typeof(System.DateTime);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(System.DateTime);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn37});
			// 
			// lblContTablaAmortizacion
			// 
			this.lblContTablaAmortizacion.AutoSize = true;
			this.lblContTablaAmortizacion.BackColor = System.Drawing.Color.Transparent;
			this.lblContTablaAmortizacion.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContTablaAmortizacion.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContTablaAmortizacion.Location = new System.Drawing.Point(16, 8);
			this.lblContTablaAmortizacion.Name = "lblContTablaAmortizacion";
			this.lblContTablaAmortizacion.Size = new System.Drawing.Size(0, 20);
			this.lblContTablaAmortizacion.TabIndex = 775;
			this.lblContTablaAmortizacion.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance58.FontData.Name = "Tahoma";
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance58;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.TbBasico);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Controls.Add(this.ultraTabPageControl1);
			this.utcReportes.Location = new System.Drawing.Point(8, 48);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(976, 360);
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.Flat;
			this.utcReportes.TabIndex = 681;
			appearance59.BackColor = System.Drawing.SystemColors.Control;
			ultraTab1.Appearance = appearance59;
			ultraTab1.TabPage = this.TbBasico;
			ultraTab1.Text = "Facturas";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Pagare ";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Tabla De Amortización";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3});
			this.utcReportes.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcReportes_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(972, 334);
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
			// txtAño
			// 
			this.txtAño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAño.Location = new System.Drawing.Point(64, 14);
			this.txtAño.MinValue = 2015;
			this.txtAño.Name = "txtAño";
			this.txtAño.PromptChar = ' ';
			this.txtAño.Size = new System.Drawing.Size(88, 21);
			this.txtAño.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAño.TabIndex = 740;
			this.txtAño.Click += new System.EventHandler(this.txtAño_Click);
			// 
			// lblPeriodo
			// 
			this.lblPeriodo.AutoSize = true;
			this.lblPeriodo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPeriodo.Location = new System.Drawing.Point(8, 16);
			this.lblPeriodo.Name = "lblPeriodo";
			this.lblPeriodo.Size = new System.Drawing.Size(45, 16);
			this.lblPeriodo.TabIndex = 859;
			this.lblPeriodo.Text = "Periodo";
			this.lblPeriodo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmReporteDataPuraCobranza
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 414);
			this.Controls.Add(this.lblPeriodo);
			this.Controls.Add(this.txtAño);
			this.Controls.Add(this.utcReportes);
			this.Name = "frmReporteDataPuraCobranza";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Data Pura Credito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteDataPuraCobranza_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteDataPuraCobranza_Load);
			this.TbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridPagare)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridTablaAmortizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAño)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

   int iYearA = 0;
		DateTime dtInicio ;
		DateTime dtFinalPeriodo ;
		private void btnConsultarPagare_Click(object sender, System.EventArgs e)
		{

			if (!Validacion.vbCampoEnteroVacio(this.txtAño, "Ingrese el periodo", 2015, DateTime.Today.Year + 1)) return;

			 iYearA = (int)this.txtAño.Value;

			 dtInicio = DateTime.Parse("01/01" + "/" + iYearA);
			 dtFinalPeriodo = DateTime.Parse("31/12" + "/" + iYearA);

			this.lblcontPagare.Text = "";

			string sSQLPa = string.Format("Exec ReporteDataPuraPagare '{0}', '{1}'", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")  );
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridPagare);

			this.lblcontPagare.Text = this.uGridPagare.Rows.Count + " Registros Encontrados";

		}

		private void btnExcelPagare_Click(object sender, System.EventArgs e)
		{
				if (!Validacion.vbCampoEnteroVacio(this.txtAño, "Ingrese el periodo", 2015, DateTime.Today.Year + 1)) return;

			iYearA = (int)this.txtAño.Value;
			dtInicio = DateTime.Parse("01/01" + "/" + iYearA);
			dtFinalPeriodo = DateTime.Parse("31/12" + "/" + iYearA);

			#region Pagare 

			SaveFileDialog sArchivoTelefono = new SaveFileDialog();

			sArchivoTelefono.Filter = "Documentos de texto (*.txt)|*.txt";
			sArchivoTelefono.FileName = "Date Pura Pagare.txt";
			sArchivoTelefono.OverwritePrompt = true;
			sArchivoTelefono.Title = "Genera Archivo Date Pura Pagare";
				                                
			if (sArchivoTelefono.ShowDialog() == DialogResult.OK)
			{					
				#region Pagos
				using (StreamWriter file = new StreamWriter(sArchivoTelefono.FileName, true))
				{
					#region Normales
					int i = 0;
					foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec ReporteDataPuraPagare '{0}', '{1}'",
						Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"))).Rows)
					{
						try
						{
								
							string iLinea = 
								dr["Cedula"].ToString() 
								+ ";" + dr["Cliente"].ToString()
								+ ";" + dr["Almacen"].ToString() 
								+ ";" + dr["Numero"].ToString() 
								+ ";" + dr["Cuota"].ToString() 
								+ ";" + dr["Fecha"].ToString() 
								+ ";" + dr["Pago"].ToString() 
								+ ";" + dr["Interes"].ToString() 
								+ ";" + dr["Capital"].ToString() 
								+ ";" + dr["Saldo"].ToString()  
								;
							
							file.WriteLine(iLinea);

							i++;

//							sArchivos = sArchivos +"Referencias/";
						}						
						catch (Exception exc)
						{
							MessageBox.Show(exc.Message + " Linea : " + i.ToString());
							continue;
						}
					}

										if (i > 0) MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					#endregion Normales

				}
				#endregion Pagos
					
			}

			#endregion Pagare 
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnConsultaTablaAmortizacion_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAño, "Ingrese el periodo", 2015, DateTime.Today.Year + 1)) return;

			this.lblContTablaAmortizacion.Text = "";

			string sSQLPa = string.Format("Exec ReporteDataPuraTablaAmortizacion ");
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridTablaAmortizacion);

			this.lblContTablaAmortizacion.Text = this.uGridTablaAmortizacion.Rows.Count + " Registros Encontrados";
		
		}

		private void frmReporteDataPuraCobranza_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200603");

			if (!Funcion.Permiso("1246", cdsSeteoF))
			{				
				MessageBox.Show("No tiene el acceso al Reporte Data Pura Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar)
			{
				this.btnExcelFacturas.Enabled = true;
				this.btnExcelPagare.Enabled = true ;
				this.btnExcelTablaAmortizacion.Enabled = true;
			}

			this.txtAño.MaxValue = DateTime.Today.Year;
			this.txtAño.Value = DateTime.Today.Year;
		}

		private void btnconsultaFacturas_Click(object sender, System.EventArgs e)
		{
			

			if (!Validacion.vbCampoEnteroVacio(this.txtAño, "Ingrese el periodo", 2015, DateTime.Today.Year + 1)) return;

			iYearA = (int)this.txtAño.Value;

			dtInicio = DateTime.Parse("01/01" + "/" + iYearA);
			dtFinalPeriodo = DateTime.Parse("31/12" + "/" + iYearA);

			this.lblcontPagare.Text = "";

			string sSQLPa = string.Format("Exec ReporteDataPuraVentas '{0}', '{1}'", Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd")  );
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.uGridDocumentos);

			this.lblcontPagare.Text = this.uGridPagare.Rows.Count + " Registros Encontrados";
		}

		private void btnExcelFacturas_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAño, "Ingrese el periodo", 2015, DateTime.Today.Year + 1)) return;
			iYearA = (int)this.txtAño.Value;

			dtInicio = DateTime.Parse("01/01" + "/" + iYearA);
			dtFinalPeriodo = DateTime.Parse("31/12" + "/" + iYearA);

			#region Tabla Ventas 

			SaveFileDialog sArchivoTelefono = new SaveFileDialog();

			sArchivoTelefono.Filter = "Documentos de texto (*.txt)|*.txt";
			sArchivoTelefono.FileName = "Data Ventas .txt";
			sArchivoTelefono.OverwritePrompt = true;
			sArchivoTelefono.Title = "Genera Archivo Ventas";
				                                
			if (sArchivoTelefono.ShowDialog() == DialogResult.OK)
			{					
				#region Pagos
				using (StreamWriter file = new StreamWriter(sArchivoTelefono.FileName, true))
				{
					#region Normales
					int i = 0;
					foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec ReporteDataPuraVentas '{0}', '{1}'",
						Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"))).Rows)
					{
						try
						{
								
							string iLinea = 
								dr["Tipo"].ToString() 
								+ ";" + dr["Fecha"].ToString() 
								+ ";" + dr["Almacen"].ToString() 
								+ ";"  + dr["Banco"].ToString() 
								+ ";"  + dr["Numero"].ToString() 
								+ ";" + dr["Cedula"].ToString() 
								+ ";" + dr["Cliente"].ToString() 
								+ ";"  + dr["Total"].ToString() 
								+ ";"  + dr["Vendedor"].ToString()  ;
							
							file.WriteLine(iLinea);

							i++;

							//							sArchivos = sArchivos +"Referencias/";
						}						
						catch (Exception exc)
						{
							MessageBox.Show(exc.Message + " Linea : " + i.ToString());
							continue;
						}
					}

					if (i > 0) MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					#endregion Normales
				}
				#endregion Pagos
					
			}

			#endregion Ventas

		}

		private void btnExcelTablaAmortizacion_Click(object sender, System.EventArgs e)
		{
				if (!Validacion.vbCampoEnteroVacio(this.txtAño, "Ingrese el periodo", 2015, DateTime.Today.Year + 1)) return;

			iYearA = (int)this.txtAño.Value;

			dtInicio = DateTime.Parse("01/01" + "/" + iYearA);
			dtFinalPeriodo = DateTime.Parse("31/12" + "/" + iYearA);

			#region Tabla Amrotizacion 

			SaveFileDialog sArchivoTelefono = new SaveFileDialog();

			sArchivoTelefono.Filter = "Documentos de texto (*.txt)|*.txt";
			sArchivoTelefono.FileName = "Data Tabla Amrotizacion .txt";
			sArchivoTelefono.OverwritePrompt = true;
			sArchivoTelefono.Title = "Genera Archivo Tabla Amrotizacion ";
				                                
			if (sArchivoTelefono.ShowDialog() == DialogResult.OK)
			{					
				#region Pagos
				using (StreamWriter file = new StreamWriter(sArchivoTelefono.FileName, true))
				{
					#region Normales
					int i = 0;
					foreach (DataRow dr in FuncionesProcedimientos.dtGeneral(string.Format("Exec ReporteDataPuraTablaAmortizacion '{0}', '{1}'",
						Convert.ToDateTime(dtInicio).ToString("yyyyMMdd"), Convert.ToDateTime(dtFinalPeriodo).ToString("yyyyMMdd"))).Rows)
					{
						try
						{
								
							string iLinea = 
								dr["Documento"].ToString() + ";" 
								+ dr["Cedula"].ToString() + ";" 
								+ dr["Cliente"].ToString() + ";" 
								+ dr["Almacen"].ToString() 
								+ ";"  + dr["Banco"].ToString() 
								+ ";"  + dr["FechaVenta"].ToString() 
								+ ";"  + dr["Numero"].ToString() 
								+ ";" + dr["Cuota"].ToString() 
								+ ";" + dr["Vence"].ToString() 
								+ ";" + dr["Valor"].ToString()  
								+ ";" + dr["Total_Interes"].ToString() 
								+ ";" + dr["Total_Capital"].ToString() 
								+ ";" +  dr["Interes"].ToString() 
								+ ";" + dr["Capital"].ToString() 
								+ ";" +  dr["Mora"].ToString() 
								+ ";" + dr["Saldo"].ToString() 
								+ ";" +  dr["Abono"].ToString() 
								+ ";" + dr["Retraso"].ToString()  ;
							
							file.WriteLine(iLinea);

							i++;

//							sArchivos = sArchivos +"Referencias/";
						}						
						catch (Exception exc)
						{
							MessageBox.Show(exc.Message + " Linea : " + i.ToString());
							continue;
						}
					}

					if (i > 0) MessageBox.Show(string.Format("Archivo generado correctamente con {0} lineas", i), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					#endregion Normales

				}
				#endregion Pagos
					
			}

			#endregion Tabla Amrotizacion
		}

		private void txtAño_Click(object sender, System.EventArgs e)
		{
			this.txtAño.SelectAll();
		}

		private void frmReporteDataPuraCobranza_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void utcReportes_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

	}
}
