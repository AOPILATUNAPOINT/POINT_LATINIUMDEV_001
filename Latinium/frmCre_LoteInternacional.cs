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
	/// Descripción breve de frmCre_LoteInternacional.
	/// </summary>
	public class frmCre_LoteInternacional : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtLote;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCapturados;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optValidos;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOperaciones;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtUltimoLote;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.ContextMenu mnuAcreditacion;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TbBasico;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnExportarDatos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDatos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Button btnExportarAnexos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGriAnexos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridFormato;
		private System.Windows.Forms.Button btnExportarFormato;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_LoteInternacional()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_LoteInternacional));
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapturaLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("C_Pendientes");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 7, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 8, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapturaLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("C_Pendientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("No.Operacion");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("T.DOC.", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificación");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo Operación");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Concesion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Vencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tasa Interes");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Inicio Actual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("No. Div. Restantes");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dia Fijo Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha 1ra Cuota");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Destino Comercial");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo Rel Cónyugue");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Id. Cónyugue");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre Conyugue");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo Rel Garante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Id. Garante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre Garante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total Pend. Cuota");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor Cuota");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tasa Inversion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencial");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("No.Dias Trans.");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes Dev.");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso Total");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Ref.");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor Presente");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TIPO PRODUCTO");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto Financiado");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Aprobacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Año Vehiculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorSolca");
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("No. Conc.");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("No. Oper.");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital Reducido");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital a Valor Presente");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes Devengado");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencial");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorSolca");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nro.Operación");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("T.DOC.");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor del Prestamo");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Efectiva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes Devengado");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Plazo");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Proximo Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tasa Interes");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dia Fijo Pago");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tasa Descuento");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo Rel.");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Identificacion Cod");
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Producto");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor Cuota");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorSolca");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nro.Operación");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("T.DOC.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor del Prestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Efectiva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes Devengado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Plazo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Proximo Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tasa Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dia Fijo Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tasa Descuento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo Rel.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificacion Cod");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Producto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorSolca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No. Conc.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No. Oper.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital Reducido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Capital a Valor Presente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes Devengado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorSolca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No.Operacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("T.DOC.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Identificación");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo Operación");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Concesion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Vencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tasa Interes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Inicio Actual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No. Div. Restantes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dia Fijo Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha 1ra Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Destino Comercial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo Rel Cónyugue");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Id. Cónyugue");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre Conyugue");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo Rel Garante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Id. Garante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre Garante");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total Pend. Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor Cuota");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tasa Inversion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("No.Dias Trans.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Interes Dev.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Ref.");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor Presente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TIPO PRODUCTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto Financiado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Aprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Año Vehiculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorSolca");
			this.lblContador = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optCapturados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optValidos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExportar = new System.Windows.Forms.Button();
			this.uGridOperaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtUltimoLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.mnuAcreditacion = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.TbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label4 = new System.Windows.Forms.Label();
			this.btnExportarDatos = new System.Windows.Forms.Button();
			this.uGridDatos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label5 = new System.Windows.Forms.Label();
			this.btnExportarAnexos = new System.Windows.Forms.Button();
			this.uGriAnexos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label8 = new System.Windows.Forms.Label();
			this.uGridFormato = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExportarFormato = new System.Windows.Forms.Button();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optCapturados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optValidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUltimoLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			this.TbBasico.SuspendLayout();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGriAnexos)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridFormato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(56, 16);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 796;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(148, 44);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 16);
			this.label1.TabIndex = 795;
			this.label1.Text = "Ultimo Lote";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(4, 44);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(26, 16);
			this.label3.TabIndex = 793;
			this.label3.Text = "Lote";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(4, 75);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 788;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(4, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 784;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtLote
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLote.Appearance = appearance1;
			this.txtLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtLote.FormatString = "";
			this.txtLote.Location = new System.Drawing.Point(60, 41);
			this.txtLote.Name = "txtLote";
			this.txtLote.PromptChar = ' ';
			this.txtLote.Size = new System.Drawing.Size(72, 21);
			this.txtLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtLote.TabIndex = 792;
			// 
			// optCapturados
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.optCapturados.Appearance = appearance2;
			this.optCapturados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optCapturados.CheckedIndex = 0;
			this.optCapturados.ItemAppearance = appearance3;
			this.optCapturados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Capturados";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "No Capturados";
			this.optCapturados.Items.Add(valueListItem1);
			this.optCapturados.Items.Add(valueListItem2);
			this.optCapturados.Items.Add(valueListItem3);
			this.optCapturados.ItemSpacingVertical = 10;
			this.optCapturados.Location = new System.Drawing.Point(412, 6);
			this.optCapturados.Name = "optCapturados";
			this.optCapturados.Size = new System.Drawing.Size(256, 26);
			this.optCapturados.TabIndex = 791;
			this.optCapturados.Text = "Todos";
			this.optCapturados.ValueChanged += new System.EventHandler(this.optCapturados_ValueChanged);
			// 
			// optValidos
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.optValidos.Appearance = appearance4;
			this.optValidos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optValidos.CheckedIndex = 1;
			this.optValidos.ItemAppearance = appearance5;
			this.optValidos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 0;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Validos";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "No Validos";
			this.optValidos.Items.Add(valueListItem4);
			this.optValidos.Items.Add(valueListItem5);
			this.optValidos.Items.Add(valueListItem6);
			this.optValidos.ItemSpacingVertical = 10;
			this.optValidos.Location = new System.Drawing.Point(412, 39);
			this.optValidos.Name = "optValidos";
			this.optValidos.Size = new System.Drawing.Size(256, 26);
			this.optValidos.TabIndex = 790;
			this.optValidos.Text = "Validos";
			this.optValidos.ValueChanged += new System.EventHandler(this.optValidos_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-4, 105);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1192, 8);
			this.groupBox1.TabIndex = 789;
			this.groupBox1.TabStop = false;
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(60, 73);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(232, 19);
			this.txtBuscar.TabIndex = 787;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(308, 72);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 786;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtpFecha
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance6;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(60, 9);
			this.dtpFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 785;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1064, 16);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(96, 23);
			this.btnExportar.TabIndex = 783;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// uGridOperaciones
			// 
			this.uGridOperaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOperaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOperaciones.DataSource = this.ultraDataSource2;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridOperaciones.DisplayLayout.Appearance = appearance7;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "...";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance8;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 60;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance9;
			ultraGridColumn4.Header.Caption = "Factura";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.Caption = "Emision";
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance10;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 250;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance11;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance12;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance13;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance14;
			ultraGridColumn11.Header.Caption = "Cuotas Pend";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 90;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance15;
			ultraGridColumn12.Header.Caption = "Dias Mora";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 300;
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
																										 ultraGridColumn13});
			ultraGridBand1.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance16;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance17;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "Totales";
			this.uGridOperaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridOperaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridOperaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOperaciones.DisplayLayout.Override.CellAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8.5F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOperaciones.DisplayLayout.Override.HeaderAppearance = appearance20;
			this.uGridOperaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOperaciones.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridOperaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOperaciones.Location = new System.Drawing.Point(8, 56);
			this.uGridOperaciones.Name = "uGridOperaciones";
			this.uGridOperaciones.Size = new System.Drawing.Size(1168, 384);
			this.uGridOperaciones.TabIndex = 782;
			this.uGridOperaciones.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridOperaciones_CellChange);
			this.uGridOperaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOperaciones_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(bool);
			ultraDataColumn2.DefaultValue = false;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn13});
			// 
			// txtUltimoLote
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtUltimoLote.Appearance = appearance23;
			this.txtUltimoLote.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtUltimoLote.Enabled = false;
			this.txtUltimoLote.FormatString = "";
			this.txtUltimoLote.Location = new System.Drawing.Point(220, 41);
			this.txtUltimoLote.Name = "txtUltimoLote";
			this.txtUltimoLote.PromptChar = ' ';
			this.txtUltimoLote.Size = new System.Drawing.Size(72, 21);
			this.txtUltimoLote.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtUltimoLote.TabIndex = 794;
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
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(1088, 32);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(88, 23);
			this.btnGenerar.TabIndex = 797;
			this.btnGenerar.Text = "Generar TXT";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// mnuAcreditacion
			// 
			this.mnuAcreditacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																										this.menuItem1,
																																										this.menuItem2,
																																										this.menuItem3,
																																										this.menuItem4});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Trama De Clientes";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 1;
			this.menuItem2.Text = "Trama De Prestamoss";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 2;
			this.menuItem3.Text = "Tabla de Amortizacióm";
			this.menuItem3.Click += new System.EventHandler(this.menuItem3_Click);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 3;
			this.menuItem4.Text = "Valor de la Prima";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance24.FontData.Name = "Tahoma";
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance24;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.TbBasico);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Controls.Add(this.ultraTabPageControl1);
			this.utcReportes.Controls.Add(this.ultraTabPageControl2);
			this.utcReportes.Location = new System.Drawing.Point(8, 120);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(1176, 480);
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.Flat;
			this.utcReportes.TabIndex = 1059;
			ultraTab1.TabPage = this.TbBasico;
			ultraTab1.Text = "Seleccion";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Datos";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Anexos";
			ultraTab4.TabPage = this.ultraTabPageControl2;
			ultraTab4.Text = "Formato";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2,
																																													ultraTab3,
																																													ultraTab4});
			this.utcReportes.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcReportes_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1172, 454);
			// 
			// TbBasico
			// 
			this.TbBasico.Controls.Add(this.uGridOperaciones);
			this.TbBasico.Controls.Add(this.btnExportar);
			this.TbBasico.Controls.Add(this.lblContador);
			this.TbBasico.Location = new System.Drawing.Point(1, 23);
			this.TbBasico.Name = "TbBasico";
			this.TbBasico.Size = new System.Drawing.Size(1172, 454);
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.label4);
			this.ultraTabPageControl10.Controls.Add(this.btnExportarDatos);
			this.ultraTabPageControl10.Controls.Add(this.uGridDatos);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(1172, 454);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.SystemColors.Control;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.label4.ForeColor = System.Drawing.Color.Firebrick;
			this.label4.Location = new System.Drawing.Point(56, 16);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(0, 17);
			this.label4.TabIndex = 1057;
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExportarDatos
			// 
			this.btnExportarDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportarDatos.CausesValidation = false;
			this.btnExportarDatos.Image = ((System.Drawing.Image)(resources.GetObject("btnExportarDatos.Image")));
			this.btnExportarDatos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarDatos.Location = new System.Drawing.Point(1064, 16);
			this.btnExportarDatos.Name = "btnExportarDatos";
			this.btnExportarDatos.Size = new System.Drawing.Size(96, 23);
			this.btnExportarDatos.TabIndex = 1045;
			this.btnExportarDatos.Text = "Exportar";
			this.btnExportarDatos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportarDatos.Click += new System.EventHandler(this.btnExportarDatos_Click);
			// 
			// uGridDatos
			// 
			this.uGridDatos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDatos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDatos.DataSource = this.ultraDataSource3;
			appearance25.BackColor = System.Drawing.Color.White;
			this.uGridDatos.DisplayLayout.Appearance = appearance25;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance26;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance27;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance28;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance29;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 4;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 5;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 6;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance30;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 7;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 8;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance31;
			ultraGridColumn23.Header.VisiblePosition = 9;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 10;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 11;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn26.CellAppearance = appearance32;
			ultraGridColumn26.Header.VisiblePosition = 12;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 13;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 14;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 15;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 16;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 17;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 18;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance33;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 19;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance34;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellButtonAppearance = appearance35;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.VisiblePosition = 20;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.VisiblePosition = 21;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance36;
			ultraGridColumn36.Header.VisiblePosition = 22;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellAppearance = appearance37;
			ultraGridColumn37.Header.VisiblePosition = 23;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.VisiblePosition = 24;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance38;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellButtonAppearance = appearance39;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 25;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 26;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance40;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellButtonAppearance = appearance41;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 27;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn42.CellAppearance = appearance42;
			ultraGridColumn42.Header.VisiblePosition = 28;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance43;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellButtonAppearance = appearance44;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 29;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 30;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 31;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance45;
			appearance46.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellButtonAppearance = appearance46;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 32;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn46});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridDatos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridDatos.DisplayLayout.GroupByBox.Hidden = true;
			appearance47.ForeColor = System.Drawing.Color.Black;
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDatos.DisplayLayout.Override.ActiveRowAppearance = appearance47;
			this.uGridDatos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDatos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance48.BackColor = System.Drawing.Color.Transparent;
			this.uGridDatos.DisplayLayout.Override.CardAreaAppearance = appearance48;
			appearance49.ForeColor = System.Drawing.Color.Black;
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDatos.DisplayLayout.Override.CellAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance50.FontData.BoldAsString = "True";
			appearance50.FontData.Name = "Arial";
			appearance50.FontData.SizeInPoints = 8F;
			appearance50.ForeColor = System.Drawing.Color.White;
			appearance50.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDatos.DisplayLayout.Override.HeaderAppearance = appearance50;
			this.uGridDatos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance51.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance51.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.RowAlternateAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.RowSelectorAppearance = appearance52;
			appearance53.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance53.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance53.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDatos.DisplayLayout.Override.SelectedRowAppearance = appearance53;
			this.uGridDatos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridDatos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDatos.Location = new System.Drawing.Point(8, 56);
			this.uGridDatos.Name = "uGridDatos";
			this.uGridDatos.Size = new System.Drawing.Size(1160, 392);
			this.uGridDatos.TabIndex = 765;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Controls.Add(this.btnExportarAnexos);
			this.ultraTabPageControl1.Controls.Add(this.uGriAnexos);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1172, 454);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.SystemColors.Control;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(56, 16);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(0, 17);
			this.label5.TabIndex = 1057;
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExportarAnexos
			// 
			this.btnExportarAnexos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportarAnexos.Image = ((System.Drawing.Image)(resources.GetObject("btnExportarAnexos.Image")));
			this.btnExportarAnexos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarAnexos.Location = new System.Drawing.Point(1064, 16);
			this.btnExportarAnexos.Name = "btnExportarAnexos";
			this.btnExportarAnexos.Size = new System.Drawing.Size(96, 23);
			this.btnExportarAnexos.TabIndex = 1004;
			this.btnExportarAnexos.Text = "Exportar";
			this.btnExportarAnexos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportarAnexos.Click += new System.EventHandler(this.btnExportarAnexos_Click);
			// 
			// uGriAnexos
			// 
			this.uGriAnexos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGriAnexos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGriAnexos.DataSource = this.ultraDataSource4;
			appearance54.BackColor = System.Drawing.Color.White;
			this.uGriAnexos.DisplayLayout.Appearance = appearance54;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn47.CellAppearance = appearance55;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Width = 100;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn48.CellAppearance = appearance56;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 155;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn49.CellAppearance = appearance57;
			ultraGridColumn49.Header.VisiblePosition = 2;
			ultraGridColumn49.Width = 200;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn50.Header.VisiblePosition = 3;
			ultraGridColumn50.Width = 150;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn51.CellAppearance = appearance58;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Width = 103;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn52.CellAppearance = appearance59;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.VisiblePosition = 5;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn53.CellAppearance = appearance60;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.VisiblePosition = 6;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn54.CellAppearance = appearance61;
			ultraGridColumn54.Header.VisiblePosition = 7;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn55.CellAppearance = appearance62;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.VisiblePosition = 8;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance63;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 9;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGriAnexos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGriAnexos.DisplayLayout.GroupByBox.Hidden = true;
			appearance64.ForeColor = System.Drawing.Color.Black;
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGriAnexos.DisplayLayout.Override.ActiveRowAppearance = appearance64;
			this.uGriAnexos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGriAnexos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance65.BackColor = System.Drawing.Color.Transparent;
			this.uGriAnexos.DisplayLayout.Override.CardAreaAppearance = appearance65;
			appearance66.ForeColor = System.Drawing.Color.Black;
			appearance66.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGriAnexos.DisplayLayout.Override.CellAppearance = appearance66;
			appearance67.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance67.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance67.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance67.FontData.BoldAsString = "True";
			appearance67.FontData.Name = "Arial";
			appearance67.FontData.SizeInPoints = 8F;
			appearance67.ForeColor = System.Drawing.Color.White;
			appearance67.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGriAnexos.DisplayLayout.Override.HeaderAppearance = appearance67;
			this.uGriAnexos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance68.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance68.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance68.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGriAnexos.DisplayLayout.Override.RowAlternateAppearance = appearance68;
			appearance69.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance69.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance69.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGriAnexos.DisplayLayout.Override.RowSelectorAppearance = appearance69;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGriAnexos.DisplayLayout.Override.SelectedRowAppearance = appearance70;
			this.uGriAnexos.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGriAnexos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGriAnexos.Location = new System.Drawing.Point(8, 56);
			this.uGriAnexos.Name = "uGriAnexos";
			this.uGriAnexos.Size = new System.Drawing.Size(1160, 392);
			this.uGriAnexos.TabIndex = 774;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.label8);
			this.ultraTabPageControl2.Controls.Add(this.uGridFormato);
			this.ultraTabPageControl2.Controls.Add(this.btnExportarFormato);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1172, 454);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.SystemColors.Control;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(56, 16);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(0, 17);
			this.label8.TabIndex = 1057;
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// uGridFormato
			// 
			this.uGridFormato.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFormato.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridFormato.DataSource = this.ultraDataSource5;
			appearance71.BackColor = System.Drawing.Color.White;
			this.uGridFormato.DisplayLayout.Appearance = appearance71;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn57.CellAppearance = appearance72;
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn58.CellAppearance = appearance73;
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn59.CellAppearance = appearance74;
			ultraGridColumn59.Header.VisiblePosition = 2;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn60.Header.VisiblePosition = 3;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance75;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.VisiblePosition = 4;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn62.Header.VisiblePosition = 5;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance76;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.VisiblePosition = 6;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn64.CellAppearance = appearance77;
			ultraGridColumn64.Header.VisiblePosition = 7;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn65.Header.VisiblePosition = 8;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance78;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 9;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn67.CellAppearance = appearance79;
			ultraGridColumn67.Header.VisiblePosition = 10;
			ultraGridColumn68.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance80;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.VisiblePosition = 11;
			ultraGridColumn69.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn69.CellAppearance = appearance81;
			ultraGridColumn69.Header.VisiblePosition = 12;
			ultraGridColumn70.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn70.CellAppearance = appearance82;
			ultraGridColumn70.Header.VisiblePosition = 13;
			ultraGridColumn71.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn71.Header.VisiblePosition = 14;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn72.CellAppearance = appearance83;
			ultraGridColumn72.Header.VisiblePosition = 15;
			ultraGridColumn73.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance84;
			ultraGridColumn73.Format = "#,##0.00";
			ultraGridColumn73.Header.VisiblePosition = 16;
			ultraGridColumn74.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance85;
			ultraGridColumn74.Format = "#,##0.00";
			ultraGridColumn74.Header.VisiblePosition = 17;
			ultraGridBand4.Columns.AddRange(new object[] {
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
																										 ultraGridColumn74});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridFormato.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridFormato.DisplayLayout.GroupByBox.Hidden = true;
			appearance86.ForeColor = System.Drawing.Color.Black;
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFormato.DisplayLayout.Override.ActiveRowAppearance = appearance86;
			this.uGridFormato.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridFormato.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance87.BackColor = System.Drawing.Color.Transparent;
			this.uGridFormato.DisplayLayout.Override.CardAreaAppearance = appearance87;
			appearance88.ForeColor = System.Drawing.Color.Black;
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFormato.DisplayLayout.Override.CellAppearance = appearance88;
			appearance89.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance89.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance89.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance89.FontData.BoldAsString = "True";
			appearance89.FontData.Name = "Arial";
			appearance89.FontData.SizeInPoints = 8F;
			appearance89.ForeColor = System.Drawing.Color.White;
			appearance89.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFormato.DisplayLayout.Override.HeaderAppearance = appearance89;
			this.uGridFormato.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance90.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance90.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance90.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFormato.DisplayLayout.Override.RowAlternateAppearance = appearance90;
			appearance91.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance91.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance91.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFormato.DisplayLayout.Override.RowSelectorAppearance = appearance91;
			appearance92.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance92.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance92.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFormato.DisplayLayout.Override.SelectedRowAppearance = appearance92;
			this.uGridFormato.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridFormato.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFormato.Location = new System.Drawing.Point(8, 56);
			this.uGridFormato.Name = "uGridFormato";
			this.uGridFormato.Size = new System.Drawing.Size(1160, 392);
			this.uGridFormato.TabIndex = 1046;
			// 
			// btnExportarFormato
			// 
			this.btnExportarFormato.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportarFormato.CausesValidation = false;
			this.btnExportarFormato.Image = ((System.Drawing.Image)(resources.GetObject("btnExportarFormato.Image")));
			this.btnExportarFormato.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportarFormato.Location = new System.Drawing.Point(1064, 16);
			this.btnExportarFormato.Name = "btnExportarFormato";
			this.btnExportarFormato.Size = new System.Drawing.Size(96, 23);
			this.btnExportarFormato.TabIndex = 1045;
			this.btnExportarFormato.Text = "Exportar";
			this.btnExportarFormato.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportarFormato.Click += new System.EventHandler(this.btnExportarFormato_Click);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.DateTime);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn31});
			// 
			// ultraDataSource4
			// 
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn46.DataType = typeof(System.Decimal);
			ultraDataColumn47.DataType = typeof(System.DateTime);
			ultraDataColumn48.DataType = typeof(System.DateTime);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn50.DataType = typeof(System.DateTime);
			ultraDataColumn53.DataType = typeof(System.DateTime);
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn59.DataType = typeof(int);
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(System.Decimal);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			ultraDataColumn68.DataType = typeof(System.DateTime);
			ultraDataColumn69.DataType = typeof(System.Decimal);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			ultraDataColumn72.DataType = typeof(System.DateTime);
			ultraDataColumn74.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn74});
			// 
			// frmCre_LoteInternacional
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1184, 598);
			this.Controls.Add(this.utcReportes);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtLote);
			this.Controls.Add(this.optCapturados);
			this.Controls.Add(this.optValidos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.txtUltimoLote);
			this.Name = "frmCre_LoteInternacional";
			this.Text = "Generación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCre_LoteInternacional_KeyDown);
			this.Load += new System.EventHandler(this.frmCre_LoteInternacional_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optCapturados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optValidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOperaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtUltimoLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			this.TbBasico.ResumeLayout(false);
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridDatos)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGriAnexos)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridFormato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmCre_LoteInternacional_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320901");
			
			if (!Funcion.Permiso("2136", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a la Generacón del Lote", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportarExcel) this.btnExportar.Enabled = true;
			this.dtpFecha.CalendarInfo.MinDate = DateTime.Today;
			this.dtpFecha.CalendarInfo.MaxDate = DateTime.Today.AddDays(21);
			this.dtpFecha.Value = DateTime.Today;

			this.txtUltimoLote.Value = Funcion.iEscalarSQL(cdsSeteoF, "Exec Cre_RetornaUltimoLote 19");		

			this.txtLote.Value = (int)this.txtUltimoLote.Value + 1;

			this.btnExportar.Enabled = true;
			this.ConsultaGrids();
		}
		private void Consulta ()
		{
			string sSQL = string.Format("Exec Cre_LoteInternacional '{0}', {1}, '{2}', {3}, {4}", 
				Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), (int)this.txtLote.Value, this.txtBuscar.Text.ToString(), 
				(int)this.optCapturados.Value, (int)this.optValidos.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridOperaciones);
			
			this.lblContador.Text = this.uGridOperaciones.Rows.Count + " Registros Encontrados";
		}
		private void ConsultaGrids()
		{
			/*Datos*/
			string sSQL = string.Format("Exec Cre_TramasXLSInternacional '{0}',{1}, {2}", 
				Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), (int)this.txtUltimoLote.Value, 1);			
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDatos);
			this.label4.Text = this.uGridDatos.Rows.Count.ToString() + " Registros Encontrados";

			/*Anexos*/
			string sSQL1 = string.Format("Exec Cre_TramasXLSInternacional '{0}',{1},{2}", 
				Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), (int)this.txtUltimoLote.Value, 2);	
			FuncionesProcedimientos.dsGeneral(sSQL1, this.uGriAnexos);
			this.label5.Text = this.uGriAnexos.Rows.Count.ToString() + " Registros Encontrados";

			/*formato*/
			string sSQL2 = string.Format("Exec Cre_TramasXLSInternacional '{0}',{1},{2}", 
				Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), (int)this.txtUltimoLote.Value, 3);			
			FuncionesProcedimientos.dsGeneral(sSQL2, this.uGridFormato);
			this.label8.Text = this.uGridFormato.Rows.Count.ToString() + " Registros Encontrados";
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridOperaciones);
		}

		private void uGridOperaciones_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "CapturaLote")
			{
				this.uGridOperaciones.UpdateData();

				int iCuotasVendidas = (int)e.Cell.Row.Cells["Cuotas"].Value - (int)e.Cell.Row.Cells["C_Pendientes"].Value;
				
				string sSQL = string.Format("Exec Cre_LoteBancoPichincha {0}, {1}, {2}, {3}, {4}, 19, ''", 
					(bool)e.Cell.Row.Cells["CapturaLote"].Value, (int)e.Cell.Row.Cells["idCompra"].Value, (decimal)e.Cell.Row.Cells["Saldo"].Value, 
					(int)e.Cell.Row.Cells["C_Pendientes"].Value, (int)this.txtLote.Value);				
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			}	
		}

		private void frmCre_LoteInternacional_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void optCapturados_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void optValidos_ValueChanged(object sender, System.EventArgs e)
		{
			this.Consulta();
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			GuardarArchivo("POIN_clientes", "Tramas de Clientes", "Clientes", 1);
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			GuardarArchivo("POIN_prestamos", "Tramas de Préstamos", "Prestamos", 2);
		}

		private void menuItem3_Click(object sender, System.EventArgs e)
		{
			GuardarArchivo("POIN_tablas", "Tablas de Amortización", "Tablas", 3);
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{		
			GuardarArchivo("POIN_valcompra", "Valores de Prima", "Primas", 4);
		}
		private void GuardarArchivo(string nombreArchivoBase, string tituloDialogo, string nombreColumna, int tipoTrama)
		{
			#region Validación común
			if (this.dtpFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpFecha.Focus();
				return;
			}
    
			if (!Validacion.vbFechaEnDocumentos(this.dtpFecha, "Ingrese la Fecha ", true, "Cobros", cdsSeteoF)) return;

			if (!Validacion.vbCampoEnteroVacio(this.txtLote, "Ingrese el Numero de Lote", 1, 1000000))
			{
				this.txtLote.Focus();
				return;
			}

			int iContLote = Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC Cre_ValidaLotes 4,{0}", (int)this.txtLote.Value));
    
			if (iContLote < 1)
			{
				MessageBox.Show(string.Format("No Existen Créditos Seleccionados para generar el Lote N. {0}", this.txtLote.Text), 
					"Validación", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtLote.Focus();
				return;
			}
			#endregion

			try
			{
				// Ruta base: Mis Documentos + carpeta con fecha
				string documentosPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
				string fechaCarpeta = Convert.ToDateTime(this.dtpFecha.Value).ToString("ddMMyyyy");
				string rutaFinal = Path.Combine(documentosPath, fechaCarpeta);
        
				Directory.CreateDirectory(rutaFinal); // Asegura que la carpeta exista

				// Configuración del diálogo
				SaveFileDialog dialogo = new SaveFileDialog();
				dialogo.Filter = "Documentos de texto (*.txt)|*.txt";
				dialogo.FileName = nombreArchivoBase + fechaCarpeta + ".txt"; // Ej: POIN_clientes01012024.txt
				dialogo.InitialDirectory = rutaFinal; // Abre directamente en la carpeta de la fecha
				dialogo.OverwritePrompt = true;
				dialogo.Title = tituloDialogo;

				if (dialogo.ShowDialog() == DialogResult.OK)
				{
					// Genera el archivo dentro de la carpeta de fecha
					using (StreamWriter archivo = new StreamWriter(dialogo.FileName, false))
					{
						int contador = 0;
						string fechaFormateada = Convert.ToDateTime(dtpFecha.Value).ToString("yyyyMMdd");
						string consulta = string.Format("EXEC Cre_TramasInternacionale '{0}',{1},{2}",
							fechaFormateada, (int)txtLote.Value, tipoTrama);
                
						DataTable datos = FuncionesProcedimientos.dtGeneral(consulta);

						foreach (DataRow fila in datos.Rows)
						{
							archivo.WriteLine(fila[nombreColumna].ToString());
							contador++;
						}

						MessageBox.Show(string.Format("{0} {1} generados correctamente.", contador, tituloDialogo), 
							"Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Error al generar {0}: {1} - {2}", tituloDialogo, ex.GetType(), ex.Message), 
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}



		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			this.mnuAcreditacion.Show(this.btnGenerar, new Point(1, 25));
		}

		private void uGridOperaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExportarDatos_Click(object sender, System.EventArgs e)
		{
			if (txtBuscar.Text != null && txtBuscar.Text.Trim() != "")
			{
				MessageBox.Show("Debe dejar el filtro vacío para exportar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.txtBuscar.Text = "";
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.uGridDatos);
		}

		private void btnExportarAnexos_Click(object sender, System.EventArgs e)
		{
			if (txtBuscar.Text != null && txtBuscar.Text.Trim() != "")
			{
				MessageBox.Show("Debe dejar el filtro vacío para exportar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.txtBuscar.Text = "";
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.uGriAnexos);
		}

		private void utcReportes_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void btnExportarFormato_Click(object sender, System.EventArgs e)
		{
			if (txtBuscar.Text != null && txtBuscar.Text.Trim() != "")
			{
				MessageBox.Show("Debe dejar el filtro vacío para exportar.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				this.txtBuscar.Text = "";
				return;
			}
			FuncionesProcedimientos.ExportaExcel(this.uGridFormato);
		}
	}
}
