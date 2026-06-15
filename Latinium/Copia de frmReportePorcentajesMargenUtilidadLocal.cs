using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReportePorcentajesMargenUtilidadLocal.
	/// </summary>
	public class frmReportePorcentajesMargenUtilidadLocal : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenCrediPOINT;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenTarjeta;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenEfectivo;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargen;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtCrediPoint;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcCrediPoint;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcEfectivo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button button3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button button4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenCrediPOINT2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenTarjeta2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenEfectivo2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargen2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtCrediPoint2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtEfectivo2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcUtTarjeta2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcCrediPoint2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcEfectivo2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcTarjeta2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda2;
		private System.Windows.Forms.Button btnExcel2;
		private System.Windows.Forms.Button btnVer2;
		private System.Windows.Forms.Label lblContador2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridporcentaje;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridTotales;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridPorcentajes;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridIYD;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportePorcentajesMargenUtilidadLocal()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportePorcentajesMargenUtilidadLocal));
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeVentasLocal");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaEfectivo");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeEfectivo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTarjeta");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeTarjeta");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCrediPoint");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCrediPoint");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalUtilidad");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeUtilidadLocal");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadEfectivo");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeEfectivoUtilidad");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadTarjeta");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeTarjetaUtilidad");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadCrediPoint");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCrediPointUtilidad");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenEfectivo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenTarjeta");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenCrediPoint");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 2, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalUtilidad", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalUtilidad", 10, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Margen", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Margen", 18, true);
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeVentasLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeUtilidadLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeEfectivoUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeTarjetaUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCrediPointUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeVentasLocal");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaEfectivo");
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeEfectivo");
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTarjeta");
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeTarjeta");
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCrediPoint");
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCrediPoint");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalUtilidad");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeUtilidadLocal");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadEfectivo");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeEfectivoUtilidad");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadTarjeta");
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeTarjetaUtilidad");
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadCrediPoint");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCrediPointUtilidad", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenEfectivo");
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenTarjeta");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenCrediPoint");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 2, true);
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalUtilidad", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalUtilidad", 10, true);
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Margen", 18, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Margen", 18, true);
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ventas");
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Efectivo");
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarjeta");
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CrediPoint");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalUtilidad");
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcUtilidad");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UEfectivo");
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UTarjeta");
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UCrediPoint");
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Margen");
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MEfectivo");
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MTarjeta");
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MCrediPoint");
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ventas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Efectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeVentasLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeUtilidadLocal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeEfectivoUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeTarjetaUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadCrediPoint");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCrediPointUtilidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Margen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenCrediPoint");
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CONCEPTO", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ID");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CONCEPTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ventas");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ventas");
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CONCEPTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ID");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CONCEPTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CONCEPTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ID");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CONCEPTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ID");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CONCEPTO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			this.txtMargenCrediPOINT = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.txtMargenTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label12 = new System.Windows.Forms.Label();
			this.txtMargenEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtMargen = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtPorcUtCrediPoint = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcUtEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcUtTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcCrediPoint = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblContador = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtMargenCrediPOINT2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtMargenTarjeta2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtMargenEfectivo2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtMargen2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.txtPorcUtCrediPoint2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcUtEfectivo2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcUtTarjeta2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcCrediPoint2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcEfectivo2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtPorcTarjeta2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBusqueda2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnExcel2 = new System.Windows.Forms.Button();
			this.btnVer2 = new System.Windows.Forms.Button();
			this.dtDesde2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtHasta2 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label28 = new System.Windows.Forms.Label();
			this.label29 = new System.Windows.Forms.Label();
			this.label30 = new System.Windows.Forms.Label();
			this.lblContador2 = new System.Windows.Forms.Label();
			this.uGridInformacion2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.chkTodos2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridLocales2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.uGridporcentaje = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.button3 = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.button4 = new System.Windows.Forms.Button();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugridTotales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugridPorcentajes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugridIYD = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenCrediPOINT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtCrediPoint)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcCrediPoint)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenCrediPOINT2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtCrediPoint2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtEfectivo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtTarjeta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcCrediPoint2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcEfectivo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcTarjeta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridporcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridTotales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentajes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridIYD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			this.SuspendLayout();
			// 
			// txtMargenCrediPOINT
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenCrediPOINT.Appearance = appearance1;
			this.txtMargenCrediPOINT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenCrediPOINT.Enabled = false;
			this.txtMargenCrediPOINT.FormatString = "#,##0.00";
			this.txtMargenCrediPOINT.Location = new System.Drawing.Point(1216, 56);
			this.txtMargenCrediPOINT.MinValue = -2147483647;
			this.txtMargenCrediPOINT.Name = "txtMargenCrediPOINT";
			this.txtMargenCrediPOINT.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenCrediPOINT.PromptChar = ' ';
			this.txtMargenCrediPOINT.Size = new System.Drawing.Size(80, 21);
			this.txtMargenCrediPOINT.TabIndex = 763;
			this.txtMargenCrediPOINT.Visible = false;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(1216, 40);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(107, 16);
			this.label13.TabIndex = 764;
			this.label13.Text = "Margen CrediPOINT";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Visible = false;
			// 
			// txtMargenTarjeta
			// 
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenTarjeta.Appearance = appearance2;
			this.txtMargenTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenTarjeta.Enabled = false;
			this.txtMargenTarjeta.FormatString = "#,##0.00";
			this.txtMargenTarjeta.Location = new System.Drawing.Point(1128, 56);
			this.txtMargenTarjeta.MinValue = -2147483647;
			this.txtMargenTarjeta.Name = "txtMargenTarjeta";
			this.txtMargenTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenTarjeta.PromptChar = ' ';
			this.txtMargenTarjeta.Size = new System.Drawing.Size(80, 21);
			this.txtMargenTarjeta.TabIndex = 761;
			this.txtMargenTarjeta.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(1128, 40);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(81, 16);
			this.label12.TabIndex = 762;
			this.label12.Text = "Margen Tarjeta";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Visible = false;
			// 
			// txtMargenEfectivo
			// 
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenEfectivo.Appearance = appearance3;
			this.txtMargenEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenEfectivo.Enabled = false;
			this.txtMargenEfectivo.FormatString = "#,##0.00";
			this.txtMargenEfectivo.Location = new System.Drawing.Point(1040, 56);
			this.txtMargenEfectivo.MinValue = -2147483647;
			this.txtMargenEfectivo.Name = "txtMargenEfectivo";
			this.txtMargenEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenEfectivo.PromptChar = ' ';
			this.txtMargenEfectivo.Size = new System.Drawing.Size(80, 21);
			this.txtMargenEfectivo.TabIndex = 759;
			this.txtMargenEfectivo.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(1040, 40);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(86, 16);
			this.label11.TabIndex = 760;
			this.label11.Text = "Margen Efectivo";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Visible = false;
			// 
			// txtMargen
			// 
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargen.Appearance = appearance4;
			this.txtMargen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargen.Enabled = false;
			this.txtMargen.FormatString = "#,##0.00";
			this.txtMargen.Location = new System.Drawing.Point(952, 56);
			this.txtMargen.MinValue = -2147483647;
			this.txtMargen.Name = "txtMargen";
			this.txtMargen.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargen.PromptChar = ' ';
			this.txtMargen.Size = new System.Drawing.Size(80, 21);
			this.txtMargen.TabIndex = 757;
			this.txtMargen.Visible = false;
			// 
			// label9
			// 
			this.label9.BackColor = System.Drawing.Color.DarkGray;
			this.label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.ForeColor = System.Drawing.Color.Firebrick;
			this.label9.Location = new System.Drawing.Point(696, 8);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(240, 16);
			this.label9.TabIndex = 756;
			this.label9.Text = "UTILIDAD";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label9.Visible = false;
			// 
			// label8
			// 
			this.label8.BackColor = System.Drawing.Color.DarkGray;
			this.label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.ForeColor = System.Drawing.Color.Firebrick;
			this.label8.Location = new System.Drawing.Point(376, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(240, 16);
			this.label8.TabIndex = 755;
			this.label8.Text = "VENTAS";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label8.Visible = false;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(856, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(93, 16);
			this.label6.TabIndex = 754;
			this.label6.Text = "% Ut CrediPOINT";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label6.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(776, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(67, 16);
			this.label5.TabIndex = 753;
			this.label5.Text = "% Ut Tarjeta";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Visible = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(696, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(72, 16);
			this.label4.TabIndex = 752;
			this.label4.Text = "% Ut Efectivo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Visible = false;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(536, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(79, 16);
			this.label3.TabIndex = 751;
			this.label3.Text = "% CrediPOINT";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label3.Visible = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(456, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(53, 16);
			this.label2.TabIndex = 750;
			this.label2.Text = "% Tarjeta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(376, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(58, 16);
			this.label1.TabIndex = 749;
			this.label1.Text = "% Efectivo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Visible = false;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(192, 16);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 741;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 40);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 740;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(8, 16);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 739;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcUtCrediPoint
			// 
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtCrediPoint.Appearance = appearance5;
			this.txtPorcUtCrediPoint.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtCrediPoint.Enabled = false;
			this.txtPorcUtCrediPoint.FormatString = "#,##0.00";
			this.txtPorcUtCrediPoint.Location = new System.Drawing.Point(856, 56);
			this.txtPorcUtCrediPoint.MinValue = -2147483647;
			this.txtPorcUtCrediPoint.Name = "txtPorcUtCrediPoint";
			this.txtPorcUtCrediPoint.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtCrediPoint.PromptChar = ' ';
			this.txtPorcUtCrediPoint.Size = new System.Drawing.Size(80, 21);
			this.txtPorcUtCrediPoint.TabIndex = 748;
			this.txtPorcUtCrediPoint.Visible = false;
			// 
			// txtPorcUtEfectivo
			// 
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtEfectivo.Appearance = appearance6;
			this.txtPorcUtEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtEfectivo.Enabled = false;
			this.txtPorcUtEfectivo.FormatString = "#,##0.00";
			this.txtPorcUtEfectivo.Location = new System.Drawing.Point(696, 56);
			this.txtPorcUtEfectivo.MinValue = -2147483647;
			this.txtPorcUtEfectivo.Name = "txtPorcUtEfectivo";
			this.txtPorcUtEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtEfectivo.PromptChar = ' ';
			this.txtPorcUtEfectivo.Size = new System.Drawing.Size(72, 21);
			this.txtPorcUtEfectivo.TabIndex = 746;
			this.txtPorcUtEfectivo.Visible = false;
			// 
			// txtPorcUtTarjeta
			// 
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtTarjeta.Appearance = appearance7;
			this.txtPorcUtTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtTarjeta.Enabled = false;
			this.txtPorcUtTarjeta.FormatString = "#,##0.00";
			this.txtPorcUtTarjeta.Location = new System.Drawing.Point(776, 56);
			this.txtPorcUtTarjeta.MinValue = -2147483647;
			this.txtPorcUtTarjeta.Name = "txtPorcUtTarjeta";
			this.txtPorcUtTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtTarjeta.PromptChar = ' ';
			this.txtPorcUtTarjeta.Size = new System.Drawing.Size(72, 21);
			this.txtPorcUtTarjeta.TabIndex = 747;
			this.txtPorcUtTarjeta.Visible = false;
			// 
			// txtPorcCrediPoint
			// 
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcCrediPoint.Appearance = appearance8;
			this.txtPorcCrediPoint.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcCrediPoint.Enabled = false;
			this.txtPorcCrediPoint.FormatString = "#,##0.00";
			this.txtPorcCrediPoint.Location = new System.Drawing.Point(536, 56);
			this.txtPorcCrediPoint.MinValue = -2147483647;
			this.txtPorcCrediPoint.Name = "txtPorcCrediPoint";
			this.txtPorcCrediPoint.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcCrediPoint.PromptChar = ' ';
			this.txtPorcCrediPoint.Size = new System.Drawing.Size(80, 21);
			this.txtPorcCrediPoint.TabIndex = 745;
			this.txtPorcCrediPoint.Visible = false;
			// 
			// txtPorcEfectivo
			// 
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcEfectivo.Appearance = appearance9;
			this.txtPorcEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcEfectivo.Enabled = false;
			this.txtPorcEfectivo.FormatString = "#,##0.00";
			this.txtPorcEfectivo.Location = new System.Drawing.Point(376, 56);
			this.txtPorcEfectivo.MinValue = -2147483647;
			this.txtPorcEfectivo.Name = "txtPorcEfectivo";
			this.txtPorcEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcEfectivo.PromptChar = ' ';
			this.txtPorcEfectivo.Size = new System.Drawing.Size(72, 21);
			this.txtPorcEfectivo.TabIndex = 743;
			this.txtPorcEfectivo.Visible = false;
			// 
			// txtPorcTarjeta
			// 
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcTarjeta.Appearance = appearance10;
			this.txtPorcTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcTarjeta.Enabled = false;
			this.txtPorcTarjeta.FormatString = "#,##0.00";
			this.txtPorcTarjeta.Location = new System.Drawing.Point(456, 56);
			this.txtPorcTarjeta.MinValue = -2147483647;
			this.txtPorcTarjeta.Name = "txtPorcTarjeta";
			this.txtPorcTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcTarjeta.PromptChar = ' ';
			this.txtPorcTarjeta.Size = new System.Drawing.Size(72, 21);
			this.txtPorcTarjeta.TabIndex = 744;
			this.txtPorcTarjeta.Visible = false;
			// 
			// txtBusqueda
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance11;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(56, 40);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda.TabIndex = 734;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1256, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 736;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(280, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 735;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtDesde
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance12;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 737;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// dtHasta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance13;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 738;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(944, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 742;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(952, 40);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(43, 16);
			this.label10.TabIndex = 758;
			this.label10.Text = "Margen";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label10.Visible = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1378, 8);
			this.groupBox1.TabIndex = 765;
			this.groupBox1.TabStop = false;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource2;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance14;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 151;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn2.CellAppearance = appearance15;
			ultraGridColumn2.Format = "#,##0";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance16;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance17;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "% Ventas";
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance18;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "V Efectivo";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance19;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "% Efectivo";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance20;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "V Tarjeta";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance21;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Tarjeta";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance22;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "V CrediPoint";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance23;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "% CrediPoint";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance24;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Total Utilidad";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance25;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "% Utilidad";
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance26;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "U Efectivo";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 80;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance27;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "% U Efectivo";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance28;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "U Tarjeta";
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 80;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance29;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "% U Tarjeta";
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance30;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "U CrediPoint";
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance31;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% U CrediPoint";
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance32;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance33;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "M Efectivo";
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance34;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "M Tarjeta";
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn21.Width = 80;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance35;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "M CrediPoint";
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
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
																										 ultraGridColumn23});
			ultraGridBand1.Header.Caption = "Periodo A";
			ultraGridBand1.HeaderVisible = true;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance36;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance37;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance38;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance39.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance39;
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance41.FontData.BoldAsString = "True";
			appearance41.FontData.Name = "Arial";
			appearance41.FontData.SizeInPoints = 8F;
			appearance41.ForeColor = System.Drawing.Color.White;
			appearance41.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance41;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance43;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(224, 80);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1120, 120);
			this.uGridInformacion.TabIndex = 766;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
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
			ultraDataColumn23.DataType = typeof(int);
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
																																 ultraDataColumn23});
			// 
			// chkTodos
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance44;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Checked = true;
			this.chkTodos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodos.Location = new System.Drawing.Point(200, 88);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 16);
			this.chkTodos.TabIndex = 768;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// uGridLocales
			// 
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource1;
			appearance45.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance45;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 179;
			ultraGridColumn26.Header.Caption = "...";
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn26.Width = 21;
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 45;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			ultraGridBand2.Header.Caption = "";
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance48.FontData.BoldAsString = "True";
			appearance48.FontData.Name = "Arial";
			appearance48.FontData.SizeInPoints = 10F;
			appearance48.ForeColor = System.Drawing.Color.White;
			appearance48.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance48;
			appearance49.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance49.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance49;
			appearance50.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance50.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance51;
			this.uGridLocales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(0, 80);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(221, 120);
			this.uGridLocales.TabIndex = 767;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(bool);
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27});
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 200);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1378, 8);
			this.groupBox2.TabIndex = 769;
			this.groupBox2.TabStop = false;
			// 
			// txtMargenCrediPOINT2
			// 
			appearance52.ForeColor = System.Drawing.Color.Black;
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenCrediPOINT2.Appearance = appearance52;
			this.txtMargenCrediPOINT2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenCrediPOINT2.Enabled = false;
			this.txtMargenCrediPOINT2.FormatString = "#,##0.00";
			this.txtMargenCrediPOINT2.Location = new System.Drawing.Point(1208, 248);
			this.txtMargenCrediPOINT2.MinValue = -2147483647;
			this.txtMargenCrediPOINT2.Name = "txtMargenCrediPOINT2";
			this.txtMargenCrediPOINT2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenCrediPOINT2.PromptChar = ' ';
			this.txtMargenCrediPOINT2.Size = new System.Drawing.Size(80, 21);
			this.txtMargenCrediPOINT2.TabIndex = 798;
			this.txtMargenCrediPOINT2.Visible = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(1208, 232);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(107, 16);
			this.label14.TabIndex = 799;
			this.label14.Text = "Margen CrediPOINT";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Visible = false;
			// 
			// txtMargenTarjeta2
			// 
			appearance53.ForeColor = System.Drawing.Color.Black;
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenTarjeta2.Appearance = appearance53;
			this.txtMargenTarjeta2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenTarjeta2.Enabled = false;
			this.txtMargenTarjeta2.FormatString = "#,##0.00";
			this.txtMargenTarjeta2.Location = new System.Drawing.Point(1120, 248);
			this.txtMargenTarjeta2.MinValue = -2147483647;
			this.txtMargenTarjeta2.Name = "txtMargenTarjeta2";
			this.txtMargenTarjeta2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenTarjeta2.PromptChar = ' ';
			this.txtMargenTarjeta2.Size = new System.Drawing.Size(80, 21);
			this.txtMargenTarjeta2.TabIndex = 796;
			this.txtMargenTarjeta2.Visible = false;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(1120, 232);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(81, 16);
			this.label15.TabIndex = 797;
			this.label15.Text = "Margen Tarjeta";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Visible = false;
			// 
			// txtMargenEfectivo2
			// 
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenEfectivo2.Appearance = appearance54;
			this.txtMargenEfectivo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenEfectivo2.Enabled = false;
			this.txtMargenEfectivo2.FormatString = "#,##0.00";
			this.txtMargenEfectivo2.Location = new System.Drawing.Point(1032, 248);
			this.txtMargenEfectivo2.MinValue = -2147483647;
			this.txtMargenEfectivo2.Name = "txtMargenEfectivo2";
			this.txtMargenEfectivo2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenEfectivo2.PromptChar = ' ';
			this.txtMargenEfectivo2.Size = new System.Drawing.Size(80, 21);
			this.txtMargenEfectivo2.TabIndex = 794;
			this.txtMargenEfectivo2.Visible = false;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(1032, 232);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(86, 16);
			this.label16.TabIndex = 795;
			this.label16.Text = "Margen Efectivo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Visible = false;
			// 
			// txtMargen2
			// 
			appearance55.ForeColor = System.Drawing.Color.Black;
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargen2.Appearance = appearance55;
			this.txtMargen2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargen2.Enabled = false;
			this.txtMargen2.FormatString = "#,##0.00";
			this.txtMargen2.Location = new System.Drawing.Point(944, 248);
			this.txtMargen2.MinValue = -2147483647;
			this.txtMargen2.Name = "txtMargen2";
			this.txtMargen2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargen2.PromptChar = ' ';
			this.txtMargen2.Size = new System.Drawing.Size(80, 21);
			this.txtMargen2.TabIndex = 792;
			this.txtMargen2.Visible = false;
			// 
			// label17
			// 
			this.label17.BackColor = System.Drawing.Color.DarkGray;
			this.label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.ForeColor = System.Drawing.Color.Firebrick;
			this.label17.Location = new System.Drawing.Point(688, 208);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(240, 16);
			this.label17.TabIndex = 791;
			this.label17.Text = "UTILIDAD";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label17.Visible = false;
			// 
			// label18
			// 
			this.label18.BackColor = System.Drawing.Color.DarkGray;
			this.label18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.ForeColor = System.Drawing.Color.Firebrick;
			this.label18.Location = new System.Drawing.Point(368, 208);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(240, 16);
			this.label18.TabIndex = 790;
			this.label18.Text = "VENTAS";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.label18.Visible = false;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(848, 232);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(93, 16);
			this.label19.TabIndex = 789;
			this.label19.Text = "% Ut CrediPOINT";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label19.Visible = false;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(768, 232);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(67, 16);
			this.label20.TabIndex = 788;
			this.label20.Text = "% Ut Tarjeta";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label20.Visible = false;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(688, 232);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(72, 16);
			this.label21.TabIndex = 787;
			this.label21.Text = "% Ut Efectivo";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label21.Visible = false;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(528, 232);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(79, 16);
			this.label22.TabIndex = 786;
			this.label22.Text = "% CrediPOINT";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label22.Visible = false;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(448, 232);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(53, 16);
			this.label23.TabIndex = 785;
			this.label23.Text = "% Tarjeta";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label23.Visible = false;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(368, 232);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(58, 16);
			this.label24.TabIndex = 784;
			this.label24.Text = "% Efectivo";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label24.Visible = false;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(184, 208);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(34, 16);
			this.label25.TabIndex = 777;
			this.label25.Text = "Hasta";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label26.Location = new System.Drawing.Point(-40, 416);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(39, 16);
			this.label26.TabIndex = 776;
			this.label26.Text = "Buscar";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(-32, 360);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(37, 16);
			this.label27.TabIndex = 775;
			this.label27.Text = "Desde";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcUtCrediPoint2
			// 
			appearance56.ForeColor = System.Drawing.Color.Black;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtCrediPoint2.Appearance = appearance56;
			this.txtPorcUtCrediPoint2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtCrediPoint2.Enabled = false;
			this.txtPorcUtCrediPoint2.FormatString = "#,##0.00";
			this.txtPorcUtCrediPoint2.Location = new System.Drawing.Point(848, 248);
			this.txtPorcUtCrediPoint2.MinValue = -2147483647;
			this.txtPorcUtCrediPoint2.Name = "txtPorcUtCrediPoint2";
			this.txtPorcUtCrediPoint2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtCrediPoint2.PromptChar = ' ';
			this.txtPorcUtCrediPoint2.Size = new System.Drawing.Size(80, 21);
			this.txtPorcUtCrediPoint2.TabIndex = 783;
			this.txtPorcUtCrediPoint2.Visible = false;
			// 
			// txtPorcUtEfectivo2
			// 
			appearance57.ForeColor = System.Drawing.Color.Black;
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtEfectivo2.Appearance = appearance57;
			this.txtPorcUtEfectivo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtEfectivo2.Enabled = false;
			this.txtPorcUtEfectivo2.FormatString = "#,##0.00";
			this.txtPorcUtEfectivo2.Location = new System.Drawing.Point(688, 248);
			this.txtPorcUtEfectivo2.MinValue = -2147483647;
			this.txtPorcUtEfectivo2.Name = "txtPorcUtEfectivo2";
			this.txtPorcUtEfectivo2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtEfectivo2.PromptChar = ' ';
			this.txtPorcUtEfectivo2.Size = new System.Drawing.Size(72, 21);
			this.txtPorcUtEfectivo2.TabIndex = 781;
			this.txtPorcUtEfectivo2.Visible = false;
			// 
			// txtPorcUtTarjeta2
			// 
			appearance58.ForeColor = System.Drawing.Color.Black;
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcUtTarjeta2.Appearance = appearance58;
			this.txtPorcUtTarjeta2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcUtTarjeta2.Enabled = false;
			this.txtPorcUtTarjeta2.FormatString = "#,##0.00";
			this.txtPorcUtTarjeta2.Location = new System.Drawing.Point(768, 248);
			this.txtPorcUtTarjeta2.MinValue = -2147483647;
			this.txtPorcUtTarjeta2.Name = "txtPorcUtTarjeta2";
			this.txtPorcUtTarjeta2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcUtTarjeta2.PromptChar = ' ';
			this.txtPorcUtTarjeta2.Size = new System.Drawing.Size(72, 21);
			this.txtPorcUtTarjeta2.TabIndex = 782;
			this.txtPorcUtTarjeta2.Visible = false;
			// 
			// txtPorcCrediPoint2
			// 
			appearance59.ForeColor = System.Drawing.Color.Black;
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcCrediPoint2.Appearance = appearance59;
			this.txtPorcCrediPoint2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcCrediPoint2.Enabled = false;
			this.txtPorcCrediPoint2.FormatString = "#,##0.00";
			this.txtPorcCrediPoint2.Location = new System.Drawing.Point(528, 248);
			this.txtPorcCrediPoint2.MinValue = -2147483647;
			this.txtPorcCrediPoint2.Name = "txtPorcCrediPoint2";
			this.txtPorcCrediPoint2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcCrediPoint2.PromptChar = ' ';
			this.txtPorcCrediPoint2.Size = new System.Drawing.Size(80, 21);
			this.txtPorcCrediPoint2.TabIndex = 780;
			this.txtPorcCrediPoint2.Visible = false;
			// 
			// txtPorcEfectivo2
			// 
			appearance60.ForeColor = System.Drawing.Color.Black;
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcEfectivo2.Appearance = appearance60;
			this.txtPorcEfectivo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcEfectivo2.Enabled = false;
			this.txtPorcEfectivo2.FormatString = "#,##0.00";
			this.txtPorcEfectivo2.Location = new System.Drawing.Point(368, 248);
			this.txtPorcEfectivo2.MinValue = -2147483647;
			this.txtPorcEfectivo2.Name = "txtPorcEfectivo2";
			this.txtPorcEfectivo2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcEfectivo2.PromptChar = ' ';
			this.txtPorcEfectivo2.Size = new System.Drawing.Size(72, 21);
			this.txtPorcEfectivo2.TabIndex = 778;
			this.txtPorcEfectivo2.Visible = false;
			// 
			// txtPorcTarjeta2
			// 
			appearance61.ForeColor = System.Drawing.Color.Black;
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPorcTarjeta2.Appearance = appearance61;
			this.txtPorcTarjeta2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcTarjeta2.Enabled = false;
			this.txtPorcTarjeta2.FormatString = "#,##0.00";
			this.txtPorcTarjeta2.Location = new System.Drawing.Point(448, 248);
			this.txtPorcTarjeta2.MinValue = -2147483647;
			this.txtPorcTarjeta2.Name = "txtPorcTarjeta2";
			this.txtPorcTarjeta2.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPorcTarjeta2.PromptChar = ' ';
			this.txtPorcTarjeta2.Size = new System.Drawing.Size(72, 21);
			this.txtPorcTarjeta2.TabIndex = 779;
			this.txtPorcTarjeta2.Visible = false;
			// 
			// txtBusqueda2
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda2.Appearance = appearance62;
			this.txtBusqueda2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda2.Location = new System.Drawing.Point(48, 240);
			this.txtBusqueda2.Name = "txtBusqueda2";
			this.txtBusqueda2.Size = new System.Drawing.Size(192, 21);
			this.txtBusqueda2.TabIndex = 770;
			// 
			// btnExcel2
			// 
			this.btnExcel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel2.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel2.Image")));
			this.btnExcel2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel2.Location = new System.Drawing.Point(1256, 208);
			this.btnExcel2.Name = "btnExcel2";
			this.btnExcel2.Size = new System.Drawing.Size(75, 24);
			this.btnExcel2.TabIndex = 772;
			this.btnExcel2.Text = "Excel";
			this.btnExcel2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel2.Visible = false;
			this.btnExcel2.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnVer2
			// 
			this.btnVer2.Image = ((System.Drawing.Image)(resources.GetObject("btnVer2.Image")));
			this.btnVer2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer2.Location = new System.Drawing.Point(248, 240);
			this.btnVer2.Name = "btnVer2";
			this.btnVer2.Size = new System.Drawing.Size(75, 24);
			this.btnVer2.TabIndex = 771;
			this.btnVer2.Text = "Ver";
			this.btnVer2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer2.Click += new System.EventHandler(this.btnVer2_Click);
			// 
			// dtDesde2
			// 
			appearance63.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde2.Appearance = appearance63;
			this.dtDesde2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtDesde2.DateButtons.Add(dateButton3);
			this.dtDesde2.Format = "dd/MM/yyyy";
			this.dtDesde2.Location = new System.Drawing.Point(48, 208);
			this.dtDesde2.Name = "dtDesde2";
			this.dtDesde2.NonAutoSizeHeight = 23;
			this.dtDesde2.Size = new System.Drawing.Size(112, 21);
			this.dtDesde2.SpinButtonsVisible = true;
			this.dtDesde2.TabIndex = 773;
			this.dtDesde2.Value = ((object)(resources.GetObject("dtDesde2.Value")));
			// 
			// dtHasta2
			// 
			appearance64.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta2.Appearance = appearance64;
			this.dtHasta2.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtHasta2.DateButtons.Add(dateButton4);
			this.dtHasta2.Format = "dd/MM/yyyy";
			this.dtHasta2.Location = new System.Drawing.Point(240, 208);
			this.dtHasta2.Name = "dtHasta2";
			this.dtHasta2.NonAutoSizeHeight = 23;
			this.dtHasta2.Size = new System.Drawing.Size(112, 21);
			this.dtHasta2.SpinButtonsVisible = true;
			this.dtHasta2.TabIndex = 774;
			this.dtHasta2.Value = ((object)(resources.GetObject("dtHasta2.Value")));
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(944, 232);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(43, 16);
			this.label28.TabIndex = 793;
			this.label28.Text = "Margen";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label28.Visible = false;
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(8, 208);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(37, 16);
			this.label29.TabIndex = 800;
			this.label29.Text = "Desde";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label30.Location = new System.Drawing.Point(8, 240);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(39, 16);
			this.label30.TabIndex = 801;
			this.label30.Text = "Buscar";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador2
			// 
			this.lblContador2.AutoSize = true;
			this.lblContador2.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador2.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador2.Location = new System.Drawing.Point(944, 208);
			this.lblContador2.Name = "lblContador2";
			this.lblContador2.Size = new System.Drawing.Size(0, 20);
			this.lblContador2.TabIndex = 802;
			this.lblContador2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridInformacion2
			// 
			this.uGridInformacion2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion2.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion2.DataSource = this.ultraDataSource2;
			appearance65.BackColor = System.Drawing.Color.White;
			this.uGridInformacion2.DisplayLayout.Appearance = appearance65;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 151;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance66;
			ultraGridColumn29.Format = "#,##0";
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 60;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance67;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 80;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance68;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "% Ventas";
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 80;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance69.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance69;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "V Efectivo";
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 80;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance70.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance70;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.Caption = "% Efectivo";
			ultraGridColumn33.Header.VisiblePosition = 6;
			ultraGridColumn33.Width = 80;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance71.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn34.CellAppearance = appearance71;
			ultraGridColumn34.Format = "#,##0.00";
			ultraGridColumn34.Header.Caption = "V Tarjeta";
			ultraGridColumn34.Header.VisiblePosition = 7;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 80;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance72.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn35.CellAppearance = appearance72;
			ultraGridColumn35.Format = "#,##0.00";
			ultraGridColumn35.Header.Caption = "% Tarjeta";
			ultraGridColumn35.Header.VisiblePosition = 8;
			ultraGridColumn35.Width = 80;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance73.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance73;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.Caption = "V CrediPoint";
			ultraGridColumn36.Header.VisiblePosition = 9;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 80;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance74.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance74;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "% CrediPoint";
			ultraGridColumn37.Header.VisiblePosition = 10;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance75.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance75;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.Caption = "Total Utilidad";
			ultraGridColumn38.Header.VisiblePosition = 11;
			ultraGridColumn38.Width = 80;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance76.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance76;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.Caption = "% Utilidad";
			ultraGridColumn39.Header.VisiblePosition = 13;
			ultraGridColumn39.Width = 80;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance77.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance77;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.Caption = "U Efectivo";
			ultraGridColumn40.Header.VisiblePosition = 12;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 80;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance78.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance78;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.Caption = "% U Efectivo";
			ultraGridColumn41.Header.VisiblePosition = 14;
			ultraGridColumn41.Width = 80;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance79.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance79;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.Caption = "U Tarjeta";
			ultraGridColumn42.Header.VisiblePosition = 15;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 80;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance80.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance80;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.Caption = "% U Tarjeta";
			ultraGridColumn43.Header.VisiblePosition = 16;
			ultraGridColumn43.Width = 80;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance81.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance81;
			ultraGridColumn44.Format = "#,##0.00";
			ultraGridColumn44.Header.Caption = "U CrediPoint";
			ultraGridColumn44.Header.VisiblePosition = 17;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 80;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance82.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance82;
			ultraGridColumn45.Format = "#,##0.00";
			ultraGridColumn45.Header.Caption = "% U CrediPoint";
			ultraGridColumn45.Header.VisiblePosition = 18;
			ultraGridColumn45.Width = 80;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance83.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance83;
			ultraGridColumn46.Format = "#,##0.00";
			ultraGridColumn46.Header.VisiblePosition = 19;
			ultraGridColumn46.Width = 80;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance84.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn47.CellAppearance = appearance84;
			ultraGridColumn47.Format = "#,##0.00";
			ultraGridColumn47.Header.Caption = "M Efectivo";
			ultraGridColumn47.Header.VisiblePosition = 20;
			ultraGridColumn47.Width = 80;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance85.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn48.CellAppearance = appearance85;
			ultraGridColumn48.Format = "#,##0.00";
			ultraGridColumn48.Header.Caption = "M Tarjeta";
			ultraGridColumn48.Header.VisiblePosition = 21;
			ultraGridColumn48.Width = 80;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance86.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn49.CellAppearance = appearance86;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "M CrediPoint";
			ultraGridColumn49.Header.VisiblePosition = 22;
			ultraGridColumn49.Width = 80;
			ultraGridColumn50.Header.VisiblePosition = 0;
			ultraGridColumn50.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn50});
			ultraGridBand3.Header.Caption = "Periodo B";
			ultraGridBand3.HeaderVisible = true;
			appearance87.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance87;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance88.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance88;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance89.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance89;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridInformacion2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridInformacion2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance90.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion2.DisplayLayout.Override.CardAreaAppearance = appearance90;
			appearance91.ForeColor = System.Drawing.Color.Black;
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion2.DisplayLayout.Override.CellAppearance = appearance91;
			appearance92.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance92.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance92.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance92.FontData.BoldAsString = "True";
			appearance92.FontData.Name = "Arial";
			appearance92.FontData.SizeInPoints = 8F;
			appearance92.ForeColor = System.Drawing.Color.White;
			appearance92.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion2.DisplayLayout.Override.HeaderAppearance = appearance92;
			this.uGridInformacion2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance93.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance93.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance93.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion2.DisplayLayout.Override.RowSelectorAppearance = appearance93;
			appearance94.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance94.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance94.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion2.DisplayLayout.Override.SelectedRowAppearance = appearance94;
			this.uGridInformacion2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion2.Location = new System.Drawing.Point(224, 272);
			this.uGridInformacion2.Name = "uGridInformacion2";
			this.uGridInformacion2.Size = new System.Drawing.Size(1120, 112);
			this.uGridInformacion2.TabIndex = 804;
			// 
			// chkTodos2
			// 
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			appearance95.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos2.Appearance = appearance95;
			this.chkTodos2.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos2.Checked = true;
			this.chkTodos2.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodos2.Location = new System.Drawing.Point(208, 272);
			this.chkTodos2.Name = "chkTodos2";
			this.chkTodos2.Size = new System.Drawing.Size(13, 16);
			this.chkTodos2.TabIndex = 806;
			this.chkTodos2.CheckedChanged += new System.EventHandler(this.chkTodos2_CheckedChanged);
			// 
			// uGridLocales2
			// 
			this.uGridLocales2.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales2.DataSource = this.ultraDataSource1;
			appearance96.BackColor = System.Drawing.Color.White;
			this.uGridLocales2.DisplayLayout.Appearance = appearance96;
			this.uGridLocales2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 100;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 179;
			ultraGridColumn53.Header.Caption = "...";
			ultraGridColumn53.Header.VisiblePosition = 2;
			ultraGridColumn53.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn53.Width = 21;
			ultraGridColumn54.Header.VisiblePosition = 3;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 45;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54});
			ultraGridBand4.Header.Caption = "";
			ultraGridBand4.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales2.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance97.ForeColor = System.Drawing.Color.Black;
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales2.DisplayLayout.Override.ActiveRowAppearance = appearance97;
			this.uGridLocales2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridLocales2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridLocales2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance98.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales2.DisplayLayout.Override.CardAreaAppearance = appearance98;
			appearance99.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance99.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance99.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance99.FontData.BoldAsString = "True";
			appearance99.FontData.Name = "Arial";
			appearance99.FontData.SizeInPoints = 10F;
			appearance99.ForeColor = System.Drawing.Color.White;
			appearance99.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales2.DisplayLayout.Override.HeaderAppearance = appearance99;
			appearance100.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance100.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance100.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales2.DisplayLayout.Override.RowAlternateAppearance = appearance100;
			appearance101.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance101.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance101.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales2.DisplayLayout.Override.RowSelectorAppearance = appearance101;
			appearance102.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance102.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance102.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales2.DisplayLayout.Override.SelectedRowAppearance = appearance102;
			this.uGridLocales2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales2.Location = new System.Drawing.Point(0, 272);
			this.uGridLocales2.Name = "uGridLocales2";
			this.uGridLocales2.Size = new System.Drawing.Size(221, 112);
			this.uGridLocales2.TabIndex = 805;
			// 
			// groupBox4
			// 
			this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox4.Location = new System.Drawing.Point(8, 384);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1378, 8);
			this.groupBox4.TabIndex = 807;
			this.groupBox4.TabStop = false;
			// 
			// uGridporcentaje
			// 
			this.uGridporcentaje.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridporcentaje.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridporcentaje.DataSource = this.ultraDataSource4;
			appearance103.BackColor = System.Drawing.Color.White;
			this.uGridporcentaje.DisplayLayout.Appearance = appearance103;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Width = 150;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance104;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridColumn56.Width = 80;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance105;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.Caption = "%Ventas";
			ultraGridColumn57.Header.VisiblePosition = 2;
			ultraGridColumn57.Width = 80;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance106;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "%Efectivo";
			ultraGridColumn58.Header.VisiblePosition = 3;
			ultraGridColumn58.Width = 80;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn59.CellAppearance = appearance107;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.Caption = "%Tarjeta";
			ultraGridColumn59.Header.VisiblePosition = 4;
			ultraGridColumn59.Width = 80;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn60.CellAppearance = appearance108;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.Caption = "%CrediPoint";
			ultraGridColumn60.Header.VisiblePosition = 5;
			ultraGridColumn60.Width = 80;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn61.CellAppearance = appearance109;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.Caption = "Total Utilidad";
			ultraGridColumn61.Header.VisiblePosition = 6;
			ultraGridColumn61.Width = 80;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn62.CellAppearance = appearance110;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "%Utilidad";
			ultraGridColumn62.Header.VisiblePosition = 7;
			ultraGridColumn62.Width = 80;
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn63.CellAppearance = appearance111;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.Caption = "%U Efectivo";
			ultraGridColumn63.Header.VisiblePosition = 8;
			ultraGridColumn63.Width = 80;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn64.CellAppearance = appearance112;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.Caption = "%U Tarjeta";
			ultraGridColumn64.Header.VisiblePosition = 9;
			ultraGridColumn64.Width = 80;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn65.CellAppearance = appearance113;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.Caption = "%U CrediPoint";
			ultraGridColumn65.Header.VisiblePosition = 10;
			ultraGridColumn65.Width = 80;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn66.CellAppearance = appearance114;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.VisiblePosition = 11;
			ultraGridColumn66.Width = 80;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn67.CellAppearance = appearance115;
			ultraGridColumn67.Format = "#,##0.00";
			ultraGridColumn67.Header.Caption = "M Efectivo";
			ultraGridColumn67.Header.VisiblePosition = 12;
			ultraGridColumn67.Width = 80;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn68.CellAppearance = appearance116;
			ultraGridColumn68.Format = "#,##0.00";
			ultraGridColumn68.Header.Caption = "M Tarjeta";
			ultraGridColumn68.Header.VisiblePosition = 13;
			ultraGridColumn68.Width = 80;
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn69.CellAppearance = appearance117;
			ultraGridColumn69.Format = "#,##0.00";
			ultraGridColumn69.Header.Caption = "M CrediPoint";
			ultraGridColumn69.Header.VisiblePosition = 14;
			ultraGridColumn69.Width = 80;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn69});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridporcentaje.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridporcentaje.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridporcentaje.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridporcentaje.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance118.BackColor = System.Drawing.Color.Transparent;
			this.uGridporcentaje.DisplayLayout.Override.CardAreaAppearance = appearance118;
			appearance119.ForeColor = System.Drawing.Color.Black;
			appearance119.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridporcentaje.DisplayLayout.Override.CellAppearance = appearance119;
			appearance120.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance120.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance120.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance120.FontData.BoldAsString = "True";
			appearance120.FontData.Name = "Arial";
			appearance120.FontData.SizeInPoints = 8F;
			appearance120.ForeColor = System.Drawing.Color.White;
			appearance120.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridporcentaje.DisplayLayout.Override.HeaderAppearance = appearance120;
			this.uGridporcentaje.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance121.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance121.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridporcentaje.DisplayLayout.Override.RowSelectorAppearance = appearance121;
			appearance122.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance122.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance122.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridporcentaje.DisplayLayout.Override.SelectedRowAppearance = appearance122;
			this.uGridporcentaje.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridporcentaje.Location = new System.Drawing.Point(264, 392);
			this.uGridporcentaje.Name = "uGridporcentaje";
			this.uGridporcentaje.Size = new System.Drawing.Size(1080, 96);
			this.uGridporcentaje.TabIndex = 808;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn42});
			// 
			// button3
			// 
			this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
			this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button3.Location = new System.Drawing.Point(16, 400);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(75, 24);
			this.button3.TabIndex = 809;
			this.button3.Text = "Ver";
			this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button3.Click += new System.EventHandler(this.button3_Click);
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
			// button4
			// 
			this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
			this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button4.Location = new System.Drawing.Point(96, 400);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(75, 24);
			this.button4.TabIndex = 810;
			this.button4.Text = "Excel";
			this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button4.Visible = false;
			this.button4.Click += new System.EventHandler(this.button4_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn44.DefaultValue = 0;
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
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn64});
			// 
			// ugridTotales
			// 
			this.ugridTotales.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridTotales.DataSource = this.ultraDataSource6;
			appearance123.BackColor = System.Drawing.Color.White;
			this.ugridTotales.DisplayLayout.Appearance = appearance123;
			ultraGridColumn70.Header.VisiblePosition = 0;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.Caption = "Concepto";
			ultraGridColumn71.Header.VisiblePosition = 1;
			ultraGridColumn71.Width = 200;
			ultraGridColumn72.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance124.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance124;
			ultraGridColumn72.Format = "#,##0.00";
			ultraGridColumn72.Header.Caption = "Total";
			ultraGridColumn72.Header.VisiblePosition = 2;
			ultraGridColumn72.Width = 110;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72});
			ultraGridBand6.Header.Caption = "Ventas Por Año";
			ultraGridBand6.HeaderVisible = true;
			ultraGridBand6.SummaryFooterCaption = "";
			this.ugridTotales.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ugridTotales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridTotales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridTotales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance125.BackColor = System.Drawing.Color.Transparent;
			this.ugridTotales.DisplayLayout.Override.CardAreaAppearance = appearance125;
			appearance126.ForeColor = System.Drawing.Color.Black;
			appearance126.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridTotales.DisplayLayout.Override.CellAppearance = appearance126;
			appearance127.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance127.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance127.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance127.FontData.BoldAsString = "True";
			appearance127.FontData.Name = "Arial";
			appearance127.FontData.SizeInPoints = 8F;
			appearance127.ForeColor = System.Drawing.Color.White;
			appearance127.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridTotales.DisplayLayout.Override.HeaderAppearance = appearance127;
			this.ugridTotales.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance128.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance128.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance128.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridTotales.DisplayLayout.Override.RowSelectorAppearance = appearance128;
			appearance129.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance129.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance129.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridTotales.DisplayLayout.Override.SelectedRowAppearance = appearance129;
			this.ugridTotales.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridTotales.Location = new System.Drawing.Point(264, 496);
			this.ugridTotales.Name = "ugridTotales";
			this.ugridTotales.Size = new System.Drawing.Size(368, 88);
			this.ugridTotales.TabIndex = 812;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn67.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn65,
																																 ultraDataColumn66,
																																 ultraDataColumn67});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource5;
			appearance130.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance130;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn73.Header.VisiblePosition = 0;
			ultraGridColumn73.Width = 8;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn73});
			ultraGridBand7.Header.Caption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			appearance131.ForeColor = System.Drawing.Color.Black;
			appearance131.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance131;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance132.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance132;
			appearance133.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance133.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance133.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance133.FontData.BoldAsString = "True";
			appearance133.FontData.Name = "Arial";
			appearance133.FontData.SizeInPoints = 10F;
			appearance133.ForeColor = System.Drawing.Color.White;
			appearance133.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance133;
			appearance134.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance134.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance134.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance134;
			appearance135.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance135.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance135.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance135;
			appearance136.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance136.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance136.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance136;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(192, 448);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(16, 24);
			this.ultraGrid1.TabIndex = 813;
			this.ultraGrid1.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn68.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn68});
			// 
			// ugridPorcentajes
			// 
			this.ugridPorcentajes.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridPorcentajes.DataSource = this.ultraDataSource7;
			appearance137.BackColor = System.Drawing.Color.White;
			this.ugridPorcentajes.DisplayLayout.Appearance = appearance137;
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.Caption = "Concepto";
			ultraGridColumn75.Header.VisiblePosition = 1;
			ultraGridColumn75.Width = 180;
			appearance138.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance138;
			ultraGridColumn76.Format = "#,##0.00";
			ultraGridColumn76.Header.Caption = "Porcentajes";
			ultraGridColumn76.Header.VisiblePosition = 2;
			ultraGridColumn76.Width = 100;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76});
			ultraGridBand8.Header.Caption = "Porcentaje de Ventas";
			ultraGridBand8.HeaderVisible = true;
			this.ugridPorcentajes.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.ugridPorcentajes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridPorcentajes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridPorcentajes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance139.BackColor = System.Drawing.Color.Transparent;
			this.ugridPorcentajes.DisplayLayout.Override.CardAreaAppearance = appearance139;
			appearance140.ForeColor = System.Drawing.Color.Black;
			appearance140.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentajes.DisplayLayout.Override.CellAppearance = appearance140;
			appearance141.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance141.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance141.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance141.FontData.BoldAsString = "True";
			appearance141.FontData.Name = "Arial";
			appearance141.FontData.SizeInPoints = 8F;
			appearance141.ForeColor = System.Drawing.Color.White;
			appearance141.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridPorcentajes.DisplayLayout.Override.HeaderAppearance = appearance141;
			this.ugridPorcentajes.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance142.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance142.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance142.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentajes.DisplayLayout.Override.RowSelectorAppearance = appearance142;
			appearance143.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance143.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance143.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentajes.DisplayLayout.Override.SelectedRowAppearance = appearance143;
			this.ugridPorcentajes.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridPorcentajes.Location = new System.Drawing.Point(640, 496);
			this.ugridPorcentajes.Name = "ugridPorcentajes";
			this.ugridPorcentajes.Size = new System.Drawing.Size(336, 88);
			this.ugridPorcentajes.TabIndex = 814;
			// 
			// ultraDataSource7
			// 
			ultraDataColumn69.DataType = typeof(int);
			ultraDataColumn71.DataType = typeof(System.Decimal);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn69,
																																 ultraDataColumn70,
																																 ultraDataColumn71});
			// 
			// ugridIYD
			// 
			this.ugridIYD.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridIYD.DataSource = this.ultraDataSource8;
			appearance144.BackColor = System.Drawing.Color.White;
			this.ugridIYD.DisplayLayout.Appearance = appearance144;
			ultraGridColumn77.Header.VisiblePosition = 0;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.Caption = "Concepto";
			ultraGridColumn78.Header.VisiblePosition = 1;
			ultraGridColumn78.Width = 180;
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn79.CellAppearance = appearance145;
			ultraGridColumn79.Format = "#,##0.00";
			ultraGridColumn79.Header.Caption = "I Y D";
			ultraGridColumn79.Header.VisiblePosition = 2;
			ultraGridColumn79.Width = 100;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79});
			ultraGridBand9.Header.Caption = "Incremento Y Decremento en Ventas";
			ultraGridBand9.HeaderVisible = true;
			this.ugridIYD.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.ugridIYD.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridIYD.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridIYD.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance146.BackColor = System.Drawing.Color.Transparent;
			this.ugridIYD.DisplayLayout.Override.CardAreaAppearance = appearance146;
			appearance147.ForeColor = System.Drawing.Color.Black;
			appearance147.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridIYD.DisplayLayout.Override.CellAppearance = appearance147;
			appearance148.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance148.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance148.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance148.FontData.BoldAsString = "True";
			appearance148.FontData.Name = "Arial";
			appearance148.FontData.SizeInPoints = 8F;
			appearance148.ForeColor = System.Drawing.Color.White;
			appearance148.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridIYD.DisplayLayout.Override.HeaderAppearance = appearance148;
			this.ugridIYD.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance149.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance149.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance149.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridIYD.DisplayLayout.Override.RowSelectorAppearance = appearance149;
			appearance150.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance150.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance150.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridIYD.DisplayLayout.Override.SelectedRowAppearance = appearance150;
			this.ugridIYD.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridIYD.Location = new System.Drawing.Point(984, 496);
			this.ugridIYD.Name = "ugridIYD";
			this.ugridIYD.Size = new System.Drawing.Size(344, 88);
			this.ugridIYD.TabIndex = 815;
			// 
			// ultraDataSource8
			// 
			ultraDataColumn72.DataType = typeof(int);
			ultraDataColumn74.DataType = typeof(System.Decimal);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn72,
																																 ultraDataColumn73,
																																 ultraDataColumn74});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource8;
			appearance151.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance151;
			ultraGridColumn80.Header.VisiblePosition = 0;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.Caption = "Concepto";
			ultraGridColumn81.Header.VisiblePosition = 1;
			ultraGridColumn81.Width = 135;
			appearance152.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn82.CellAppearance = appearance152;
			ultraGridColumn82.Format = "#,##0.00";
			ultraGridColumn82.Header.Caption = "Total";
			ultraGridColumn82.Header.VisiblePosition = 2;
			ultraGridColumn82.Width = 98;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn80,
																											ultraGridColumn81,
																											ultraGridColumn82});
			ultraGridBand10.Header.Caption = "Utilidad Por Año";
			ultraGridBand10.HeaderVisible = true;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance153.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance153;
			appearance154.ForeColor = System.Drawing.Color.Black;
			appearance154.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance154;
			appearance155.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance155.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance155.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance155.FontData.BoldAsString = "True";
			appearance155.FontData.Name = "Arial";
			appearance155.FontData.SizeInPoints = 8F;
			appearance155.ForeColor = System.Drawing.Color.White;
			appearance155.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance155;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance156.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance156.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance156.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance156;
			appearance157.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance157.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance157.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance157;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(0, 432);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(256, 152);
			this.ultraGrid2.TabIndex = 816;
			// 
			// frmReportePorcentajesMargenUtilidadLocal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1344, 590);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.ugridIYD);
			this.Controls.Add(this.ugridPorcentajes);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.ugridTotales);
			this.Controls.Add(this.button4);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.uGridporcentaje);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.uGridInformacion2);
			this.Controls.Add(this.chkTodos2);
			this.Controls.Add(this.uGridLocales2);
			this.Controls.Add(this.lblContador2);
			this.Controls.Add(this.label30);
			this.Controls.Add(this.label29);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label27);
			this.Controls.Add(this.label28);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.txtMargenCrediPOINT2);
			this.Controls.Add(this.txtMargenTarjeta2);
			this.Controls.Add(this.txtMargenEfectivo2);
			this.Controls.Add(this.txtMargen2);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.txtPorcUtCrediPoint2);
			this.Controls.Add(this.txtPorcUtEfectivo2);
			this.Controls.Add(this.txtPorcUtTarjeta2);
			this.Controls.Add(this.txtPorcCrediPoint2);
			this.Controls.Add(this.txtPorcEfectivo2);
			this.Controls.Add(this.txtPorcTarjeta2);
			this.Controls.Add(this.txtBusqueda2);
			this.Controls.Add(this.btnExcel2);
			this.Controls.Add(this.btnVer2);
			this.Controls.Add(this.dtDesde2);
			this.Controls.Add(this.dtHasta2);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.uGridLocales);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtMargenCrediPOINT);
			this.Controls.Add(this.txtMargenTarjeta);
			this.Controls.Add(this.txtMargenEfectivo);
			this.Controls.Add(this.txtMargen);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtPorcUtCrediPoint);
			this.Controls.Add(this.txtPorcUtEfectivo);
			this.Controls.Add(this.txtPorcUtTarjeta);
			this.Controls.Add(this.txtPorcCrediPoint);
			this.Controls.Add(this.txtPorcEfectivo);
			this.Controls.Add(this.txtPorcTarjeta);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.lblContador);
			this.KeyPreview = true;
			this.Name = "frmReportePorcentajesMargenUtilidadLocal";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Porcentajes Margen De Utilidad Por Locales";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportePorcentajesMargenUtilidadLocal_KeyDown);
			this.Load += new System.EventHandler(this.frmReportePorcentajesMargenUtilidadLocal_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtMargenCrediPOINT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtCrediPoint)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcCrediPoint)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenCrediPOINT2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargen2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtCrediPoint2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtEfectivo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcUtTarjeta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcCrediPoint2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcEfectivo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcTarjeta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridporcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridTotales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentajes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridIYD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		private Acceso miAcceso;
		private void frmReportePorcentajesMargenUtilidadLocal_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200521");
			if (miAcceso.BExportar)
			{
				this.btnExcel.Visible =  true;
				this.btnExcel2.Visible =  true;
				this.button4.Visible =  true;

			}

			if(miAcceso.Buscar)
			{
				#region Habilita Periodo A
					 this.label8.Visible =  true;
				   this.label1.Visible =  true;
           this.txtPorcEfectivo.Visible =  true;
           this.label2.Visible =  true;
           this. txtPorcTarjeta.Visible =  true;
           this.label3.Visible =  true;
           this.txtPorcCrediPoint.Visible =  true;
           this. label4.Visible =  true;
           this.txtPorcUtEfectivo.Visible =  true;
           this.label5.Visible =  true;
           this.txtPorcUtTarjeta.Visible =  true;
           this.label6.Visible =  true;
           this.txtPorcUtCrediPoint.Visible =  true;
           this.label10.Visible =  true;
           this.txtMargen.Visible =  true;
           this.label11.Visible =  true;
           this.txtMargenEfectivo.Visible =  true;
           this.label12.Visible =  true;
           this.txtMargenTarjeta.Visible =  true;
           this.label13.Visible =  true;
           this.txtMargenCrediPOINT.Visible =  true;
           this.label9.Visible =  true;
				#endregion Habilita Periodo A
			}
			if(miAcceso.Imprimir)
			{
				#region Habilita Periodo B
					this.label18.Visible =  true;
				  this.label18.Visible =  true;
          this.label17.Visible =  true;
          this.label24.Visible =  true;
          this.label23.Visible =  true;
          this.label22.Visible =  true;
          this.label21.Visible =  true;
          this.label20.Visible =  true;
          this.label19.Visible =  true;
          this.label28.Visible =  true;
          this.label16.Visible =  true;
          this.label15.Visible =  true;
          this.label14.Visible =  true;
          this.txtMargenCrediPOINT2.Visible =  true;
          this.txtMargenTarjeta2.Visible =  true;
          this.txtMargenEfectivo2.Visible =  true;
          this.txtMargen2.Visible =  true;
          this.txtPorcUtCrediPoint2.Visible =  true;
          this.txtPorcUtTarjeta2.Visible =  true;
          this.txtPorcUtEfectivo2.Visible =  true;
          this.txtPorcCrediPoint2.Visible =  true;
          this.txtPorcTarjeta2.Visible =  true;
          this.txtPorcEfectivo2.Visible =  true;
				#endregion Habilita Periodo B
			}

			#region Parte uno
      DateTime dtInicio = DateTime.Today.AddYears(-1);
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;	
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;

			this.dtDesde2.Value = "01/" + DateTime.Now.Month.ToString() + "/" + dtInicio.Year.ToString();	
			this.dtHasta2.Value = DateTime.Today.AddYears(-1);

			string sSQL = string.Format("Exec ValidaLocalesRangoMorosidad {0} ", MenuLatinium.IdUsuario);		
			
			this.uGridLocales.DataSource = FuncionesProcedimientos.dtGeneral(sSQL);

//			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Convert(Bit, 1) As Sel From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");
	   
			#endregion Parte uno

			#region Parte dos 

			this.dtDesde2.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtHasta2.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			this.dtDesde2.CalendarInfo.MaxDate = DateTime.Today;			
			this.dtHasta2.CalendarInfo.MaxDate = DateTime.Today;

//			this.dtDesde2.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
//			this.dtHasta2.Value = DateTime.Today;

			string sSQLa = string.Format("Exec ValidaLocalesRangoMorosidad {0} ", MenuLatinium.IdUsuario);		
			
			this.uGridLocales2.DataSource = FuncionesProcedimientos.dtGeneral(sSQLa);

//			this.uGridLocales2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Nombre, Convert(Bit, 1) As Sel From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");
	  
			#endregion Parte dos 
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (this.dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde.Value) > Convert.ToDateTime(this.dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}
			#endregion Valida Fechas
						
			string sLocales = "";			
			this.lblContador.Text = "";
						
			#region Locales
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales.Length == 0) sLocales = dr.Cells["Bodega"].Value.ToString(); else if (sLocales.Length > 0) sLocales = sLocales + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Locales
			
			#region Consulta
			this.uGridInformacion.DataSource = ultraDataSource2;

			this.Cursor = Cursors.WaitCursor;
		
			string sSQL = string.Format("Exec ResumenPorcentajeUtilidadPorPeriodo '{0}', '{1}', {2}, '{3}', '{4}'", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 0, this.txtBusqueda.Text.ToString(), sLocales);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);
			
			this.lblContador.Text = this.uGridInformacion.Rows.Count + " Registros Encontrados";

			this.Cursor = Cursors.Default;
			#endregion Consulta

			#region Porcentajes Totales
			this.txtPorcEfectivo.Value = 0.00m;
			this.txtPorcTarjeta.Value = 0.00m;
			this.txtPorcCrediPoint.Value = 0.00m;
			this.txtPorcUtEfectivo.Value = 0.00m;
			this.txtPorcUtTarjeta.Value = 0.00m;
			this.txtPorcUtCrediPoint.Value = 0.00m;

			if (this.uGridInformacion.Rows.Count > 0)
			{
				decimal dSumaTotal = 0.00m;
				decimal dSumaEfectivo = 0.00m;
				decimal dSumaTarjeta = 0.00m;
				decimal dSumaCrediPoint = 0.00m;
				decimal dSumaTotalUtilidad = 0.00m;
				decimal dSumaUtEfectivo = 0.00m;
				decimal dSumaUtTarjeta = 0.00m;
				decimal dSumaUtCrediPoint = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{
					dSumaTotal = dSumaTotal + Convert.ToDecimal(dr.Cells["Total"].Value);
					dSumaEfectivo = dSumaEfectivo + Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value);
					dSumaTarjeta = dSumaTarjeta + Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value);
					dSumaCrediPoint = dSumaCrediPoint + Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value);

					dSumaTotalUtilidad = dSumaTotalUtilidad + Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value);
					dSumaUtEfectivo = dSumaUtEfectivo + Convert.ToDecimal(dr.Cells["UtilidadEfectivo"].Value);
					dSumaUtTarjeta = dSumaUtTarjeta + Convert.ToDecimal(dr.Cells["UtilidadTarjeta"].Value);
					dSumaUtCrediPoint = dSumaUtCrediPoint + Convert.ToDecimal(dr.Cells["UtilidadCrediPoint"].Value);
				}			

				this.txtPorcEfectivo.Value = dSumaEfectivo/dSumaTotal * 100.00m;
				this.txtPorcTarjeta.Value = dSumaTarjeta/dSumaTotal * 100.00m;
				this.txtPorcCrediPoint.Value = dSumaCrediPoint/dSumaTotal * 100.00m;

				this.txtPorcUtEfectivo.Value = dSumaUtEfectivo/dSumaTotalUtilidad * 100.00m;
				this.txtPorcUtTarjeta.Value = dSumaUtTarjeta/dSumaTotalUtilidad * 100.00m;
				this.txtPorcUtCrediPoint.Value = dSumaUtCrediPoint/dSumaTotalUtilidad * 100.00m;

				this.txtMargen.Value = dSumaTotalUtilidad/dSumaTotal * 100.00m;

				this.txtMargenEfectivo.Value = dSumaUtEfectivo/dSumaEfectivo * 100.00m;
				this.txtMargenTarjeta.Value = dSumaUtTarjeta/dSumaTarjeta * 100.00m;
				this.txtMargenCrediPOINT.Value = dSumaUtCrediPoint/dSumaCrediPoint * 100.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{					
					dr.Cells["PorcentajeVentasLocal"].Value = (Convert.ToDecimal(dr.Cells["Total"].Value)/dSumaTotal) * 100;

					dr.Cells["PorcentajeUtilidadLocal"].Value = (Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value)/dSumaTotalUtilidad) * 100;

					dr.Cells["Margen"].Value = (Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value)/Convert.ToDecimal(dr.Cells["Total"].Value)) * 100;

					if (Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value) > 0.00m)
						dr.Cells["MargenEfectivo"].Value = (Convert.ToDecimal(dr.Cells["UtilidadEfectivo"].Value)/Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value)) * 100;//VentaEfectivo
					if (Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value) > 0.00m)
						dr.Cells["MargenTarjeta"].Value = (Convert.ToDecimal(dr.Cells["UtilidadTarjeta"].Value)/Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value)) * 100;//VentaTarjeta
					if (Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value) > 0.00m)
						dr.Cells["MargenCrediPoint"].Value = (Convert.ToDecimal(dr.Cells["UtilidadCrediPoint"].Value)/Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value)) * 100;//VentaCrediPoint
				}
			}
			#endregion Porcentajes Totales
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion2);
		}

		private void btnVer2_Click(object sender, System.EventArgs e)
		{
			#region Valida Fechas
			if (this.dtDesde2.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde2.Focus();
				return;
			}

			if (this.dtHasta2.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta2.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtDesde2.Value) > Convert.ToDateTime(this.dtHasta2.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta2.Focus();
				return;
			}
			#endregion Valida Fechas
						
			string sLocales2 = "";			
			this.lblContador2.Text = "";
						
			#region Locales
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales2.Rows.All)
			{
				if ((bool) dr.Cells["Sel"].Value)
				{
					if (sLocales2.Length == 0) sLocales2 = dr.Cells["Bodega"].Value.ToString(); else if (sLocales2.Length > 0) sLocales2 = sLocales2 + ", " + dr.Cells["Bodega"].Value.ToString();
				}
			}

			if (sLocales2.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Locales
			
			#region Consulta
			this.uGridInformacion2.DataSource = ultraDataSource2;

			this.Cursor = Cursors.WaitCursor;
		
			string sSQL2 = string.Format("Exec ResumenPorcentajeUtilidadPorPeriodo '{0}', '{1}', {2}, '{3}', '{4}'", 
				Convert.ToDateTime(this.dtDesde2.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta2.Value).ToString("yyyyMMdd"), 0, this.txtBusqueda2.Text.ToString(), sLocales2);
			FuncionesProcedimientos.dsGeneral(sSQL2, this.uGridInformacion2);
			
			this.lblContador2.Text = this.uGridInformacion2.Rows.Count + " Registros Encontrados";

			this.Cursor = Cursors.Default;
			#endregion Consulta

			#region Porcentajes Totales
			this.txtPorcEfectivo2.Value = 0.00m;
			this.txtPorcTarjeta2.Value = 0.00m;
			this.txtPorcCrediPoint2.Value = 0.00m;
			this.txtPorcUtEfectivo2.Value = 0.00m;
			this.txtPorcUtTarjeta2.Value = 0.00m;
			this.txtPorcUtCrediPoint2.Value = 0.00m;

			if (this.uGridInformacion2.Rows.Count > 0)
			{
				decimal dSumaTotal2 = 0.00m;
				decimal dSumaEfectivo2 = 0.00m;
				decimal dSumaTarjeta2 = 0.00m;
				decimal dSumaCrediPoint2 = 0.00m;
				decimal dSumaTotalUtilidad2 = 0.00m;
				decimal dSumaUtEfectivo2 = 0.00m;
				decimal dSumaUtTarjeta2 = 0.00m;
				decimal dSumaUtCrediPoint2 = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion2.Rows.All)
				{
					dSumaTotal2 = dSumaTotal2 + Convert.ToDecimal(dr.Cells["Total"].Value);
					dSumaEfectivo2 = dSumaEfectivo2 + Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value);
					dSumaTarjeta2 = dSumaTarjeta2 + Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value);
					dSumaCrediPoint2 = dSumaCrediPoint2 + Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value);

					dSumaTotalUtilidad2 = dSumaTotalUtilidad2 + Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value);
					dSumaUtEfectivo2 = dSumaUtEfectivo2 + Convert.ToDecimal(dr.Cells["UtilidadEfectivo"].Value);
					dSumaUtTarjeta2 = dSumaUtTarjeta2 + Convert.ToDecimal(dr.Cells["UtilidadTarjeta"].Value);
					dSumaUtCrediPoint2 = dSumaUtCrediPoint2 + Convert.ToDecimal(dr.Cells["UtilidadCrediPoint"].Value);
				}			

				this.txtPorcEfectivo2.Value = dSumaEfectivo2/dSumaTotal2 * 100.00m;
				this.txtPorcTarjeta2.Value = dSumaTarjeta2/dSumaTotal2 * 100.00m;
				this.txtPorcCrediPoint2.Value = dSumaCrediPoint2/dSumaTotal2 * 100.00m;

				this.txtPorcUtEfectivo2.Value = dSumaUtEfectivo2/dSumaTotalUtilidad2 * 100.00m;
				this.txtPorcUtTarjeta2.Value = dSumaUtTarjeta2/dSumaTotalUtilidad2 * 100.00m;
				this.txtPorcUtCrediPoint2.Value = dSumaUtCrediPoint2/dSumaTotalUtilidad2 * 100.00m;

				this.txtMargen2.Value = dSumaTotalUtilidad2/dSumaTotal2 * 100.00m;

				this.txtMargenEfectivo2.Value = dSumaUtEfectivo2/dSumaEfectivo2 * 100.00m;
				this.txtMargenTarjeta2.Value = dSumaUtTarjeta2/dSumaTarjeta2 * 100.00m;
				this.txtMargenCrediPOINT2.Value = dSumaUtCrediPoint2/dSumaCrediPoint2 * 100.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion2.Rows.All)
				{					
					dr.Cells["PorcentajeVentasLocal"].Value = (Convert.ToDecimal(dr.Cells["Total"].Value)/dSumaTotal2) * 100;

					dr.Cells["PorcentajeUtilidadLocal"].Value = (Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value)/dSumaTotalUtilidad2) * 100;

					dr.Cells["Margen"].Value = (Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value)/Convert.ToDecimal(dr.Cells["Total"].Value)) * 100;

					if (Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value) > 0.00m)
						dr.Cells["MargenEfectivo"].Value = (Convert.ToDecimal(dr.Cells["UtilidadEfectivo"].Value)/Convert.ToDecimal(dr.Cells["VentaEfectivo"].Value)) * 100;//VentaEfectivo
					if (Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value) > 0.00m)
						dr.Cells["MargenTarjeta"].Value = (Convert.ToDecimal(dr.Cells["UtilidadTarjeta"].Value)/Convert.ToDecimal(dr.Cells["VentaTarjeta"].Value)) * 100;//VentaTarjeta
					if (Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value) > 0.00m)
						dr.Cells["MargenCrediPoint"].Value = (Convert.ToDecimal(dr.Cells["UtilidadCrediPoint"].Value)/Convert.ToDecimal(dr.Cells["VentaCrediPoint"].Value)) * 100;//VentaCrediPoint
				}
			}
			#endregion Porcentajes Totales
		}

		private void chkTodos2_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales2.Rows.All)
			{
				if (this.chkTodos2.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			int icont = 0;
			if (Convert.ToDateTime(this.dtDesde.Value) == Convert.ToDateTime(this.dtDesde2.Value) )
//				& Convert.ToDateTime(this.dtHasta.Value) = Convert.ToDateTime(this.dtHasta2.Value ))
			{
				icont = icont +1 ;
			}
			if (Convert.ToDateTime(this.dtHasta.Value) == Convert.ToDateTime(this.dtHasta2.Value) )
				//				& Convert.ToDateTime(this.dtHasta.Value) = Convert.ToDateTime(this.dtHasta2.Value ))
			{
				icont = icont +1 ;
			}
			if (icont >= 2)
			{
				MessageBox.Show("Las Fechas del Periodo A y B Son las mismas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			
			}
//		  this.label31.Text = "";
			this.btnVer_Click  (sender,e);
			this.btnVer2_Click  (sender,e);
	
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size=75000;"))
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


					
					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
					
						string sSQLDetalless = string.Format("Exec GuardaReportePorcMargenUtilidad {0}, {1}, {2}, '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}",
							0,
							1, 
							0, 
							dr.Cells["Local"].Value.ToString(), 
							Convert.ToDecimal(dr.Cells["Total"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeVentasLocal"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeEfectivo"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeTarjeta"].Value),
							Convert.ToDecimal(dr.Cells["PorcentajeCrediPoint"].Value), 
							Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeUtilidadLocal"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeEfectivoUtilidad"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeTarjetaUtilidad"].Value),
							Convert.ToDecimal(dr.Cells["PorcentajeCrediPointUtilidad"].Value),
							Convert.ToDecimal(dr.Cells["Margen"].Value), 
							Convert.ToDecimal(dr.Cells["MargenEfectivo"].Value), 
							Convert.ToDecimal(dr.Cells["MargenTarjeta"].Value), 
							Convert.ToDecimal(dr.Cells["MargenCrediPoint"].Value)
							
							);
						oCmdActualiza.CommandText = sSQLDetalless;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Detalle	
          
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion2.Rows.All)
					{
					
						string sSQLDetalles = string.Format("Exec GuardaReportePorcMargenUtilidad2 {0}, {1}, {2}, '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}",
							0,
							2, 
							0, 
							dr.Cells["Local"].Value.ToString(), 
							Convert.ToDecimal(dr.Cells["Total"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeVentasLocal"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeEfectivo"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeTarjeta"].Value),
							Convert.ToDecimal(dr.Cells["PorcentajeCrediPoint"].Value), 
							Convert.ToDecimal(dr.Cells["TotalUtilidad"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeUtilidadLocal"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeEfectivoUtilidad"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeTarjetaUtilidad"].Value), 
							Convert.ToDecimal(dr.Cells["PorcentajeCrediPointUtilidad"].Value),
							Convert.ToDecimal(dr.Cells["Margen"].Value), 
							Convert.ToDecimal(dr.Cells["MargenEfectivo"].Value), 
							Convert.ToDecimal(dr.Cells["MargenTarjeta"].Value), 
							Convert.ToDecimal(dr.Cells["MargenCrediPoint"].Value)
							);
						oCmdActualiza.CommandText = sSQLDetalles;
						oCmdActualiza.ExecuteNonQuery();
					}


					oTransaction.Commit();

				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}
      string sLocales3 = "";		
			#region Locales
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion2.Rows.All)
			{
				
					if (sLocales3.Length == 0) sLocales3 = dr.Cells["Bodega"].Value.ToString(); else if (sLocales3.Length > 0) sLocales3 = sLocales3 + ", " + dr.Cells["Bodega"].Value.ToString();
			}

			if (sLocales3.Length == 0)
			{
				MessageBox.Show("Seleccione al menos un Local", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Locales

			this.ugridTotales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaReportePorcTotalMargenUtilidad '{0}', '{1}', {2}, '{3}', '{4}',1,1", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 0, 
				this.txtBusqueda2.Text.ToString(), sLocales3));

			this.ugridPorcentajes.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaReportePorcTotalMargenUtilidad '{0}', '{1}', {2}, '{3}', '{4}',1,2", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 0, 
				this.txtBusqueda2.Text.ToString(), sLocales3));

			this.ugridIYD.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaReportePorcTotalMargenUtilidad '{0}', '{1}', {2}, '{3}', '{4}',1,3", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 0, 
				this.txtBusqueda2.Text.ToString(), sLocales3));

			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaReportePorcTotalMargenUtilidad '{0}', '{1}', {2}, '{3}', '{4}',2,0,3", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 0, 
				this.txtBusqueda2.Text.ToString(), sLocales3));

			

//			this.label31.Text = this.uGridporcentaje.Rows.Count + " Registros Encontrados";
      
      this.uGridporcentaje.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaReportePorcMargenUtilidad2 {0}", 0));
//		  this.label31.Text = this.uGridporcentaje.Rows.Count + " Registros Encontrados";

		}

		private void button4_Click(object sender, System.EventArgs e)
		{
					FuncionesProcedimientos.ExportaExcel(this.uGridporcentaje);
		}

		private void frmReportePorcentajesMargenUtilidadLocal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
