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
	/// Descripción breve de ReporteDATANuovo.
	/// </summary>
	public class ReporteDATANuovo : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodos;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbColatreal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDYD;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optGestion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReporteDATANuovo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ReporteDATANuovo));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Colateral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Colateral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotaEntrada");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Gracias");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoNuovo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Edad");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Score");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Colateral");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario_Aprueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_1");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_2");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_3");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_4");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_5");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_6");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_7");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_8");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_9");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_10");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_11");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_12");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_13");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_14");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_15");
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_16");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_17");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_18");
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuovo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codNuovo");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotaEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Gracias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodigoNuovo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Edad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Colateral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario_Aprueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nuovo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("codNuovo");
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Num");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_1");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_2");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_3");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_4");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_5");
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_6");
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_7");
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_8");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_9");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_10");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_11");
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_12");
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_13");
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_14");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_15");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_16");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_17");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora_Cuota_18");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Num");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora_Cuota_18");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.chkTodos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbColatreal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.uGridDYD = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optGestion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColatreal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDYD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).BeginInit();
			this.SuspendLayout();
			// 
			// chkTodos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			appearance1.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkTodos.Appearance = appearance1;
			this.chkTodos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkTodos.Enabled = false;
			this.chkTodos.Location = new System.Drawing.Point(232, 168);
			this.chkTodos.Name = "chkTodos";
			this.chkTodos.Size = new System.Drawing.Size(13, 13);
			this.chkTodos.TabIndex = 689;
			this.chkTodos.CheckedChanged += new System.EventHandler(this.chkTodos_CheckedChanged);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(176, 16);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 685;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(8, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 684;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridLocales
			// 
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance2;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 100;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 199;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn3.Width = 20;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			ultraGridBand1.Header.Caption = "LOCALES";
			ultraGridBand1.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance3.ForeColor = System.Drawing.Color.Black;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLocales.DisplayLayout.Override.ActiveRowAppearance = appearance3;
			this.uGridLocales.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridLocales.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridLocales.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridLocales.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLocales.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLocales.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 176);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(240, 328);
			this.uGridLocales.TabIndex = 686;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn3.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// dtHasta
			// 
			appearance9.FontData.Name = "Tahoma";
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance9;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 14);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 683;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance10.FontData.Name = "Tahoma";
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance10;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 14);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 682;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 50);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(53, 16);
			this.label1.TabIndex = 695;
			this.label1.Text = "Vendedor";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 82);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(50, 16);
			this.label2.TabIndex = 696;
			this.label2.Text = "Colateral";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 114);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(89, 16);
			this.label3.TabIndex = 698;
			this.label3.Text = "Usuario Aprueba";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbUsuario
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbUsuario.Appearance = appearance11;
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 198;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbUsuario.DisplayMember = "Nombre";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(136, 112);
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(200, 21);
			this.cmbUsuario.TabIndex = 697;
			this.cmbUsuario.ValueMember = "Usuario";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5});
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
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// cmbVendedor
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance12;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.ultraDataSource3;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 239;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(96, 48);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(240, 21);
			this.cmbVendedor.TabIndex = 699;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn8.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn8,
																																 ultraDataColumn9});
			// 
			// cmbColatreal
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbColatreal.Appearance = appearance13;
			this.cmbColatreal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbColatreal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColatreal.DataSource = this.ultraDataSource4;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 235;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbColatreal.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbColatreal.DisplayMember = "Nombre";
			this.cmbColatreal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbColatreal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColatreal.Location = new System.Drawing.Point(96, 80);
			this.cmbColatreal.Name = "cmbColatreal";
			this.cmbColatreal.Size = new System.Drawing.Size(240, 21);
			this.cmbColatreal.TabIndex = 700;
			this.cmbColatreal.ValueMember = "Colateral";
			// 
			// uGridDYD
			// 
			this.uGridDYD.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDYD.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDYD.DataSource = this.ultraDataSource5;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridDYD.DisplayLayout.Appearance = appearance14;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 29;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Width = 150;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.ForeColor = System.Drawing.Color.Black;
			ultraGridColumn13.CellAppearance = appearance15;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "# Factura";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance16;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 5;
			ultraGridColumn15.Width = 90;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance17;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Cuota Entrada";
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance18;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Valor Cuotas";
			ultraGridColumn17.Header.VisiblePosition = 7;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn18.CellAppearance = appearance19;
			ultraGridColumn18.Header.Caption = "Dias Gracias";
			ultraGridColumn18.Header.VisiblePosition = 8;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Codigo Nuovo";
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Width = 90;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "Cliente";
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Width = 150;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.Caption = "Cedula";
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn22.Width = 90;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance20;
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn23.Width = 90;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance21;
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn24.Width = 90;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn25.CellAppearance = appearance22;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Width = 90;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.Caption = "Usuario Aprueba";
			ultraGridColumn26.Header.VisiblePosition = 16;
			ultraGridColumn26.Width = 90;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellAppearance = appearance23;
			ultraGridColumn27.Header.Caption = "Cuota 1";
			ultraGridColumn27.Header.VisiblePosition = 19;
			ultraGridColumn27.Width = 90;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn28.CellAppearance = appearance24;
			ultraGridColumn28.Header.Caption = "Cuota 2";
			ultraGridColumn28.Header.VisiblePosition = 20;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn29.CellAppearance = appearance25;
			ultraGridColumn29.Header.Caption = "Cuota 3";
			ultraGridColumn29.Header.VisiblePosition = 21;
			ultraGridColumn29.Width = 90;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn30.CellAppearance = appearance26;
			ultraGridColumn30.Header.Caption = "Cuota 4";
			ultraGridColumn30.Header.VisiblePosition = 22;
			ultraGridColumn30.Width = 90;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn31.CellAppearance = appearance27;
			ultraGridColumn31.Header.Caption = "Cuota 5";
			ultraGridColumn31.Header.VisiblePosition = 23;
			ultraGridColumn31.Width = 90;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn32.CellAppearance = appearance28;
			ultraGridColumn32.Header.Caption = "Cuota 6";
			ultraGridColumn32.Header.VisiblePosition = 24;
			ultraGridColumn32.Width = 90;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn33.CellAppearance = appearance29;
			ultraGridColumn33.Header.Caption = "Cuota 7";
			ultraGridColumn33.Header.VisiblePosition = 25;
			ultraGridColumn33.Width = 90;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance30;
			ultraGridColumn34.Header.Caption = "Cuota 8";
			ultraGridColumn34.Header.VisiblePosition = 26;
			ultraGridColumn34.Width = 90;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn35.CellAppearance = appearance31;
			ultraGridColumn35.Header.Caption = "Cuota 9";
			ultraGridColumn35.Header.VisiblePosition = 27;
			ultraGridColumn35.Width = 90;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance32;
			ultraGridColumn36.Header.Caption = "Cuota 10";
			ultraGridColumn36.Header.VisiblePosition = 28;
			ultraGridColumn36.Width = 90;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn37.CellAppearance = appearance33;
			ultraGridColumn37.Header.Caption = "Cuota 11";
			ultraGridColumn37.Header.VisiblePosition = 29;
			ultraGridColumn37.Width = 90;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn38.CellAppearance = appearance34;
			ultraGridColumn38.Header.Caption = "Cuota 12";
			ultraGridColumn38.Header.VisiblePosition = 30;
			ultraGridColumn38.Width = 90;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn39.CellAppearance = appearance35;
			ultraGridColumn39.Header.Caption = "Cuota 13";
			ultraGridColumn39.Header.VisiblePosition = 31;
			ultraGridColumn39.Width = 90;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn40.CellAppearance = appearance36;
			ultraGridColumn40.Header.Caption = "Cuota 14";
			ultraGridColumn40.Header.VisiblePosition = 32;
			ultraGridColumn40.Width = 90;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn41.CellAppearance = appearance37;
			ultraGridColumn41.Header.Caption = "Cuota 15";
			ultraGridColumn41.Header.VisiblePosition = 33;
			ultraGridColumn41.Width = 90;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn42.CellAppearance = appearance38;
			ultraGridColumn42.Header.Caption = "Cuota 16";
			ultraGridColumn42.Header.VisiblePosition = 34;
			ultraGridColumn42.Width = 90;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance39.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn43.CellAppearance = appearance39;
			ultraGridColumn43.Header.Caption = "Cuota 17";
			ultraGridColumn43.Header.VisiblePosition = 35;
			ultraGridColumn43.Width = 90;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance40.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn44.CellAppearance = appearance40;
			ultraGridColumn44.Header.Caption = "Cuota 18";
			ultraGridColumn44.Header.VisiblePosition = 36;
			ultraGridColumn44.Width = 90;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 18;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn46.Header.Caption = "Codigo";
			ultraGridColumn46.Header.VisiblePosition = 17;
			ultraGridBand5.Columns.AddRange(new object[] {
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
																										 ultraGridColumn46});
			ultraGridBand5.SummaryFooterCaption = "TOTAL GENERAL";
			this.uGridDYD.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			appearance41.ForeColor = System.Drawing.Color.Black;
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.ActiveRowAppearance = appearance41;
			this.uGridDYD.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDYD.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDYD.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance42.BackColor = System.Drawing.Color.Transparent;
			this.uGridDYD.DisplayLayout.Override.CardAreaAppearance = appearance42;
			appearance43.ForeColor = System.Drawing.Color.Black;
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.CellAppearance = appearance43;
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDYD.DisplayLayout.Override.GroupByRowAppearance = appearance44;
			appearance45.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance45.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance45.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance45.FontData.BoldAsString = "True";
			appearance45.FontData.Name = "Arial";
			appearance45.FontData.SizeInPoints = 8.5F;
			appearance45.ForeColor = System.Drawing.Color.White;
			appearance45.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDYD.DisplayLayout.Override.HeaderAppearance = appearance45;
			this.uGridDYD.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance46.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance46.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowAlternateAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.RowSelectorAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDYD.DisplayLayout.Override.SelectedRowAppearance = appearance48;
			this.uGridDYD.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDYD.Location = new System.Drawing.Point(264, 176);
			this.uGridDYD.Name = "uGridDYD";
			this.uGridDYD.Size = new System.Drawing.Size(944, 328);
			this.uGridDYD.TabIndex = 701;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn42.DataType = typeof(int);
			ultraDataColumn43.DataType = typeof(int);
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(bool);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn46});
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(264, 144);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 703;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click_1);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(448, 112);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(75, 24);
			this.btnExportar.TabIndex = 702;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(352, 48);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 704;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(448, 144);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 987;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource6;
			appearance49.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance49;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 220;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn49.CellAppearance = appearance50;
			ultraGridColumn49.Format = "#,##0.00";
			ultraGridColumn49.Header.Caption = "Cuota 1";
			ultraGridColumn49.Header.VisiblePosition = 2;
			ultraGridColumn49.Width = 90;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance51.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn50.CellAppearance = appearance51;
			ultraGridColumn50.Format = "#,##0.00";
			ultraGridColumn50.Header.Caption = "Cuota 2";
			ultraGridColumn50.Header.VisiblePosition = 3;
			ultraGridColumn50.Width = 90;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn51.CellAppearance = appearance52;
			ultraGridColumn51.Format = "#,##0.00";
			ultraGridColumn51.Header.Caption = "Cuota 3";
			ultraGridColumn51.Header.VisiblePosition = 4;
			ultraGridColumn51.Width = 90;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance53.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn52.CellAppearance = appearance53;
			ultraGridColumn52.Format = "#,##0.00";
			ultraGridColumn52.Header.Caption = "Cuota 4";
			ultraGridColumn52.Header.VisiblePosition = 5;
			ultraGridColumn52.Width = 90;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn53.CellAppearance = appearance54;
			ultraGridColumn53.Format = "#,##0.00";
			ultraGridColumn53.Header.Caption = "Cuota 5";
			ultraGridColumn53.Header.VisiblePosition = 6;
			ultraGridColumn53.Width = 90;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn54.CellAppearance = appearance55;
			ultraGridColumn54.Format = "#,##0.00";
			ultraGridColumn54.Header.Caption = "Cuota 6";
			ultraGridColumn54.Header.VisiblePosition = 7;
			ultraGridColumn54.Width = 90;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance56.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn55.CellAppearance = appearance56;
			ultraGridColumn55.Format = "#,##0.00";
			ultraGridColumn55.Header.Caption = "Cuota 7";
			ultraGridColumn55.Header.VisiblePosition = 8;
			ultraGridColumn55.Width = 90;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance57.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn56.CellAppearance = appearance57;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.Caption = "Cuota 8";
			ultraGridColumn56.Header.VisiblePosition = 9;
			ultraGridColumn56.Width = 90;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn57.CellAppearance = appearance58;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.Caption = "Cuota 9";
			ultraGridColumn57.Header.VisiblePosition = 10;
			ultraGridColumn57.Width = 90;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance59.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn58.CellAppearance = appearance59;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.Caption = "Cuota 10";
			ultraGridColumn58.Header.VisiblePosition = 11;
			ultraGridColumn58.Width = 90;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance60.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn59.CellAppearance = appearance60;
			ultraGridColumn59.Format = "#,##0.00";
			ultraGridColumn59.Header.Caption = "Cuota 11";
			ultraGridColumn59.Header.VisiblePosition = 12;
			ultraGridColumn59.Width = 90;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance61.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn60.CellAppearance = appearance61;
			ultraGridColumn60.Format = "#,##0.00";
			ultraGridColumn60.Header.Caption = "Cuota 12";
			ultraGridColumn60.Header.VisiblePosition = 13;
			ultraGridColumn60.Width = 90;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance62.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn61.CellAppearance = appearance62;
			ultraGridColumn61.Format = "#,##0.00";
			ultraGridColumn61.Header.Caption = "Cuota 13";
			ultraGridColumn61.Header.VisiblePosition = 14;
			ultraGridColumn61.Width = 90;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance63.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn62.CellAppearance = appearance63;
			ultraGridColumn62.Format = "#,##0.00";
			ultraGridColumn62.Header.Caption = "Cuota 14";
			ultraGridColumn62.Header.VisiblePosition = 15;
			ultraGridColumn62.Width = 90;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn63.CellAppearance = appearance64;
			ultraGridColumn63.Format = "#,##0.00";
			ultraGridColumn63.Header.Caption = "Cuota 15";
			ultraGridColumn63.Header.VisiblePosition = 16;
			ultraGridColumn63.Width = 90;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn64.CellAppearance = appearance65;
			ultraGridColumn64.Format = "#,##0.00";
			ultraGridColumn64.Header.Caption = "Cuota 16";
			ultraGridColumn64.Header.VisiblePosition = 17;
			ultraGridColumn64.Width = 90;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn65.CellAppearance = appearance66;
			ultraGridColumn65.Format = "#,##0.00";
			ultraGridColumn65.Header.Caption = "Cuota 17";
			ultraGridColumn65.Header.VisiblePosition = 18;
			ultraGridColumn65.Width = 90;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn66.CellAppearance = appearance67;
			ultraGridColumn66.Format = "#,##0.00";
			ultraGridColumn66.Header.Caption = "Cuota 18";
			ultraGridColumn66.Header.VisiblePosition = 19;
			ultraGridColumn66.Width = 90;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn66});
			ultraGridBand6.SummaryFooterCaption = "TOTAL GENERAL";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			appearance68.ForeColor = System.Drawing.Color.Black;
			appearance68.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance68;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance69.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance69;
			appearance70.ForeColor = System.Drawing.Color.Black;
			appearance70.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance70;
			appearance71.ForeColor = System.Drawing.Color.Black;
			appearance71.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.GroupByRowAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance72.FontData.BoldAsString = "True";
			appearance72.FontData.Name = "Arial";
			appearance72.FontData.SizeInPoints = 8.5F;
			appearance72.ForeColor = System.Drawing.Color.White;
			appearance72.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance72;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance73.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance73.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance73;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance74.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance75;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(528, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(672, 160);
			this.ultraGrid1.TabIndex = 988;
			// 
			// ultraDataSource6
			// 
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn50.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(System.Decimal);
			ultraDataColumn57.DataType = typeof(System.Decimal);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			ultraDataColumn59.DataType = typeof(System.Decimal);
			ultraDataColumn60.DataType = typeof(System.Decimal);
			ultraDataColumn61.DataType = typeof(System.Decimal);
			ultraDataColumn62.DataType = typeof(System.Decimal);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn65.DataType = typeof(System.Decimal);
			ultraDataColumn66.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn66});
			// 
			// optGestion
			// 
			appearance76.ForeColorDisabled = System.Drawing.Color.Black;
			this.optGestion.Appearance = appearance76;
			this.optGestion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optGestion.CheckedIndex = 0;
			this.optGestion.ItemAppearance = appearance77;
			this.optGestion.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Con Nuovo";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Sin Nuovo";
			this.optGestion.Items.Add(valueListItem1);
			this.optGestion.Items.Add(valueListItem2);
			this.optGestion.Items.Add(valueListItem3);
			this.optGestion.ItemSpacingVertical = 8;
			this.optGestion.Location = new System.Drawing.Point(8, 144);
			this.optGestion.Name = "optGestion";
			this.optGestion.Size = new System.Drawing.Size(240, 26);
			this.optGestion.TabIndex = 989;
			this.optGestion.Text = "Todos";
			this.optGestion.ValueChanged += new System.EventHandler(this.optGestion_ValueChanged);
			// 
			// ReporteDATANuovo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1216, 510);
			this.Controls.Add(this.optGestion);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.uGridDYD);
			this.Controls.Add(this.cmbColatreal);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.chkTodos);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.uGridLocales);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "ReporteDATANuovo";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Nuovo";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ReporteDATANuovo_KeyDown);
			this.Load += new System.EventHandler(this.ReporteDATANuovo_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColatreal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDYD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optGestion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		private void ReporteDATANuovo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "20053902");

			if (!Funcion.Permiso("1879", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a NUOVO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExportar.Visible = true;

			//miAcceso = new Acceso(cdsSeteoF, "0712");

				this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/08/2022");
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/08/2022");				

			this.dtDesde.CalendarInfo.MaxDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;

			this.dtDesde.Value = DateTime.Today;
			this.dtHasta.Value = DateTime.Today;

			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ResumenDeVentasLocales {0}", MenuLatinium.IdUsuario));

			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Nombre From dbo.Personal Where Vendedor = 1 Order By Nombre");

			this.cmbColatreal.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "	SELECT 'SI'  AS Colateral , 'SI' AS Nombre UNION  SELECT 'NO'  AS Colateral , 'NO' AS Nombre");

			this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "			SELECT Usuario, Usuario as Nombre FROM dbo.Compra C  WHERE EXISTS (SELECT idCompra FROM dbo.DetalleNuovo WHERE idCompra = C.IDCOMPRA) group by Usuario order by Usuario");

	
		 

			if (this.uGridLocales.Rows.Count > 1) 
			{
				this.chkTodos.Enabled = true;
				this.uGridLocales.DisplayLayout.Bands[0].Columns["Sel"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
			}

			//if (miAcceso.BExportar) this.btnExportar.Enabled = true;
	
			this.chkTodos.Checked = true;		
		}

		private void chkTodos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
			{
				if (this.chkTodos.Checked) dr.Cells["Sel"].Value = 1;
				else dr.Cells["Sel"].Value = 0;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
		 this.Consulta ();
		}

		private void Consulta()
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione la Fecha Inicial", false, "", cdsSeteoF)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione la Fecha Final", false, "", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La fecha inicial no puede ser posterior a la fecha final", false)) return;


			string sLocales = "";			
			int idVendedor = 0;
			string scoltarela = "";
			string sUsuario = "";
			string sSQL = "";

			if (this.cmbVendedor.ActiveRow != null) idVendedor = (int)this.cmbVendedor.Value;
			if (this.cmbColatreal.ActiveRow != null) scoltarela = (string)this.cmbColatreal.Value;
			if (this.cmbUsuario.ActiveRow != null) sUsuario = (string)this.cmbUsuario.Value;
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
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
			
				sSQL = string.Format("Exec [ReporteDATANuovo] '{0}', '{1}', '{2}', {3}, '{4}', '{5}',{6}",
					scoltarela ,Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), idVendedor,sUsuario, sLocales, (int) this.optGestion.Value);
				
				
				FuncionesProcedimientos.dsGeneral(sSQL, this.uGridDYD);				

			if ((int) this.uGridDYD.Rows.Count > 0 )
			{
				string sSQLa = "";
				sSQLa = string.Format("Exec [ReporteDATANuovoDatos] '{0}', '{1}', '{2}', {3}, '{4}', '{5}'",
					scoltarela ,Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), idVendedor,sUsuario, sLocales);
				
				
				FuncionesProcedimientos.dsGeneral(sSQLa, this.ultraGrid1);
				this.DiseñoGridColor(this.uGridDYD);
			}

//				this.Text = this.Text + " - " + this.uGridHorizontal.Rows.Count + " Registros Encontrados";
			
		}

		private void btnVer_Click_1(object sender, System.EventArgs e)
		{
			this.Consulta();
				this.lblContador.Text = this.uGridDYD.Rows.Count + " Clientes Encontrados";
      
			
		}

		private void DiseñoGridColor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_1"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_2"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_3"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_4"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_5"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_6"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_7"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_8"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_9"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_10"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_11"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_12"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_13"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_14"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_15"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_16"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_17"].Hidden = true;
			this.uGridDYD.DisplayLayout.Bands[0].Columns["Mora_18"].Hidden = true;

			string sColumna = "";
			int iContColumnas = (int)this.uGridDYD.DisplayLayout.Bands[0].Columns.Count;	

		
					
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDYD.Rows.All)
			{
				for (int i = 17; i < iContColumnas; i++)
				{

					sColumna = this.uGridDYD.DisplayLayout.Bands[0].Columns[i].Header.Caption;			

			
					#region Mora_1
					if( sColumna == "Mora_1")
					{
						if (dr.Cells["Dias_Mora_Cuota_1"].Value.ToString().Trim().Length == 0) 
						{
							dr.Cells["Dias_Mora_Cuota_1"].Appearance.ForeColor = Color.Green;							
						}
						else 	if (dr.Cells["Dias_Mora_Cuota_1"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_1"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_1"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_1"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_1

					#region Mora_2
					if( sColumna == "Mora_2")
					{
						if (dr.Cells["Dias_Mora_Cuota_2"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_2"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_2"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_2"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_2

					#region Mora_3
					if( sColumna == "Mora_3")
					{
						if (dr.Cells["Dias_Mora_Cuota_3"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_3"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_3"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_3"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_3

					#region Mora_4
					if( sColumna == "Mora_4")
					{
						if (dr.Cells["Dias_Mora_Cuota_4"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_4"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_4"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_4"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_4

					#region Mora_5
					if( sColumna == "Mora_5")
					{
						if (dr.Cells["Dias_Mora_Cuota_5"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_5"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_5"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_5"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_5		

					#region Mora_6
					if( sColumna == "Mora_6")
					{
						if (dr.Cells["Dias_Mora_Cuota_6"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_6"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_6"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_6"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_6

					#region Mora_7
					if( sColumna == "Mora_7")
					{
						if (dr.Cells["Dias_Mora_Cuota_7"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_7"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_7"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_7"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_7		

					#region Mora_8
					if( sColumna == "Mora_8")
					{
						if (dr.Cells["Dias_Mora_Cuota_8"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_8"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_8"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_8"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_8		

					#region Mora_9
					if( sColumna == "Mora_9")
					{
						if (dr.Cells["Dias_Mora_Cuota_9"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_9"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_9"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_9"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_9
							
					#region Mora_10
					if( sColumna == "Mora_10")
					{
						if (dr.Cells["Dias_Mora_Cuota_10"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_10"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_10"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_10"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_10
							
					#region Mora_11
					if( sColumna == "Mora_11")
					{
						if (dr.Cells["Dias_Mora_Cuota_11"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_11"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_11"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_11"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_11
							
					#region Mora_12
					if( sColumna == "Mora_12")
					{
						if (dr.Cells["Dias_Mora_Cuota_12"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_12"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_12"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_12"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_12

					#region Mora_13
					if( sColumna == "Mora_13")
					{
						if (dr.Cells["Dias_Mora_Cuota_13"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_13"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_13"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_13"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_13

					#region Mora_14
					if( sColumna == "Mora_14")
					{
						if (dr.Cells["Dias_Mora_Cuota_14"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_14"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_14"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_14"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_14

					#region Mora_15
					if( sColumna == "Mora_15")
					{
						if (dr.Cells["Dias_Mora_Cuota_15"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_15"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_15"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_15"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_15		

					#region Mora_16
					if( sColumna == "Mora_16")
					{
						if (dr.Cells["Dias_Mora_Cuota_16"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_16"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_16"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_16"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_16

					#region Mora_17
					if( sColumna == "Mora_17")
					{
						if (dr.Cells["Dias_Mora_Cuota_17"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_17"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_17"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_17"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_17

					#region Mora_18
					if( sColumna == "Mora_18")
					{
						if (dr.Cells["Dias_Mora_Cuota_18"].Value.ToString().Trim().Length > 0) 
						{
							int dNumero = Convert.ToInt32 ( dr.Cells["Mora_18"].Value.ToString ());

							if(dNumero == 0 )
							{
								dr.Cells["Dias_Mora_Cuota_18"].Appearance.ForeColor = Color.Red;
							}	
							if(dNumero == 1 )
							{
								dr.Cells["Dias_Mora_Cuota_18"].Appearance.ForeColor = Color.Blue;
							}		
						}
					}
					#endregion Mora_18



					//					if (dr.Cells[sColumna].Value.ToString().Trim().Length > 0) 
					//					{
					//						decimal dNumero = Convert.ToDecimal ( dr.Cells[sColumna].Value.ToString ());
					//								
					//						if(dNumero == 0)
					//						{
					//							dr.Cells[sColumna].Appearance.BackColor = Color.White;
					//							dr.Cells[sColumna].Appearance.BackColor2 = Color.White;
					//						}
					//					}
				}
			}

			

		}

		private void ReporteDATANuovo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if ((int) this.uGridDYD.Rows.Count > 0 )
			{
				FuncionesProcedimientos.ExportaExcel(this.uGridDYD);	
				FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);	
			}
		}

		private void optGestion_ValueChanged(object sender, System.EventArgs e)
		{
			this.btnVer_Click_1 (sender , e );
			
		}
	}
}
