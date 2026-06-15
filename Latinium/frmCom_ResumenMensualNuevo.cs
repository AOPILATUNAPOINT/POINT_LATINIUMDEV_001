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
	/// Descripción breve de frmCom_ResumenComisionesNuevo.
	/// </summary>
	public class frmCom_ResumenComisionesNuevo : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVendedores;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCargos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkLocales;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCargos;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optRegiones;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optConsulta;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLocales;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesExportar;
		private System.Windows.Forms.MenuItem mnuAccionesBloquear;
		private System.Windows.Forms.MenuItem mnuAccionesCerrar;
		private System.Windows.Forms.MenuItem mnuAccionesImprimir;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_ResumenComisionesNuevo()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_ResumenComisionesNuevo));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_AsignacionDeVendedores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Promedio");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaDiaria");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaMensual");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCredito");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaGEx");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Por_Cuota_Cero");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCuotaCero");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cartera_Flujo");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCarteraFlujo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTotal");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCredito");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaCarpa");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaGEx");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeDiario");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeMensual");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCredito");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeGEx");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCuotaCeroFinal");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeCarteraFlujoFinal");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionMensual");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionCredito");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BonoSubjefe");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionGEx");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Contador");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasTrabajados");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta_General");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Creditos");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CeroPagos");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcCeroPagos");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella1", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella2", 1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estrella3", 2);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnResumen", 3);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 32, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 32, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_AsignacionDeVendedores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Por_Cuota_Cero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCuotaCero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cartera_Flujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCarteraFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeDiario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCuotaCeroFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCarteraFlujoFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoSubjefe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasTrabajados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta_General");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Creditos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CeroPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcCeroPagos");
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_AsignacionDeVendedores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Promedio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Por_Cuota_Cero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cartera_Flujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaCarpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeDiario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCuotaCeroFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeCarteraFlujoFinal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionMensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BonoSubjefe");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasTrabajados");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta_General");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Creditos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CeroPagos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcCeroPagos");
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnComandos = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.uGridVendedores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkCargos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkLocales = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.uGridCargos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optRegiones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optConsulta = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridLocales = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesExportar = new System.Windows.Forms.MenuItem();
			this.mnuAccionesBloquear = new System.Windows.Forms.MenuItem();
			this.mnuAccionesCerrar = new System.Windows.Forms.MenuItem();
			this.mnuAccionesImprimir = new System.Windows.Forms.MenuItem();
			((System.ComponentModel.ISupportInitialize)(this.uGridVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(976, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 20);
			this.lblEstado.TabIndex = 768;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Enabled = false;
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(1200, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.Size = new System.Drawing.Size(88, 23);
			this.btnComandos.TabIndex = 767;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click_1);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 35);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 17);
			this.label1.TabIndex = 766;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(176, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 17);
			this.label4.TabIndex = 754;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 17);
			this.label5.TabIndex = 753;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// uGridVendedores
			// 
			this.uGridVendedores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVendedores.DataSource = this.ultraDataSource3;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridVendedores.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 3;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridColumn3.Width = 190;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 22;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance2;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance3;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Meta Diaria";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance4;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Meta Mensual";
			ultraGridColumn11.Header.VisiblePosition = 9;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Meta Credito";
			ultraGridColumn12.Header.VisiblePosition = 10;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Meta Gex";
			ultraGridColumn13.Header.VisiblePosition = 11;
			ultraGridColumn13.Width = 85;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "% Cuota Cero";
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 90;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Meta Cuota Cero";
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance9;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Cartera Flujo";
			ultraGridColumn16.Header.VisiblePosition = 14;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Meta Flujo";
			ultraGridColumn17.Header.VisiblePosition = 15;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance11;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Venta Total";
			ultraGridColumn18.Header.VisiblePosition = 16;
			ultraGridColumn18.Width = 80;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance12;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Venta Credito";
			ultraGridColumn19.Header.VisiblePosition = 17;
			ultraGridColumn19.Width = 80;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance13;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Venta Carpa";
			ultraGridColumn20.Header.VisiblePosition = 18;
			ultraGridColumn20.Width = 80;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance14;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Venta Gex";
			ultraGridColumn21.Header.VisiblePosition = 19;
			ultraGridColumn21.Width = 88;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance15;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 20;
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance16;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "% Diario";
			ultraGridColumn23.Header.VisiblePosition = 21;
			ultraGridColumn23.Width = 70;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance17;
			ultraGridColumn24.Format = "#,##0.00";
			ultraGridColumn24.Header.Caption = "% Mensual";
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 70;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn25.CellAppearance = appearance18;
			ultraGridColumn25.Format = "#,##0.00";
			ultraGridColumn25.Header.Caption = "% Credito";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 70;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance19;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.Caption = "% Gex";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 85;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance20;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.Caption = "% Cuota Cero Final";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 90;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance21;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.Caption = "% Cartera Flujo Final";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 90;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn29.CellAppearance = appearance22;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.Caption = "Comision Mensual";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 80;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance23;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Comision Credito";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 80;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn31.CellAppearance = appearance24;
			ultraGridColumn31.Format = "#,##0.00";
			ultraGridColumn31.Header.Caption = "Bono SubJefe";
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn32.CellAppearance = appearance25;
			ultraGridColumn32.Format = "#,##0.00";
			ultraGridColumn32.Header.Caption = "Comision Gex";
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 85;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn33.CellAppearance = appearance26;
			ultraGridColumn33.Format = "#,##0.00";
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 85;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance27;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn35.CellAppearance = appearance28;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance29;
			ultraGridColumn36.Format = "#,##0";
			ultraGridColumn36.Header.Caption = "Dias Trabajados";
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Width = 85;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance30;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.Caption = "Meta General";
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Width = 90;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance31;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Width = 90;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance32;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 70;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn40.CellAppearance = appearance33;
			ultraGridColumn40.Format = "#,##0.00";
			ultraGridColumn40.Header.Caption = "% Cero Pagos";
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Width = 90;
			ultraGridColumn41.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn41.Header.Caption = "";
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn41.Width = 15;
			ultraGridColumn42.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn42.Header.Caption = "";
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn42.Width = 15;
			ultraGridColumn43.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn43.Header.Caption = "";
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn43.Width = 15;
			ultraGridColumn44.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance34.Image = ((object)(resources.GetObject("appearance34.Image")));
			appearance34.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance34.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn44.CellButtonAppearance = appearance34;
			ultraGridColumn44.Header.Caption = "...";
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn44.Width = 30;
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
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance35;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridVendedores.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance36.ForeColor = System.Drawing.Color.Black;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVendedores.DisplayLayout.Override.ActiveRowAppearance = appearance36;
			this.uGridVendedores.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVendedores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVendedores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.uGridVendedores.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance38.FontData.BoldAsString = "True";
			appearance38.FontData.Name = "Arial";
			appearance38.FontData.SizeInPoints = 7.75F;
			appearance38.ForeColor = System.Drawing.Color.White;
			appearance38.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVendedores.DisplayLayout.Override.HeaderAppearance = appearance38;
			this.uGridVendedores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVendedores.DisplayLayout.Override.RowAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVendedores.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			this.uGridVendedores.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVendedores.Location = new System.Drawing.Point(8, 72);
			this.uGridVendedores.Name = "uGridVendedores";
			this.uGridVendedores.Size = new System.Drawing.Size(1280, 488);
			this.uGridVendedores.TabIndex = 756;
			this.uGridVendedores.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridVendedores_InitializeRow_1);
			this.uGridVendedores.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVendedores_ClickCellButton_1);
			this.uGridVendedores.Click += new System.EventHandler(this.uGridVendedores_Click);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
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
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn25.DataType = typeof(System.Decimal);
			ultraDataColumn26.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(System.Decimal);
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn33.DataType = typeof(System.Decimal);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn40});
			// 
			// txtBuscar
			// 
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(56, 32);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(272, 22);
			this.txtBuscar.TabIndex = 765;
			// 
			// chkCargos
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			appearance41.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkCargos.Appearance = appearance41;
			this.chkCargos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkCargos.Checked = true;
			this.chkCargos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCargos.Location = new System.Drawing.Point(168, 80);
			this.chkCargos.Name = "chkCargos";
			this.chkCargos.Size = new System.Drawing.Size(13, 13);
			this.chkCargos.TabIndex = 764;
			this.chkCargos.CheckedChanged += new System.EventHandler(this.chkCargos_CheckedChanged);
			// 
			// chkLocales
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			appearance42.TextHAlign = Infragistics.Win.HAlign.Center;
			this.chkLocales.Appearance = appearance42;
			this.chkLocales.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkLocales.Checked = true;
			this.chkLocales.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkLocales.Enabled = false;
			this.chkLocales.Location = new System.Drawing.Point(168, 248);
			this.chkLocales.Name = "chkLocales";
			this.chkLocales.Size = new System.Drawing.Size(13, 13);
			this.chkLocales.TabIndex = 763;
			this.chkLocales.CheckedChanged += new System.EventHandler(this.chkLocales_CheckedChanged);
			// 
			// uGridCargos
			// 
			this.uGridCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCargos.DataSource = this.ultraDataSource4;
			appearance43.BackColor = System.Drawing.Color.White;
			this.uGridCargos.DisplayLayout.Appearance = appearance43;
			this.uGridCargos.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn45.Header.VisiblePosition = 0;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 73;
			ultraGridColumn46.Header.VisiblePosition = 1;
			ultraGridColumn46.Width = 150;
			ultraGridColumn47.Header.Caption = "...";
			ultraGridColumn47.Header.VisiblePosition = 2;
			ultraGridColumn47.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn47.Width = 20;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47});
			ultraGridBand2.Header.Caption = "CARGOS";
			ultraGridBand2.HeaderVisible = true;
			this.uGridCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance44.ForeColor = System.Drawing.Color.Black;
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCargos.DisplayLayout.Override.ActiveRowAppearance = appearance44;
			this.uGridCargos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCargos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCargos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance45.BackColor = System.Drawing.Color.Transparent;
			this.uGridCargos.DisplayLayout.Override.CardAreaAppearance = appearance45;
			appearance46.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance46.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance46.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance46.FontData.BoldAsString = "True";
			appearance46.FontData.Name = "Arial";
			appearance46.FontData.SizeInPoints = 8.5F;
			appearance46.ForeColor = System.Drawing.Color.White;
			appearance46.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCargos.DisplayLayout.Override.HeaderAppearance = appearance46;
			appearance47.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance47.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.RowAlternateAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargos.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.uGridCargos.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCargos.Location = new System.Drawing.Point(8, 72);
			this.uGridCargos.Name = "uGridCargos";
			this.uGridCargos.Size = new System.Drawing.Size(191, 168);
			this.uGridCargos.TabIndex = 762;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn41.DataType = typeof(int);
			ultraDataColumn41.DefaultValue = 0;
			ultraDataColumn43.DataType = typeof(bool);
			ultraDataColumn43.DefaultValue = false;
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43});
			// 
			// optRegiones
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.optRegiones.Appearance = appearance50;
			this.optRegiones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optRegiones.CheckedIndex = 0;
			this.optRegiones.Enabled = false;
			this.optRegiones.ItemAppearance = appearance51;
			this.optRegiones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Sierra";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Costa";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Oriente";
			this.optRegiones.Items.Add(valueListItem1);
			this.optRegiones.Items.Add(valueListItem2);
			this.optRegiones.Items.Add(valueListItem3);
			this.optRegiones.Items.Add(valueListItem4);
			this.optRegiones.ItemSpacingVertical = 10;
			this.optRegiones.Location = new System.Drawing.Point(448, 31);
			this.optRegiones.Name = "optRegiones";
			this.optRegiones.Size = new System.Drawing.Size(280, 24);
			this.optRegiones.TabIndex = 760;
			this.optRegiones.Text = "Todos";
			this.optRegiones.Visible = false;
			this.optRegiones.ValueChanged += new System.EventHandler(this.optRegiones_ValueChanged);
			// 
			// optConsulta
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.optConsulta.Appearance = appearance52;
			this.optConsulta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optConsulta.CheckedIndex = 0;
			this.optConsulta.Enabled = false;
			this.optConsulta.ItemAppearance = appearance53;
			this.optConsulta.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem5.DataValue = 0;
			valueListItem5.DisplayText = "Resumen Por Vendedor";
			this.optConsulta.Items.Add(valueListItem5);
			this.optConsulta.ItemSpacingVertical = 10;
			this.optConsulta.Location = new System.Drawing.Point(448, 6);
			this.optConsulta.Name = "optConsulta";
			this.optConsulta.Size = new System.Drawing.Size(160, 24);
			this.optConsulta.TabIndex = 759;
			this.optConsulta.Text = "Resumen Por Vendedor";
			this.optConsulta.Visible = false;
			this.optConsulta.Click += new System.EventHandler(this.optConsulta_Click);
			this.optConsulta.ValueChanged += new System.EventHandler(this.optConsulta_ValueChanged_1);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 56);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1296, 8);
			this.groupBox1.TabIndex = 758;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(744, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 757;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(336, 32);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 755;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtHasta
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance54;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 752;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance55;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 751;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridLocales
			// 
			this.uGridLocales.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridLocales.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLocales.DataSource = this.ultraDataSource2;
			appearance56.BackColor = System.Drawing.Color.White;
			this.uGridLocales.DisplayLayout.Appearance = appearance56;
			this.uGridLocales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 55;
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.Width = 141;
			ultraGridColumn50.Header.Caption = "...";
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn50.Width = 29;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50});
			ultraGridBand3.Header.Caption = "LOCALES";
			ultraGridBand3.HeaderVisible = true;
			this.uGridLocales.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
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
			this.uGridLocales.Font = new System.Drawing.Font("Tahoma", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLocales.Location = new System.Drawing.Point(8, 240);
			this.uGridLocales.Name = "uGridLocales";
			this.uGridLocales.Size = new System.Drawing.Size(191, 328);
			this.uGridLocales.TabIndex = 742;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(bool);
			ultraDataColumn46.DefaultValue = false;
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(int);
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn53.DefaultValue = 0;
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
			ultraDataColumn61.DataType = typeof(System.Decimal);
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
			ultraDataColumn72.DataType = typeof(System.Decimal);
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
			ultraDataColumn78.DataType = typeof(int);
			ultraDataColumn79.DataType = typeof(int);
			ultraDataColumn80.DataType = typeof(int);
			ultraDataColumn80.DefaultValue = 0;
			ultraDataColumn81.DataType = typeof(System.Decimal);
			ultraDataColumn82.DataType = typeof(int);
			ultraDataColumn83.DataType = typeof(int);
			ultraDataColumn83.DefaultValue = 0;
			ultraDataColumn84.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn84});
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
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesExportar,
																																								this.mnuAccionesBloquear,
																																								this.mnuAccionesCerrar,
																																								this.mnuAccionesImprimir});
			// 
			// mnuAccionesExportar
			// 
			this.mnuAccionesExportar.Enabled = false;
			this.mnuAccionesExportar.Index = 0;
			this.mnuAccionesExportar.Text = "Exportar";
			this.mnuAccionesExportar.Click += new System.EventHandler(this.mnuAccionesExportar_Click_1);
			// 
			// mnuAccionesBloquear
			// 
			this.mnuAccionesBloquear.Enabled = false;
			this.mnuAccionesBloquear.Index = 1;
			this.mnuAccionesBloquear.Text = "Bloquear Periodo";
			this.mnuAccionesBloquear.Click += new System.EventHandler(this.mnuAccionesBloquear_Click_1);
			// 
			// mnuAccionesCerrar
			// 
			this.mnuAccionesCerrar.Enabled = false;
			this.mnuAccionesCerrar.Index = 2;
			this.mnuAccionesCerrar.Text = "Cerrar Periodo";
			this.mnuAccionesCerrar.Click += new System.EventHandler(this.mnuAccionesCerrar_Click_1);
			// 
			// mnuAccionesImprimir
			// 
			this.mnuAccionesImprimir.Enabled = false;
			this.mnuAccionesImprimir.Index = 3;
			this.mnuAccionesImprimir.Text = "Imprimir";
			this.mnuAccionesImprimir.Click += new System.EventHandler(this.mnuAccionesImprimir_Click_1);
			// 
			// frmCom_ResumenComisionesNuevo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1296, 566);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.uGridVendedores);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.chkCargos);
			this.Controls.Add(this.chkLocales);
			this.Controls.Add(this.uGridCargos);
			this.Controls.Add(this.optRegiones);
			this.Controls.Add(this.optConsulta);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridLocales);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.KeyPreview = true;
			this.Name = "frmCom_ResumenComisionesNuevo";
			this.Text = "Resumen Mensual De Comisiones";
			this.Load += new System.EventHandler(this.frmCom_ResumenComisionesNuevo_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optRegiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optConsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLocales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;
		int iSemana = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void  frmCom_ResumenComisionesNuevo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0739");
			
			if (!Funcion.Permiso("2144", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Resumen de Comisiones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				

				return;
			}
			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridCargos);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridLocales);

			#region Accesos
			if (miAcceso.BExportar) 
			{
				this.btnComandos.Enabled = true;
				this.mnuAccionesExportar.Enabled = true;
			}

			if (miAcceso.BCom_FiltrarPorCargo) FuncionesProcedimientos.EstadoGrids(true, this.uGridCargos);

			if (miAcceso.BCom_CerrarPeriodo)
			{
				this.btnComandos.Enabled = true;
				this.mnuAccionesBloquear.Enabled = true;
				this.mnuAccionesCerrar.Enabled = true;
				this.mnuAccionesImprimir.Enabled = true;
			}

			if (miAcceso.BCom_Observaciones)
			{
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnResumen"].Hidden = false;
			}

			#endregion Accesos
									
			DateTime dtFecha = DateTime.Parse("01/07/2018");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtHasta.Value = DateTime.Today;
			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			int iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);		

			if (iBodega == 0) 
			{
				FuncionesProcedimientos.EstadoGrids(true, this.uGridLocales);

				this.chkLocales.Enabled = true;

				this.uGridVendedores.Location = new Point(208, 72);				
				this.uGridVendedores.Size = new Size(1112, 488);

				this.optConsulta.Visible = true;
				this.optRegiones.Visible = true;
			}

			if (iBodega > 0)
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Almacen"].Hidden = true;

			this.optRegiones.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Com_ResumenSemanalValidaRegion {0}, {1}", MenuLatinium.IdUsuario, MenuLatinium.iNivel));
				
			if (MenuLatinium.iNivel != 23 && MenuLatinium.iNivel != 25) this.optConsulta.Enabled = true;

			this.Cargos();

			this.Locales();	

			if ((int)this.optRegiones.Value == 0) this.optRegiones.Enabled = true;
			//this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("19/07/2018"); //Convert.ToDateTime(this.dtDesde.Value);
			this.dtHasta.Value = DateTime.Today;			
			//this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
			//dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Text = "Resumen Mensual De Comisiones";
			string sLocales = "";
			
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

			string sCargos = "";

			if ((int)this.optConsulta.Value == 0 )
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCargos.Rows.All)
				{
					if ((bool) dr.Cells["Sel"].Value)
					{
						if (sCargos.Length == 0) sCargos = dr.Cells["idCargo"].Value.ToString();
						else if (sCargos.Length > 0) sCargos = sCargos + ", " + dr.Cells["idCargo"].Value.ToString();
					}
				}

				if (sCargos.Length == 0)
				{
					MessageBox.Show("Seleccione al menos un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

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

			string sSQL = "";

			if ((int)this.optConsulta.Value == 0)
				sSQL = string.Format("Exec Com_ResumenDeComisionesMensualesVentas '{0}', '{1}', '{2}', '{3}', '{4}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales, sCargos, this.txtBuscar.Text.ToString().Trim());
//			if ((int)this.optConsulta.Value == 1)
//				sSQL = string.Format("Exec Com_ResumenDeComisionesMensualesVentas '{0}', '{1}', '{2}'", 
//					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), sLocales);
//			if ((int)this.optConsulta.Value == 2)
//				sSQL = string.Format("Exec [Com_ResumenMensualPorRegion] '{0}', '{1}'", 
//					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));

			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridVendedores);

			this.Text = this.Text + " - " + this.uGridVendedores.Rows.Count + " Registros Encontrados";


			if (this.uGridVendedores.Rows != null && this.uGridVendedores.Rows.Count > 0)
			{
				this.DiseñoGridColor(this.uGridVendedores);
			}
		}

		private void DiseñoGridColor(UltraGrid uGrid)
		{
			foreach (UltraGridRow dr in uGrid.Rows)
			{
				if (dr.Cells["idCargo"].Value == null) continue; // Evita errores de nulos

				int idCargo = Convert.ToInt32(dr.Cells["idCargo"].Value);
				if (idCargo == 1 || idCargo == 4 || idCargo == 21)
				{
					decimal metaCuota0 = dr.Cells["MetaCuotaCero"].Value != DBNull.Value ? Convert.ToDecimal(dr.Cells["MetaCuotaCero"].Value) : 0;
					decimal porcentajeCuotaCeroFinal = dr.Cells["PorcentajeCuotaCeroFinal"].Value != DBNull.Value ? Convert.ToDecimal(dr.Cells["PorcentajeCuotaCeroFinal"].Value) : 0;
					decimal metaCarteraFlujo = dr.Cells["MetaCarteraFlujo"].Value != DBNull.Value ? Convert.ToDecimal(dr.Cells["MetaCarteraFlujo"].Value) : 0;
					decimal porcentajeCarteraFlujoFinal = dr.Cells["PorcentajeCarteraFlujoFinal"].Value != DBNull.Value ? Convert.ToDecimal(dr.Cells["PorcentajeCarteraFlujoFinal"].Value) : 0;

					decimal totalPirc0 = (metaCuota0 > 0) ? (porcentajeCuotaCeroFinal * 100) / metaCuota0 : 0;
					decimal totalCarteraFlujo = (metaCarteraFlujo > 0) ? (porcentajeCarteraFlujoFinal * 100) / metaCarteraFlujo : 0;

					AplicarColor(dr.Cells["PorcentajeCuotaCeroFinal"], totalPirc0);
					AplicarColor(dr.Cells["PorcentajeCarteraFlujoFinal"], totalCarteraFlujo);
				}
			}
		}

		private void AplicarColor(UltraGridCell cell, decimal valor)
		{
			if (cell.Value == null || cell.Value == DBNull.Value || cell.Value.ToString().Trim().Length == 0)
			{
				cell.Appearance.BackColor = Color.White;
				return;
			}

			if (valor < 85)
				cell.Appearance.BackColor = Color.Red;
			else if (valor < 95)
				cell.Appearance.BackColor = Color.Orange;
			else if (valor < 98)
				cell.Appearance.BackColor = Color.Yellow;
			else
				cell.Appearance.BackColor = Color.LimeGreen;
		}

		private void Cargos()
		{
			this.uGridCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ResumenComisionesListaCargos {0}, {1}, {2}", 
				MenuLatinium.IdUsuario, MenuLatinium.iNivel, (int)this.optRegiones.Value));
		}

		private void Locales()
		{
			string sSQL = string.Format("Exec Com_ResumenComisionesListaLocales {0}, {1}, {2}, '{3}'", 
				MenuLatinium.IdUsuario, iSemana, (int)this.optRegiones.Value, Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			this.uGridLocales.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/07/2017");
				this.dtHasta.CalendarInfo.MaxDate = DateTime.Parse("31/12/2030");

				if (Convert.ToDateTime(this.dtDesde.Value).Month == DateTime.Today.Month) 
				{
					this.dtHasta.Value = DateTime.Today;
					this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
				}
				else
				{
					this.dtHasta.Value = DateTime.DaysInMonth(Convert.ToDateTime(this.dtDesde.Value).Year, Convert.ToDateTime(this.dtDesde.Value).Month).ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
					this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
				}

				this.dtHasta.CalendarInfo.MinDate = Convert.ToDateTime(this.dtDesde.Value);						
			}
			else this.dtHasta.Value = System.DBNull.Value;
		}

		private void uGridVendedores_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if ((int)e.Row.Cells["idCargo"].Value != 4 && (int)e.Row.Cells["idCargo"].Value != 5 && (int)e.Row.Cells["idCargo"].Value != 7)
			{
				if ((int)e.Row.Cells["Contador"].Value == 1) e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png");
				if ((int)e.Row.Cells["Contador"].Value == 2)
				{
					e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
				}
				if ((int)e.Row.Cells["Contador"].Value == 3) 
				{
					e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
				}

				if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 100.00m) e.Row.Appearance.BackColor = Color.LimeGreen;
				else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 100.00m) e.Row.Appearance.BackColor = Color.Yellow;
				else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 85.00m) e.Row.Appearance.BackColor = Color.Orange;			
				else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 70.00m) e.Row.Appearance.BackColor = Color.Red;
			}
			//			if (e.Row.Cells["idCargo"].Value != System.DBNull.Value)
			//			{
			//				if ((int)e.Row.Cells["idCargo"].Value != 4 && (int)e.Row.Cells["idCargo"].Value != 5 && (int)e.Row.Cells["idCargo"].Value != 7)
			//				{
			//					if ((int)e.Row.Cells["Contador"].Value == 1) e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png");
			//					if ((int)e.Row.Cells["Contador"].Value == 2)
			//					{
			//						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
			//						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
			//					}
			//					if ((int)e.Row.Cells["Contador"].Value == 3) 
			//					{
			//						e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
			//						e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
			//						e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.187\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
			//					}					
			//			
			//					if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 100.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.LimeGreen;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 100.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.Yellow;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 85.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.Orange;			
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 70.00m) e.Row.Cells["PorcentajeDiario"].Appearance.BackColor = Color.Red;
			//
			//					if (Convert.ToDecimal(e.Row.Cells["PorcentajeMensual"].Value) >= 100.00m) e.Row.Cells["PorcentajeMensual"].Appearance.BackColor = Color.LimeGreen;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeMensual"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeMensual"].Value) < 100.00m) e.Row.Cells["PorcentajeMensual"].Appearance.BackColor = Color.Yellow;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeMensual"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeMensual"].Value) < 85.00m) e.Row.Cells["PorcentajeMensual"].Appearance.BackColor = Color.Orange;			
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeMensual"].Value) < 70.00m) e.Row.Cells["PorcentajeMensual"].Appearance.BackColor = Color.Red;
			//
			//					if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) >= 100.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.LimeGreen;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) < 100.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.Yellow;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) < 85.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.Orange;			
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeContado"].Value) < 70.00m) e.Row.Cells["PorcentajeContado"].Appearance.BackColor = Color.Red;
			//
			//					if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) >= 100.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.LimeGreen;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) < 100.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.Yellow;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) < 85.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.Orange;			
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeCredito"].Value) < 70.00m) e.Row.Cells["PorcentajeCredito"].Appearance.BackColor = Color.Red;
			//
			//					if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) >= 100.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.LimeGreen;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) < 100.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.Yellow;
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) < 85.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.Orange;			
			//					else if (Convert.ToDecimal(e.Row.Cells["PorcentajeGEx"].Value) < 70.00m) e.Row.Cells["PorcentajeGEx"].Appearance.BackColor = Color.Red;
			//				}
			//			}
		}

		private void uGridVendedores_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void dtDesde_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void optRegiones_ValueChanged(object sender, System.EventArgs e)
		{
			Cargos();

			Locales();

			this.btnVer_Click(sender, e);
		}

		private void uGridLocales_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkLocales_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
				dr.Cells["Sel"].Value = this.chkLocales.Checked;
		}

		private void chkCargos_CheckedChanged(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCargos.Rows.All)
				dr.Cells["Sel"].Value = this.chkCargos.Checked;
		}

		private void optConsulta_ValueChanged(object sender, System.EventArgs e)
		{
			if ((int)this.optConsulta.Value == 0 ) this.txtBuscar.Enabled = false;
			//this.txtBuscar.Enabled = !(bool)this.optConsulta.Value;
			this.chkCargos.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridCargos);

			if ((int)this.optConsulta.Value == 1) this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendedor"].Header.Caption = "Jefe De Almacen";
			if ((int)this.optConsulta.Value == 1) this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendedor"].Header.Caption = "Jefe Regional";
			
			if ((int)this.optConsulta.Value == 0) 
			{
				this.chkCargos.Enabled = true;
				FuncionesProcedimientos.EstadoGrids(true, this.uGridCargos);
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendedor"].Header.Caption = "Vendedor";
			}
			bool bConsultaEstado = false;
			if ((int) this.optConsulta.Value >= 1  ) bConsultaEstado = true;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["Cargo"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionMensual"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionCredito"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionCarpa"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["BonoCarpa"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionGEx"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["ComisionDyD"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["Comision"].Hidden = bConsultaEstado;

			this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estrella1"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estrella2"].Hidden = bConsultaEstado;
			this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estrella3"].Hidden = bConsultaEstado;	
		
			this.btnVer_Click(sender, e);
		}

		private void frmCom_ResumenMensual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridVendedores_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnResumen")
			{	
				frmCom_ResumenMensualPorVendedor RMPV = new frmCom_ResumenMensualPorVendedor((int)e.Cell.Row.Cells["idPersonal"].Value);
				RMPV.ShowDialog();
			}
		}

		private void mnuAccionesExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridVendedores);
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			this.lblEstado.Text = "";
			this.mnuAccionesImprimir.Enabled = false;
			this.mnuAccionesBloquear.Enabled = false;					
			this.mnuAccionesCerrar.Enabled = true;				

			if (miAcceso.BCom_CerrarPeriodo)
			{
				if (!Funcion.ValidaPeriodo(Convert.ToDateTime(this.dtHasta.Value), cdsSeteoF, "Venta"))
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec Com_ValidaUltimoDiaMes '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")), true))
					{
						this.mnuAccionesBloquear.Enabled = true;
					}
				}
				else
				{
					this.lblEstado.Text = "PERIODO BLOQUEADO";

					int iMes = Convert.ToDateTime(this.dtHasta.Value).Month;
					int iAño = Convert.ToDateTime(this.dtHasta.Value).Year;
					string sPeriodo = iMes.ToString() + "/" + iAño.ToString();

					if (!Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec Com_ValidaPeriodoCerrado '{0}'", sPeriodo), true))
					{
						this.mnuAccionesCerrar.Enabled = true;				
					}
					else
					{
						this.lblEstado.Text = "PERIODO CERRADO";
						this.mnuAccionesImprimir.Enabled = true;
					}
				}
			}
		}

		private void btnComandos_Click_1(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void uGridVendedores_Click(object sender, System.EventArgs e)
		{
		
		}

		private void uGridVendedores_ClickCellButton_1(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnResumen")
			{	
				frmCom_ResumenMensualPorVendedor RMPV = new frmCom_ResumenMensualPorVendedor((int)e.Cell.Row.Cells["idPersonal"].Value);
				RMPV.ShowDialog();
			}
		}

		private void uGridVendedores_InitializeRow_1(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if ((int)e.Row.Cells["idCargo"].Value != 4 && (int)e.Row.Cells["idCargo"].Value != 5 && (int)e.Row.Cells["idCargo"].Value != 7)
			{
				if ((int)e.Row.Cells["Contador"].Value == 1) e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png");
				if ((int)e.Row.Cells["Contador"].Value == 2)
				{
					e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
				}
				if ((int)e.Row.Cells["Contador"].Value == 3) 
				{
					e.Row.Cells["Estrella1"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					e.Row.Cells["Estrella2"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
					e.Row.Cells["Estrella3"].ButtonAppearance.Image = Image.FromFile(@"\\192.168.1.10\Proyectos\Latinium\Imagenes\Estrella_Comisiones\estrella_comisiones.png"); 
				}

				if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 100.00m) e.Row.Appearance.BackColor = Color.LimeGreen;
				else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 85.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 100.00m) e.Row.Appearance.BackColor = Color.Yellow;
				else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) >= 70.00m && Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 85.00m) e.Row.Appearance.BackColor = Color.Orange;			
				else if (Convert.ToDecimal(e.Row.Cells["PorcentajeDiario"].Value) < 70.00m) e.Row.Appearance.BackColor = Color.Red;
			}

		}

		private void mnuAccionesExportar_Click_1(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridVendedores);
		}

		private void mnuAccionesBloquear_Click_1(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de Bloquear el periodo del '{0}' al '{1}'\n\n", Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy"), Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				string sSQL = string.Format("Exec Com_CierraPeriodoVentas '{0}', 1", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				MessageBox.Show("Periodo Bloqueado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.mnuAccionesBloquear.Enabled = false;
				this.mnuAccionesCerrar.Enabled = true;
			}
		}

		private void mnuAccionesCerrar_Click_1(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.uGridVendedores.Rows.Count == 0)
			{
				MessageBox.Show("No existen registros de vendedores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}

			if ((int)this.optConsulta.Value == 1)
			{
				MessageBox.Show("Para cerrar comsiones debe seleccionar la opción RESUMEN POR VENDEDOR", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optConsulta.Focus();
				return;
			}

			if ((int)this.optRegiones.Value != 0)
			{
				MessageBox.Show("Para cerrar comsiones debe seleccionar la opción TODOS en regiones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.optRegiones.Focus();
				return;
			}

			int iLocales = 0;
			int iCargos = 0;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLocales.Rows.All)
				if (!(bool) dr.Cells["Sel"].Value) iLocales++;

			if (iLocales > 0)
			{
				MessageBox.Show("Para cerrar comsiones debe seleccionar todos los almacenes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
						
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCargos.Rows.All)
				if (!(bool) dr.Cells["Sel"].Value) iCargos++;

			if (iCargos > 0)
			{
				MessageBox.Show("Para cerrar comsiones debe seleccionar todos los cargos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (Convert.ToDateTime(this.dtDesde.Value) != DateTime.Parse("01/" + Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString()))
			{
				MessageBox.Show("La fecha inicial de consulta debe ser el primer dia del mes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesde.Focus();
				return;
			}

			if (Convert.ToDateTime(this.dtHasta.Value) != DateTime.Parse(DateTime.DaysInMonth(Convert.ToDateTime(this.dtHasta.Value).Year, Convert.ToDateTime(this.dtHasta.Value).Month) + "/" + Convert.ToDateTime(this.dtHasta.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtHasta.Value).Year.ToString()))
			{
				MessageBox.Show("La fecha final de consulta debe ser el ultimo dia del mes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtHasta.Focus();
				return;
			}

			if (this.txtBuscar.Text.ToString().Trim().Length > 0)
			{
				MessageBox.Show("Para cerrar el Periodo no puede tener filtros de busqueda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtBuscar.Focus();
				return;
			}
			#endregion Validacion
			
			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de Cerrar el periodo del '{0}' al '{1}'\n\n", Convert.ToDateTime(this.dtDesde.Value).ToString("dd/MM/yyyy"), Convert.ToDateTime(this.dtHasta.Value).ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				#region Guarda Cierre de Periodo
				string sPeriodo = Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();
				bool bComisiona = false;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridVendedores.Rows.All)
				{
					#region Valida Estrellas
					int iAccion = 1;
					int iEstrellas = (int)dr.Cells["Contador"].Value; 
					int idCom_Rango = (int)dr.Cells["idCom_Rango"].Value;

					if (Convert.ToDecimal(dr.Cells["PorcentajeMensual"].Value) >= 100.00m) 
					{
						bComisiona = true; 
						if (iEstrellas < 3) iEstrellas++;
					}
					else 
					{
						bComisiona = false;
						if (iEstrellas > 0) iEstrellas--;
						if (iEstrellas == 1) iEstrellas = 0;
					}
					#endregion Valida Estrellas

					#region Cierra Periodo
					/*@idCom_ResumenMensualComisiones Int, @Periodo Varchar(7), @Desde Datetime, @Hasta Datetime, @Bodega Int, @idPersonal Int, @idCargo Int, @idCom_Rango Int, 7 
	@Meta Decimal(18, 2), @MetaCarpa Decimal(18, 2), @MetaCredito Decimal(18, 2), @MetaGEx Decimal(18, 2), 11
	@TotalVendido Decimal(18, 2), @Carpa Decimal(18, 2), @VentasCarpa Int, @VentaCredito Decimal(18, 2), @VentaGEx Decimal(18, 2), 16
	@Porcentaje Decimal(18, 2), @PorcentajeCredito Decimal(18, 2), @PorcentajeCarpa Decimal(18, 2), @PorcentajeGEx Decimal(18, 2), 20
	@ComisionGeneral Decimal(18, 2), @ComisionCredito Decimal(18, 2), @ComisionCarpa Decimal(18, 2), @BonoCarpa Decimal(18, 2), @ComisionGEx Decimal(18, 2), @Comision Decimal(18, 2), 26
	@Comisiona Bit, @Estado Int, @Estrellas Int, @Numero Int 30*/

					/*idCom_AsignacionDeVendedores, Bodega, Almacen, idCargo, Cargo, idCom_Rango, idPersonal, Vendedor, 
		MetaDiaria, MetaMensual, MetaCarpa, MetaCredito, MetaGEx,
		VentaTotal, VentaCredito, VentaCarpa, VentasCarpa, VentaGEx, 
		Diferencia, PorcentajeDiario, PorcentajeMensual, PorcentajeCredito, PorcentajeCarpa, PorcentajeGEx,
		ISNULL(ComisionMensual, 0) As ComisionMensual, ISNULL(ComisionCredito, 0) As ComisionCredito, ISNULL(ComisionCarpa, 0) As ComisionCarpa, ISNULL(BonoCarpa, 0) As BonoCarpa, ISNULL(ComisionGEx, 0) As ComisionGEx, 
		ISNULL(Comision, 0) As Comision, ISNULL(Contador, 0) As Contador, Numero*/
					//
					string sSQL = string.Format("Exec Com_CierraPeriodo {0}, '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}",
						0, sPeriodo, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
						(int)dr.Cells["Bodega"].Value, (int)dr.Cells["idPersonal"].Value, (int)dr.Cells["idCargo"].Value, (int)dr.Cells["idCom_Rango"].Value,
						Convert.ToDecimal(dr.Cells["MetaMensual"].Value), Convert.ToDecimal(dr.Cells["MetaCarpa"].Value), Convert.ToDecimal(dr.Cells["MetaCredito"].Value), Convert.ToDecimal(dr.Cells["MetaGEx"].Value), 
						Convert.ToDecimal(dr.Cells["VentaTotal"].Value), Convert.ToDecimal(dr.Cells["VentaCarpa"].Value), Convert.ToDecimal(dr.Cells["VentasCarpa"].Value), Convert.ToDecimal(dr.Cells["VentaCredito"].Value), Convert.ToDecimal(dr.Cells["VentaGEx"].Value),
						Convert.ToDecimal(dr.Cells["PorcentajeMensual"].Value), Convert.ToDecimal(dr.Cells["PorcentajeCredito"].Value), Convert.ToDecimal(dr.Cells["PorcentajeCarpa"].Value), Convert.ToDecimal(dr.Cells["PorcentajeGEx"].Value), 
						Convert.ToDecimal(dr.Cells["ComisionMensual"].Value), Convert.ToDecimal(dr.Cells["ComisionCredito"].Value), Convert.ToDecimal(dr.Cells["ComisionCarpa"].Value), Convert.ToDecimal(dr.Cells["BonoCarpa"].Value), Convert.ToDecimal(dr.Cells["ComisionGEx"].Value), Convert.ToDecimal(dr.Cells["Comision"].Value),
						bComisiona, 1, iEstrellas, (int)dr.Cells["Numero"].Value,
						Convert.ToDecimal(dr.Cells["PorcentajeCuotaCeroFinal"].Value), Convert.ToDecimal(dr.Cells["PorcentajeCarteraFlujoFinal"].Value),
						Convert.ToDecimal(dr.Cells["PorcentajeCarteraRatioFinal"].Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
					#endregion Cierra Periodo

					#region Valida Rango Vendedor
					if (iEstrellas == 3)
					{
						if (Convert.ToInt32( dr.Cells["idCom_Rango"].Value) == 2)
						{
							iAccion = 2;
							idCom_Rango = 1;
						}						
					}
					else if (iEstrellas < 2)
					{
						if ((int)dr.Cells["idCom_Rango"].Value == 1)
						{
							iAccion = 3;
							idCom_Rango = 2;							
						}						
					}					
					#endregion Valida Rango Vendedor
					
					#region Inserta Historial
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Meta From Com_CargosDeVentas Where idCargo = {0}", (int)dr.Cells["idCargo"].Value), true))
					{
						string sSQLH = string.Format("Exec Com_GuardaHistorialRango '{0}', {1}, {2}, {3}, {4}, {5}", 
							sPeriodo, (int)dr.Cells["idPersonal"].Value, (int)dr.Cells["Bodega"].Value, (int)dr.Cells["idCargo"].Value, idCom_Rango, iAccion);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLH);	
					}
					#endregion Inserta Historial
				}
				#endregion Guarda Cierre de Periodo

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Com_BloqueaFacturas '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
								
				string sSQLActRangos = string.Format("Exec Com_CierrePeriodoActualizaRango '{0}'", sPeriodo);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActRangos);

				string sSQLMail = string.Format("Exec Com_MailRangosVendedores '{0}'", sPeriodo);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLMail);

				MessageBox.Show("Periodo Cerrado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.mnuAccionesImprimir_Click_1(sender, e);

				this.mnuAccionesCerrar.Enabled = false;
			}
		}

		private void mnuAccionesImprimir_Click_1(object sender, System.EventArgs e)
		{
			string sPeriodo = Convert.ToDateTime(this.dtDesde.Value).Month.ToString() + "/" + Convert.ToDateTime(this.dtDesde.Value).Year.ToString();

			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Periodo = new ParameterField ();
			ParameterDiscreteValue dPeriodo = new ParameterDiscreteValue ();
			Periodo.ParameterFieldName = "@Periodo";
			dPeriodo.Value = sPeriodo;
			Periodo.CurrentValues.Add (dPeriodo);
			paramFields.Add (Periodo);
			
			Reporte miRep = new Reporte("Com_ResumenCierrePeriodo.rpt", "");
			miRep.MdiParent = this.MdiParent;			
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
						
			Cursor = Cursors.Default;
		}

		private void optConsulta_ValueChanged_1(object sender, System.EventArgs e)
		{

		}

	
		private void optConsulta_Click(object sender, System.EventArgs e)
		{

		}
	
	}
}
