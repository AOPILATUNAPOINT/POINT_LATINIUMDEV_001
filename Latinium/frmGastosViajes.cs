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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmGastosViajes.
	/// </summary>
	public class frmGastosViajes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCedula;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.GroupBox groupBox2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaSalida;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRetorna;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAcreditado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorS;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtLDestino;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridDetalle;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidGastosviajes;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPersonal;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEgreso;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidDeposito;
		private System.Windows.Forms.Label lblEstado;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsiento;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEgresoReposicion;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCobradores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentroCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnDeposito;
		private System.Windows.Forms.Button btnEmpresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Button btnRevisar;
		public System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label67;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtRetorna;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtInicia;
		private System.Windows.Forms.Label label15;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigen;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTransporte;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.LinkLabel linkLabel1;

		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmGastosViajes()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmGastosViajes));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleGViaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosViaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCobradores");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RazonSocial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serie");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SF");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 5, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleGViaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosViaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RazonSocial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serie");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SF");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngresoReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosViajes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Definido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngresoReembolsoViaticos", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGastosViajes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Definido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AdministrativoSolicitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GerentesSolicitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Administrativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gerentes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotalAcreditado");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngresoReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosViajes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Definido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AdministrativoSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GerentesSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Administrativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gerentes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotalAcreditado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngresoReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGastosViajes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idReembolsoViaticos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Definido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AdministrativoSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GerenteSolicitado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Administrativo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Gerentes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotalAcreditado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Linea_Producto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio_Contado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoPaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoMaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sexo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto_Credito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor_Seguro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total_Deuda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Calificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Vivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo_Vivienda_Meses");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa_Trabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo_Trabajo_Meses");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Negocio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo_Negocio_Meses");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFechaSalida = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCedula = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.dtFechaRetorna = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnAnular = new System.Windows.Forms.Button();
			this.txtAcreditado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtValorS = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtLDestino = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.uGridDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtidGastosviajes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidEgreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidDeposito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtidAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidEgresoReposicion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbCobradores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnDeposito = new System.Windows.Forms.Button();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnEmpresa = new System.Windows.Forms.Button();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnRevisar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.label67 = new System.Windows.Forms.Label();
			this.dtRetorna = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.dtInicia = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label15 = new System.Windows.Forms.Label();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.optTransporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label16 = new System.Windows.Forms.Label();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetorna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAcreditado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorS)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGastosviajes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDeposito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEgresoReposicion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRetorna)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(248, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 17);
			this.label1.TabIndex = 200;
			this.label1.Text = "Cedula";
			// 
			// dtFechaSalida
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaSalida.Appearance = appearance1;
			this.dtFechaSalida.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaSalida.DateButtons.Add(dateButton1);
			this.dtFechaSalida.Enabled = false;
			this.dtFechaSalida.Format = "dd/MM/yyyy";
			this.dtFechaSalida.Location = new System.Drawing.Point(96, 38);
			this.dtFechaSalida.Name = "dtFechaSalida";
			this.dtFechaSalida.NonAutoSizeHeight = 23;
			this.dtFechaSalida.Size = new System.Drawing.Size(136, 21);
			this.dtFechaSalida.SpinButtonsVisible = true;
			this.dtFechaSalida.TabIndex = 198;
			this.dtFechaSalida.Value = ((object)(resources.GetObject("dtFechaSalida.Value")));
			this.dtFechaSalida.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaSalida_KeyDown);
			this.dtFechaSalida.ValueChanged += new System.EventHandler(this.dtFechaSalida_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 199;
			this.label6.Text = "Fecha";
			// 
			// dtFecha
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance2;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(136, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 197;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(456, 11);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 17);
			this.label3.TabIndex = 778;
			this.label3.Text = "Nombre";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance3;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(552, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(296, 22);
			this.txtNombre.TabIndex = 784;
			// 
			// txtCedula
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCedula.Appearance = appearance4;
			this.txtCedula.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCedula.Enabled = false;
			this.txtCedula.Location = new System.Drawing.Point(304, 8);
			this.txtCedula.MaxLength = 10;
			this.txtCedula.Name = "txtCedula";
			this.txtCedula.Size = new System.Drawing.Size(136, 22);
			this.txtCedula.TabIndex = 783;
			this.txtCedula.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCedula_KeyDown);
			this.txtCedula.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCedula_KeyPress);
			this.txtCedula.Validating += new System.ComponentModel.CancelEventHandler(this.txtCedula_Validating);
			this.txtCedula.Validated += new System.EventHandler(this.txtCedula_Validated);
			this.txtCedula.ValueChanged += new System.EventHandler(this.txtCedula_ValueChanged);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(8, 124);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(85, 17);
			this.label9.TabIndex = 786;
			this.label9.Text = "Lugar de Origen";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(8, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(65, 16);
			this.label5.TabIndex = 788;
			this.label5.Text = "Fecha Inicia";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(1024, 136);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(76, 17);
			this.label7.TabIndex = 790;
			this.label7.Text = "Observaciones";
			this.label7.Visible = false;
			// 
			// txtObservaciones
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance5;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(488, 200);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(32, 16);
			this.txtObservaciones.TabIndex = 789;
			this.txtObservaciones.Visible = false;
			this.txtObservaciones.ValueChanged += new System.EventHandler(this.txtObservaciones_ValueChanged);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(456, 40);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(79, 16);
			this.label8.TabIndex = 792;
			this.label8.Text = "Fecha Retorna";
			// 
			// dtFechaRetorna
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRetorna.Appearance = appearance6;
			this.dtFechaRetorna.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaRetorna.DateButtons.Add(dateButton3);
			this.dtFechaRetorna.Enabled = false;
			this.dtFechaRetorna.Format = "dd/MM/yyyy";
			this.dtFechaRetorna.Location = new System.Drawing.Point(552, 38);
			this.dtFechaRetorna.Name = "dtFechaRetorna";
			this.dtFechaRetorna.NonAutoSizeHeight = 23;
			this.dtFechaRetorna.Size = new System.Drawing.Size(136, 21);
			this.dtFechaRetorna.SpinButtonsVisible = true;
			this.dtFechaRetorna.TabIndex = 791;
			this.dtFechaRetorna.Value = ((object)(resources.GetObject("dtFechaRetorna.Value")));
			this.dtFechaRetorna.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaRetorna_KeyDown);
			this.dtFechaRetorna.ValueChanged += new System.EventHandler(this.dtFechaRetorna_ValueChanged);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(744, 472);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(80, 23);
			this.btnAnular.TabIndex = 858;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			this.btnAnular.Click += new System.EventHandler(this.btnRechazar_Click);
			// 
			// txtAcreditado
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAcreditado.Appearance = appearance7;
			this.txtAcreditado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAcreditado.Enabled = false;
			this.txtAcreditado.FormatString = "#,##0.00";
			this.txtAcreditado.Location = new System.Drawing.Point(344, 64);
			this.txtAcreditado.MinValue = 0;
			this.txtAcreditado.Name = "txtAcreditado";
			this.txtAcreditado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAcreditado.PromptChar = ' ';
			this.txtAcreditado.Size = new System.Drawing.Size(96, 22);
			this.txtAcreditado.TabIndex = 854;
			this.txtAcreditado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAcreditado_KeyDown);
			this.txtAcreditado.ValueChanged += new System.EventHandler(this.txtAcreditado_ValueChanged);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Cuenta";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Enabled = false;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(192, 192);
			this.cmbCuenta.MaxDropDownItems = 30;
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(24, 21);
			this.cmbCuenta.TabIndex = 843;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			this.cmbCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCuenta_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(224, 192);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(78, 17);
			this.label13.TabIndex = 841;
			this.label13.Text = "Valor Autorizar";
			this.label13.Visible = false;
			// 
			// cmbFormaPago
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance8;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource4;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 176;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(104, 192);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(24, 21);
			this.cmbFormaPago.TabIndex = 839;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			// 
			// ultraDataSource4
			// 
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(8, 192);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(80, 17);
			this.label10.TabIndex = 836;
			this.label10.Text = "Forma de Pago";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label10.Visible = false;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(144, 192);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(39, 17);
			this.label11.TabIndex = 833;
			this.label11.Text = "Cuenta";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Visible = false;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(8, 67);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(79, 17);
			this.label12.TabIndex = 839;
			this.label12.Text = "90% Solicitado";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance9;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(1235, 472);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 26);
			this.txtTotal.TabIndex = 844;
			this.txtTotal.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 472);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 24);
			this.btnNuevo.TabIndex = 851;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(96, 472);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(79, 24);
			this.btnConsultar.TabIndex = 850;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(269, 472);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(77, 24);
			this.btnEditar.TabIndex = 849;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(354, 472);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(84, 24);
			this.btnGuardar.TabIndex = 848;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(537, 472);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(82, 24);
			this.btnCancelar.TabIndex = 846;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 176);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1323, 8);
			this.groupBox2.TabIndex = 852;
			this.groupBox2.TabStop = false;
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
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(720, 200);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(56, 22);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 855;
			this.txtEstado.Visible = false;
			// 
			// txtValorS
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorS.Appearance = appearance10;
			this.txtValorS.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorS.Enabled = false;
			this.txtValorS.FormatString = "#,##0.00";
			this.txtValorS.Location = new System.Drawing.Point(96, 64);
			this.txtValorS.Name = "txtValorS";
			this.txtValorS.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorS.PromptChar = ' ';
			this.txtValorS.Size = new System.Drawing.Size(136, 22);
			this.txtValorS.TabIndex = 856;
			this.txtValorS.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValorS_KeyDown);
			this.txtValorS.ValueChanged += new System.EventHandler(this.txtValorS_ValueChanged);
			// 
			// txtLDestino
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLDestino.Appearance = appearance11;
			this.txtLDestino.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLDestino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLDestino.Enabled = false;
			this.txtLDestino.Location = new System.Drawing.Point(552, 96);
			this.txtLDestino.Multiline = true;
			this.txtLDestino.Name = "txtLDestino";
			this.txtLDestino.Size = new System.Drawing.Size(296, 72);
			this.txtLDestino.TabIndex = 857;
			this.txtLDestino.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtLDestino_KeyDown_1);
			// 
			// uGridDetalle
			// 
			this.uGridDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDetalle.DataSource = this.ultraDataSource1;
			appearance12.BackColor = System.Drawing.Color.White;
			this.uGridDetalle.DisplayLayout.Appearance = appearance12;
			this.uGridDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 119;
			ultraGridColumn6.Header.VisiblePosition = 4;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 122;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.Header.Caption = "Cuenta";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 220;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Width = 200;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance13;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 13;
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.Header.Caption = "Centro De Costo";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn11.Width = 101;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 69;
			ultraGridColumn13.FieldLen = 13;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.Caption = "Proveedor";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.Width = 250;
			ultraGridColumn15.FieldLen = 6;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 70;
			ultraGridColumn16.FieldLen = 9;
			ultraGridColumn16.Header.Caption = "Número";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.FieldLen = 37;
			ultraGridColumn17.Header.Caption = "Autorización";
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.Width = 90;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 30;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance14;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand3.SummaryFooterCaption = "Totales";
			this.uGridDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridDetalle.DisplayLayout.Override.CardAreaAppearance = appearance16;
			this.uGridDetalle.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 8.5F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDetalle.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridDetalle.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDetalle.Location = new System.Drawing.Point(8, 224);
			this.uGridDetalle.Name = "uGridDetalle";
			this.uGridDetalle.Size = new System.Drawing.Size(1307, 240);
			this.uGridDetalle.TabIndex = 858;
			this.uGridDetalle.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDetalle_AfterRowInsert);
			this.uGridDetalle.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDetalle_KeyPress);
			this.uGridDetalle.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDetalle_BeforeCellUpdate);
			this.uGridDetalle.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDetalle_BeforeRowsDeleted);
			this.uGridDetalle.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDetalle_KeyDown);
			this.uGridDetalle.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDetalle_AfterCellUpdate);
			this.uGridDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDetalle_InitializeLayout);
			this.uGridDetalle.Click += new System.EventHandler(this.p);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(bool);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn18});
			// 
			// txtidGastosviajes
			// 
			this.txtidGastosviajes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidGastosviajes.Enabled = false;
			this.txtidGastosviajes.Location = new System.Drawing.Point(656, 200);
			this.txtidGastosviajes.Name = "txtidGastosviajes";
			this.txtidGastosviajes.PromptChar = ' ';
			this.txtidGastosviajes.Size = new System.Drawing.Size(56, 22);
			this.txtidGastosviajes.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidGastosviajes.TabIndex = 861;
			this.txtidGastosviajes.Visible = false;
			// 
			// txtidPersonal
			// 
			this.txtidPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPersonal.Enabled = false;
			this.txtidPersonal.Location = new System.Drawing.Point(784, 200);
			this.txtidPersonal.Name = "txtidPersonal";
			this.txtidPersonal.PromptChar = ' ';
			this.txtidPersonal.Size = new System.Drawing.Size(56, 22);
			this.txtidPersonal.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidPersonal.TabIndex = 864;
			this.txtidPersonal.Visible = false;
			// 
			// txtidEgreso
			// 
			this.txtidEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEgreso.Enabled = false;
			this.txtidEgreso.Location = new System.Drawing.Point(528, 200);
			this.txtidEgreso.Name = "txtidEgreso";
			this.txtidEgreso.PromptChar = ' ';
			this.txtidEgreso.Size = new System.Drawing.Size(56, 22);
			this.txtidEgreso.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidEgreso.TabIndex = 865;
			this.txtidEgreso.Visible = false;
			// 
			// txtidDeposito
			// 
			this.txtidDeposito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidDeposito.Enabled = false;
			this.txtidDeposito.Location = new System.Drawing.Point(592, 200);
			this.txtidDeposito.Name = "txtidDeposito";
			this.txtidDeposito.PromptChar = ' ';
			this.txtidDeposito.Size = new System.Drawing.Size(56, 22);
			this.txtidDeposito.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidDeposito.TabIndex = 866;
			this.txtidDeposito.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(1067, 69);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 867;
			// 
			// txtidAsiento
			// 
			this.txtidAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAsiento.Enabled = false;
			this.txtidAsiento.Location = new System.Drawing.Point(856, 200);
			this.txtidAsiento.Name = "txtidAsiento";
			this.txtidAsiento.PromptChar = ' ';
			this.txtidAsiento.Size = new System.Drawing.Size(56, 22);
			this.txtidAsiento.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidAsiento.TabIndex = 868;
			this.txtidAsiento.Visible = false;
			// 
			// txtidEgresoReposicion
			// 
			this.txtidEgresoReposicion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEgresoReposicion.Enabled = false;
			this.txtidEgresoReposicion.Location = new System.Drawing.Point(920, 336);
			this.txtidEgresoReposicion.Name = "txtidEgresoReposicion";
			this.txtidEgresoReposicion.PromptChar = ' ';
			this.txtidEgresoReposicion.Size = new System.Drawing.Size(56, 22);
			this.txtidEgresoReposicion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidEgresoReposicion.TabIndex = 869;
			this.txtidEgresoReposicion.Visible = false;
			// 
			// cmbCobradores
			// 
			this.cmbCobradores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCobradores.DataSource = this.ultraDataSource3;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 240;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCobradores.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCobradores.DisplayMember = "Descripcion";
			this.cmbCobradores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCobradores.Location = new System.Drawing.Point(448, 352);
			this.cmbCobradores.Name = "cmbCobradores";
			this.cmbCobradores.Size = new System.Drawing.Size(112, 56);
			this.cmbCobradores.TabIndex = 981;
			this.cmbCobradores.ValueMember = "idCuenta";
			this.cmbCobradores.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(1147, 10);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 983;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance21;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(1203, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 23);
			this.txtNumero.TabIndex = 982;
			// 
			// btnDeposito
			// 
			this.btnDeposito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDeposito.CausesValidation = false;
			this.btnDeposito.Enabled = false;
			this.btnDeposito.Image = ((System.Drawing.Image)(resources.GetObject("btnDeposito.Image")));
			this.btnDeposito.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDeposito.Location = new System.Drawing.Point(446, 472);
			this.btnDeposito.Name = "btnDeposito";
			this.btnDeposito.Size = new System.Drawing.Size(83, 24);
			this.btnDeposito.TabIndex = 984;
			this.btnDeposito.Text = "&Deposito";
			this.btnDeposito.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDeposito.Click += new System.EventHandler(this.btnDeposito_Click);
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataSource = this.ultraDataSource5;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(568, 352);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(160, 56);
			this.cmbCentroCosto.TabIndex = 985;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			this.cmbCentroCosto.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(627, 472);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(81, 24);
			this.btnSalir.TabIndex = 986;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnEmpresa
			// 
			this.btnEmpresa.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.btnEmpresa.Enabled = false;
			this.btnEmpresa.Image = ((System.Drawing.Image)(resources.GetObject("btnEmpresa.Image")));
			this.btnEmpresa.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEmpresa.Location = new System.Drawing.Point(704, 63);
			this.btnEmpresa.Name = "btnEmpresa";
			this.btnEmpresa.Size = new System.Drawing.Size(25, 24);
			this.btnEmpresa.TabIndex = 1007;
			this.btnEmpresa.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEmpresa.Click += new System.EventHandler(this.btnEmpresa_Click);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(System.Decimal);
			ultraDataColumn28.DataType = typeof(System.Decimal);
			ultraDataColumn29.DataType = typeof(bool);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(bool);
			ultraDataColumn32.DataType = typeof(System.Decimal);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource8;
			appearance22.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance22;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn25.Header.VisiblePosition = 2;
			ultraGridColumn26.Header.VisiblePosition = 4;
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn28.Header.VisiblePosition = 5;
			ultraGridColumn29.Header.VisiblePosition = 6;
			ultraGridColumn30.Header.VisiblePosition = 7;
			ultraGridColumn31.Header.VisiblePosition = 8;
			ultraGridColumn32.Header.VisiblePosition = 10;
			ultraGridColumn33.Header.VisiblePosition = 11;
			ultraGridColumn34.Header.VisiblePosition = 9;
			ultraGridColumn35.Header.VisiblePosition = 12;
			ultraGridBand6.Columns.AddRange(new object[] {
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
																										 ultraGridColumn35});
			ultraGridBand6.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance24.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance24;
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance26.FontData.BoldAsString = "True";
			appearance26.FontData.Name = "Arial";
			appearance26.FontData.SizeInPoints = 8F;
			appearance26.ForeColor = System.Drawing.Color.White;
			appearance26.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance26;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance27.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance27.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(880, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(435, 64);
			this.ultraGrid1.TabIndex = 1008;
			this.ultraGrid1.Visible = false;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn40.DataType = typeof(System.Decimal);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn43.DataType = typeof(bool);
			ultraDataColumn44.DataType = typeof(bool);
			ultraDataColumn45.DataType = typeof(System.Decimal);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn45});
			// 
			// ultraDataSource7
			// 
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(System.Decimal);
			ultraDataColumn51.DataType = typeof(System.Decimal);
			ultraDataColumn52.DataType = typeof(System.Decimal);
			ultraDataColumn53.DataType = typeof(System.Decimal);
			ultraDataColumn54.DataType = typeof(System.Decimal);
			ultraDataColumn55.DataType = typeof(System.Decimal);
			ultraDataColumn56.DataType = typeof(bool);
			ultraDataColumn57.DataType = typeof(bool);
			ultraDataColumn58.DataType = typeof(System.Decimal);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn58});
			// 
			// btnRevisar
			// 
			this.btnRevisar.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.btnRevisar.Enabled = false;
			this.btnRevisar.Image = ((System.Drawing.Image)(resources.GetObject("btnRevisar.Image")));
			this.btnRevisar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnRevisar.Location = new System.Drawing.Point(621, 192);
			this.btnRevisar.Name = "btnRevisar";
			this.btnRevisar.Size = new System.Drawing.Size(25, 24);
			this.btnRevisar.TabIndex = 1009;
			this.btnRevisar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnRevisar.Visible = false;
			this.btnRevisar.Click += new System.EventHandler(this.btnRevisar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(183, 472);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(78, 24);
			this.btnImprimir.TabIndex = 1010;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn59.DataType = typeof(int);
			ultraDataColumn60.DataType = typeof(int);
			ultraDataColumn63.DataType = typeof(System.Decimal);
			ultraDataColumn64.DataType = typeof(System.Decimal);
			ultraDataColumn69.DataType = typeof(System.DateTime);
			ultraDataColumn71.DataType = typeof(System.DateTime);
			ultraDataColumn74.DataType = typeof(System.Decimal);
			ultraDataColumn76.DataType = typeof(System.Decimal);
			ultraDataColumn77.DataType = typeof(System.Decimal);
			ultraDataColumn78.DataType = typeof(int);
			ultraDataColumn81.DataType = typeof(int);
			ultraDataColumn84.DataType = typeof(int);
			ultraDataColumn86.DataType = typeof(int);
			ultraDataColumn87.DataType = typeof(System.Decimal);
			ultraDataColumn88.DataType = typeof(bool);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn84,
																																 ultraDataColumn85,
																																 ultraDataColumn86,
																																 ultraDataColumn87,
																																 ultraDataColumn88});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(248, 67);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(86, 17);
			this.label2.TabIndex = 1011;
			this.label2.Text = "Valor Acreditado";
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(248, 40);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(28, 17);
			this.label67.TabIndex = 1012;
			this.label67.Text = "Hora";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtRetorna
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtRetorna.Appearance = appearance30;
			this.dtRetorna.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtRetorna.DateButtons.Add(dateButton4);
			this.dtRetorna.Enabled = false;
			this.dtRetorna.Format = "HH:mm";
			this.dtRetorna.Location = new System.Drawing.Point(744, 38);
			this.dtRetorna.Name = "dtRetorna";
			this.dtRetorna.NonAutoSizeHeight = 23;
			this.dtRetorna.Size = new System.Drawing.Size(88, 21);
			this.dtRetorna.SpinButtonsVisible = true;
			this.dtRetorna.TabIndex = 1013;
			this.dtRetorna.Value = ((object)(resources.GetObject("dtRetorna.Value")));
			this.dtRetorna.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtRetorna_KeyDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(704, 40);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(28, 17);
			this.label4.TabIndex = 1014;
			this.label4.Text = "Hora";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtInicia
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtInicia.Appearance = appearance31;
			this.dtInicia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtInicia.DateButtons.Add(dateButton5);
			this.dtInicia.Enabled = false;
			this.dtInicia.Format = "HH:mm";
			this.dtInicia.Location = new System.Drawing.Point(304, 38);
			this.dtInicia.Name = "dtInicia";
			this.dtInicia.NonAutoSizeHeight = 23;
			this.dtInicia.Size = new System.Drawing.Size(88, 21);
			this.dtInicia.SpinButtonsVisible = true;
			this.dtInicia.TabIndex = 1015;
			this.dtInicia.Value = ((object)(resources.GetObject("dtInicia.Value")));
			this.dtInicia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtInicia_KeyDown);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(456, 124);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(74, 17);
			this.label15.TabIndex = 1016;
			this.label15.Text = "Lugar Destino";
			// 
			// txtOrigen
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrigen.Appearance = appearance32;
			this.txtOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigen.Enabled = false;
			this.txtOrigen.Location = new System.Drawing.Point(96, 96);
			this.txtOrigen.Multiline = true;
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.Size = new System.Drawing.Size(344, 72);
			this.txtOrigen.TabIndex = 1017;
			// 
			// optTransporte
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTransporte.Appearance = appearance33;
			this.optTransporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTransporte.CheckedIndex = 0;
			this.optTransporte.Enabled = false;
			this.optTransporte.ItemAppearance = appearance34;
			this.optTransporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Propio";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Publico";
			this.optTransporte.Items.Add(valueListItem1);
			this.optTransporte.Items.Add(valueListItem2);
			this.optTransporte.ItemSpacingVertical = 10;
			this.optTransporte.Location = new System.Drawing.Point(552, 63);
			this.optTransporte.Name = "optTransporte";
			this.optTransporte.Size = new System.Drawing.Size(136, 24);
			this.optTransporte.TabIndex = 1018;
			this.optTransporte.Text = "Propio";
			this.optTransporte.ValueChanged += new System.EventHandler(this.optTransporte_ValueChanged);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(456, 67);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(58, 17);
			this.label16.TabIndex = 1019;
			this.label16.Text = "Transporte";
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(1147, 144);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(72, 16);
			this.linkLabel1.TabIndex = 1020;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "Informe";
			this.linkLabel1.Visible = false;
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// frmGastosViajes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1323, 502);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label67);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.optTransporte);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.dtInicia);
			this.Controls.Add(this.dtRetorna);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnRevisar);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btnEmpresa);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.btnDeposito);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbCobradores);
			this.Controls.Add(this.txtidEgresoReposicion);
			this.Controls.Add(this.txtidAsiento);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtidDeposito);
			this.Controls.Add(this.txtidEgreso);
			this.Controls.Add(this.txtidPersonal);
			this.Controls.Add(this.txtidGastosviajes);
			this.Controls.Add(this.uGridDetalle);
			this.Controls.Add(this.txtLDestino);
			this.Controls.Add(this.txtValorS);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.dtFechaRetorna);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.txtCedula);
			this.Controls.Add(this.dtFechaSalida);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbCuenta);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtAcreditado);
			this.Controls.Add(this.btnAnular);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Location = new System.Drawing.Point(400, 0);
			this.Name = "frmGastosViajes";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Solicitud Gastos Viajes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGastosViajes_KeyDown);
			this.Load += new System.EventHandler(this.frmGastosViajes_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFechaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCedula)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetorna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAcreditado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorS)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidGastosviajes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidDeposito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEgresoReposicion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCobradores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtRetorna)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtInicia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private bool bActualiza = true;
		int idBloqueaTransacciones = 0;
		int Grupo = 0;
		int Cont = 0;
		int idPersonal = 0;
		decimal VSolicitado = 0.00m;
		decimal dAcreditado = 0.00m;
		string sNotasAnulacion="";
		string sCarpeta = "";
		string sUrlNube = "";
		
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		int Exportar = 0;
		private void frmGastosViajes_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0218");

			if (!Funcion.Permiso("1248", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Solicitud De Gastos De Viaje", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.BExportar) Exportar = 1;

			this.dtFechaSalida.Value = DateTime.Now.AddDays(+1);
			this.dtFechaRetorna.Value = DateTime.Now.AddDays(+2 );
			this.dtInicia.Value = DateTime.Now.AddDays(+1);
			this.dtRetorna.Value = DateTime.Now.AddDays(+2);

			

			this.cmbCobradores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ConsultaGastosViajes 1 "));
			this.cmbCentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idProyecto, Nombre From Proyecto Where Visible = 1 Order by Nombre");
			this.cmbCuenta.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("EXEC ConsultaGastosViajes 2 "));
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC ConsultaGastosViajes 3 "));
			//    
			//			string SqlUs = string.Format("Select idGrupo From Usuario Where Nombre = '{0}' And Activo = 1", MenuLatinium.stUsuario);
			//			Grupo = Funcion.iEscalarSQL(cdsSeteoF, SqlUs);
			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\DetalleGastosDeViajes2.rpt";
		}

		private void Encabezado()
		{
			this.txtCedula.Text = " ";	
			this.txtNombre.Text =  " ";
			this.dtFechaSalida.Value = System.DBNull.Value; 
			this.dtFechaRetorna.Value = System.DBNull.Value;
			this.dtInicia.Value = System.DBNull.Value; 
			this.dtRetorna.Value = System.DBNull.Value; 
			this.txtValorS.Value =  0.00m;
			this.txtLDestino.Text = " ";
			this.txtOrigen.Text = " ";
			this.optTransporte.Value = 1;
			this.txtObservaciones.Text = " ";
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.txtAcreditado.Value =  0.00m;
			this.txtTotal.Value = 0.00m;
			this.lblEstado.Text = " ";


			//			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleSuministros  0");

			this.txtCedula.Enabled = true;
			this.dtFechaSalida.Enabled = true;	
			this.dtFechaRetorna.Enabled = true;	
			this.dtInicia.Enabled = true;	
			this.dtRetorna.Enabled = true;	
			//			this.txtValorS.Enabled = true;
			this.txtLDestino.Enabled = true;
			this.txtOrigen.Enabled = true;
			this.optTransporte.Enabled = true;
			this.txtObservaciones.Enabled = true;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnEmpresa.Enabled = true;
			//			this.txtNotas.Enabled = false;	
		}

		

		private void BQEncabezado()
		{
			this.txtCedula.Enabled = false;
			this.dtFechaSalida.Enabled = false;	
			this.dtFechaRetorna.Enabled = false;	
			this.dtInicia.Enabled = false;	
			this.dtRetorna.Enabled = false;	
			//			this.txtValorS.Enabled = false;
			this.txtLDestino.Enabled = false;
			this.txtObservaciones.Enabled = false;
			this.btnEmpresa.Enabled = false;
			this.txtOrigen.Enabled = false;
			this.optTransporte.Enabled = false;

			this.btnGuardar.Enabled = false;
			this.btnConsultar.Enabled = true;
			this.btnNuevo.Enabled = true;

		}

		private void ControlesPago(bool bActiva)
		{
			if(bActiva)
			{
				this.cmbFormaPago.Value = System.DBNull.Value;
				this.txtAcreditado.Value = 0.00m;
				this.cmbCuenta.Value = System.DBNull.Value;
			
				this.cmbFormaPago.Enabled = true;
				this.btnRevisar.Enabled = true;
				this.cmbCuenta.Enabled = true;
				
			}

		}


		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void txtCedula_Validated(object sender, System.EventArgs e)
		{
			if(!bEdicion) return;
			if (!Validacion.vbTexto(this.txtCedula, 3, 10, "Cedula del Personal")) 
			{ 
				this.txtCedula.Focus();
				return;
			}

			bool bEmpleado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaEmpleado '{0}'", this.txtCedula.Text.ToString()), true);
			if (bEmpleado)
			{
				string sSQLoc = string.Format("Select (ApellidoPaterno + ' ' + ApellidoMaterno + ' ' + PrimerNombre + ' ' + SegundoNombre) as Nombre from Nomina where  NIdentificacion = '{0}' ", this.txtCedula.Value);	
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQLoc);

				string sSQPer = string.Format("Select idPersonal from Nomina where  NIdentificacion = '{0}' ", this.txtCedula.Value);	
				this.txtidPersonal.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQPer);
			}

			if(this.txtNombre.Text.Equals(""))
			{
				MessageBox.Show("Empleado  Inactivo" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCedula.Text = "";
				this.txtCedula.Focus();
			}
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			this.dtFechaSalida.CalendarInfo.MinDate = DateTime.Now.AddDays(+1);
			this.dtFechaRetorna.CalendarInfo.MinDate = DateTime.Now.AddDays(+1);
		
			//			this.dtFechaSalida.Value = DateTime.Today;	
			//			this.dtFechaRetorna.Value = DateTime.Today;		
	
			this.dtFechaSalida.Value = DateTime.Now.AddDays(+ 1);
			this.dtFechaRetorna.Value = DateTime.Now.AddDays(+ 1);
			this.dtInicia.Value = DateTime.Now.AddDays(+ 1);
			this.dtRetorna.Value = DateTime.Now.AddDays(+ 1);
			this.lblEstado.Text = "PENDIENTE";

			#region controles
			this.txtCedula.Enabled = true;
			this.dtFechaSalida.Enabled = true;	
			this.dtFechaRetorna.Enabled = true;	
			this.dtInicia.Enabled = true;	
			this.dtRetorna.Enabled = true;	
			//			this.txtValorS.Enabled = true;
			this.txtLDestino.Enabled = true;
			this.txtOrigen.Enabled = true;
			this.optTransporte.Enabled = true;
			this.btnEmpresa.Enabled = true;
			this.txtObservaciones.Enabled = false;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;	
			#endregion controles
			
			bNuevo = true;
			bEdicion = true;

			this.txtCedula.Focus();			
		}

		private void txtCedula_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if(!bEdicion) return;
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) 
			{
				e.Cancel = true;    
			}
		}

		private void txtCedula_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaSalida.Focus();
		
		}

		private void txtCedula_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtCedula.Focus();
		}

		private void dtFechaSalida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtInicia.Focus();
		}

		private void dtFechaSalida_ValueChanged(object sender, System.EventArgs e)
		{
			this.dtFechaRetorna.Focus();
		}

		private void dtFechaRetorna_ValueChanged(object sender, System.EventArgs e)
		{
			this.txtValorS.Focus();
			this.txtValorS.SelectAll();
		}
		string txtDirectorioarchivo ="";

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{			
			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaSalida, "Seleccione la Fecha De Inicio", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaRetorna, "Seleccione la Fecha De Retorno", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtInicia, "Seleccione la Hora De Inicio", false, "", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtRetorna, "Seleccione la Hora De Retorno", false, "", cdsSeteoF)) return;
			//			if ((DateTime) this.dtFechaSalida.Value == DateTime.Today )
			//			{
			//				MessageBox.Show("La Fecha De Salida no puede ser la Actual ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				return;
			//
			//			}
			if (!Validacion.vbIdentificacion(this.txtCedula, 5, cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtFechaSalida.Value), Convert.ToDateTime(this.dtFechaRetorna.Value), this.dtFechaSalida, "La fecha de salida no puede ser posterior a la fecha de retorna", false)) return;
			if (!Validacion.vbTexto(this.txtLDestino, 10, 100, "Destino")) return;
			if (!Validacion.vbTexto(this.txtOrigen, 10, 100, "Origen")) return;

			if (!Validacion.vbCampoDecimalVacio (this.txtValorS,"Ingrese el Valor a solicitar",1,10000)) return;			
			#endregion Validacion
			
			if (bNuevo) this.txtEstado.Value = 1;

			if((int)this.txtEstado.Value == 1 && !bNuevo)
			{
				if (miAcceso.CambiarFecha)
				{				
					//					if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione la Forma de Pago")) return;
					//					if (!Validacion.vbComboVacio(this.cmbCuenta, "Seleccione la Cuenta")) return;
					//					if (!Validacion.vbCampoDecimalVacio (this.txtAcreditado,"Ingrese el Valor Acreditar",1,10000)) return;
					//					if (!Validacion.vbTexto(this.txtObservaciones, 10, 100, "Observaciones")) return;				
					//
					//					this.txtEstado.Value = 2;
				}
			}
			
			if((int)this.txtEstado.Value == 4 || (int)this.txtEstado.Value == 5)
			{
				#region Valida Grid
				if ((int) this.uGridDetalle.Rows.Count == 0)
				{
					MessageBox.Show("Ingrese El Detalle De Los Gastos Realizados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridDetalle.Focus ();
					return;
				}

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				{	
					#region Ingreso Con Factura
					if (!(bool)dr.Cells["SF"].Value )
					{
						#region Valida Ingreso de RUC
						if (dr.Cells["Ruc"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese el RUC del proveedor.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}
						#endregion Valida Ingreso de RUC

						#region Valida identificación
						int idTipoRUC = 1;
						if (dr.Cells["Ruc"].Value.ToString().Length == 10) idTipoRUC = 2;
						string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", dr.Cells["Ruc"].Value.ToString(), idTipoRUC);
						string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
						if (VlCed.StartsWith("Error"))
						{
							MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}
						#endregion Valida identificación

						#region Busca Proveedor
						if (dr.Cells["Ruc"].Value.ToString().Length > 0)
						{
							string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", dr.Cells["Ruc"].Value.ToString(), 4, 0, false);
							SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
							drC.Read();
							if (!drC.HasRows)
							{
								drC.Close();
								MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.uGridDetalle.ActiveRow = dr;
								this.uGridDetalle.ActiveRow.Selected = true;
								return;
							}
							drC.Close();
						}
						#endregion Busca Proveedor

						#region Autorización
						if (dr.Cells["Autorizacion"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese la Autorización.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["Autorizacion"].Value.ToString().Length != 10)
						{
							if (dr.Cells["Autorizacion"].Value.ToString().Length != 37)
							{
								MessageBox.Show("La Autorización es de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
								this.uGridDetalle.ActiveRow = dr;
								this.uGridDetalle.ActiveRow.Selected = true;
								return;
							}
						}
						#endregion Autorización

						#region Serie
						if (dr.Cells["Serie"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese la Serie.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}
						if (dr.Cells["Serie"].Value.ToString().Length < 6)
						{
							MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}
						#endregion Serie

						#region Numero
						if (dr.Cells["Numero"].Value.ToString().Length == 0)
						{
							MessageBox.Show("Ingrese el número de la Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}

						char Pad = '0';
						dr.Cells["Numero"].Value = dr.Cells["Numero"].Value.ToString().PadLeft(9, Pad);
					
						int idGastosViaje= 0;
						int idDetalleGastosViajes = 0;

						if (dr.Cells["idDetalleGViaje"].Value != System.DBNull.Value) idDetalleGastosViajes = (int)dr.Cells["idDetalleGViaje"].Value; 
						if (dr.Cells["idGastosViaje"].Value != System.DBNull.Value) idGastosViaje = (int)dr.Cells["idGastosViaje"].Value; 

						string sSQLValNum = string.Format("Exec ValidaNumeracionCompraGastosDeViaje {0}, {1}, {2}, '{3}', '{4}', '{5}'", 
							idDetalleGastosViajes, idGastosViaje, 
							(int)dr.Cells["idProveedor"].Value, 
							dr.Cells["Serie"].Value.ToString(), 
							dr.Cells["Numero"].Value.ToString(), 
							dr.Cells["Autorizacion"].Value.ToString());
						string sMensValNum = Funcion.sEscalarSQL(cdsSeteoF, sSQLValNum, true);
						if (sMensValNum.Length > 0)
						{
							MessageBox.Show(sMensValNum, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							this.uGridDetalle.ActiveRow = dr;
							this.uGridDetalle.ActiveRow.Selected = true;
							return;
						}
						#endregion Numero
					}
					#endregion Ingreso Con Factura

					if (dr.Cells["idCuenta"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione la cuenta del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}		
				
					if ((int)dr.Cells["idCuenta"].Value == 0)
					{
						MessageBox.Show("Seleccione la cuenta del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}	
					if (dr.Cells["Fecha"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (Convert.ToDateTime(dr.Cells["Fecha"].Value) < Convert.ToDateTime(this.dtFechaSalida.Value))
					{
						MessageBox.Show("La Fecha del Detalle no puede ser menor a la FECHA DE SALIDA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (Convert.ToDateTime(dr.Cells["Fecha"].Value) > Convert.ToDateTime(this.dtFechaRetorna.Value))
					{
						MessageBox.Show("La Fecha del Detalle no puede ser mayor a la FECHA DE RETORNO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}	
					if (dr.Cells["idProyecto"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione el Centro De Costo del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}		
					if ((int)dr.Cells["idProyecto"].Value == 0)
					{
						MessageBox.Show("Seleccione el Centro De Costo del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}	
					if (dr.Cells["Detalle"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Escriba el Detalle ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Detalle"].Value.ToString().Trim().Length < 10 )
					{
						MessageBox.Show("El Detalle debe tener minimo 10 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Detalle"].Value.ToString().Trim().Length > 100 )
					{
						MessageBox.Show("El Detalle debe tener maximo 100 caracteres ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (dr.Cells["Valor"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese el Valor del Gasto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}
					if (Convert.ToDecimal(dr.Cells["Valor"].Value) == 0.00m)
					{
						MessageBox.Show("El monto del Gasto no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridDetalle.ActiveRow = dr;
						this.uGridDetalle.ActiveRow.Selected = true;
						return;
					}					
				}
				#endregion Valida Grid

				if ((int)this.txtEstado.Value == 4)				
					if (DialogResult.Yes == MessageBox.Show("¿Desea cerrar los gastos?, Si presiona SI ya no podra modificar o registrar mas gastos.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
					{
						try
						{
							MessageBox.Show("Acontinuación seleccione el Informe de Actividades en formato PDF.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

							OpenFileDialog buscar = new OpenFileDialog ();
							buscar.Filter = "Text files (*.pdf)|*.pdf";
							if (buscar.ShowDialog () == DialogResult.OK)
							{
								txtDirectorioarchivo = buscar.FileName;
								//txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",txtDirectorioarchivo,@"\"));
								this.txtEstado.Value = 5;
							}
						}
						catch (Exception ex)
						{
							MessageBox.Show ("Error: "+ ex.Message);
						}
					}
			}

			decimal dSaldo = 0.00m;
			decimal dS = 0.00m;
			this.txtTotal.Value = 0;

			if ((int)this.txtEstado.Value == 5)
			{ 
				if (miAcceso.BAPAutorizar)
					if (DialogResult.Yes == MessageBox.Show("¿Desea generar la liquidacion de los gastos?, Si presiona SI ya no podra modificar o registrar mas gastos.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
						this.txtEstado.Value = 6;
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
					this.txtTotal.Value = Convert.ToDecimal(dr.Cells["Valor"].Value) + Convert.ToDecimal(this.txtTotal.Value);

				dSaldo = Convert.ToDecimal(this.txtAcreditado.Value) - Convert.ToDecimal(this.txtTotal.Value);
			}

			this.txtTotal.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				this.txtTotal.Value = Convert.ToDecimal(dr.Cells["Valor"].Value) + Convert.ToDecimal(this.txtTotal.Value);

			dS = Convert.ToDecimal(this.txtAcreditado.Value) - Convert.ToDecimal(this.txtTotal.Value);
				
			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
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

					#region Numeracion
					if (bNuevo)
					{
						string ssSQL = string.Format("Exec NumeracionLocales 66, 0, 0");
						oCmdActualiza.CommandText = ssSQL;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();						
					}
					#endregion Numeracion

					#region Graba Maestro
					int idFormaDePago = 0;
					int idCuenta = 0;
					if (this.cmbCuenta.ActiveRow != null) idCuenta = (int) this.cmbCuenta.Value;
					if(this.cmbFormaPago.ActiveRow != null) idFormaDePago = (int) this.cmbFormaPago.Value;

					string  sSQL1 = string.Format("Exec GrabaSocilitudGastosViaje {0}, '{1}', {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', {21}" ,
						(int)this.txtidGastosviajes.Value, 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
						(int)this.txtidPersonal.Value,
						Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd"), 
						Convert.ToDateTime(this.dtFechaRetorna.Value).ToString("yyyyMMdd"),
						this.txtLDestino.Text.ToString(), 
						this.txtObservaciones.Text.ToString(), 
						Convert.ToDecimal (this.txtValorS.Value), 
						Convert.ToDecimal (this.txtAcreditado.Value), 
						idFormaDePago, 
						idCuenta,
						(int)this.txtidEgreso.Value, 
						(int)this.txtidDeposito.Value, 
						(int)this.txtEstado.Value,
						(int) this.txtidAsiento.Value,
						(int) this.txtidEgresoReposicion.Value,
						sNotasAnulacion,
						this.txtNumero.Text,
						Convert.ToDateTime(this.dtInicia.Value).ToString("yyyyMMdd HH:mm:ss"),
						Convert.ToDateTime(this.dtRetorna.Value).ToString("yyyyMMdd HH:mm:ss"),
						Convert.ToString( this.txtOrigen.Text),
						(int) this.optTransporte.Value
						);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidGastosviajes.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					#region Graba Ingreso Gastos
					if((int) this.ultraGrid1.Rows.Count > 0)
					{						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{
							string  sSQLDeta = string.Format("Exec GuardaDettaleIngresoReembolsoViaticos {0}, {1}, '{2}', {3}, {4}, {5}, {6}, {7}, {8},{9}" ,
								Convert.ToInt32 ( dr.Cells["idIngresoReembolsoViaticos"].Value), 
								Convert.ToInt32 ( dr.Cells["idReembolsoViaticos"].Value), 
								Convert.ToString  ( dr.Cells["Concepto"].Value),
								Convert.ToDecimal ( dr.Cells["AdministrativoSolicitado"].Value),
								Convert.ToDecimal ( dr.Cells["GerentesSolicitado"].Value)
								,Convert.ToDecimal ( dr.Cells["Administrativo"].Value)
								,Convert.ToDecimal ( dr.Cells["Gerentes"].Value)
								, (int)this.txtidGastosviajes.Value,
								Convert.ToDecimal  ( dr.Cells["SubTotal"].Value),
								Convert.ToDecimal ( dr.Cells["SubTotalAcreditado"].Value));
							//							Funcion.EjecutaSQL(cdsSeteoF, sSQLDeta);

							oCmdActualiza.CommandText = sSQLDeta;
							oCmdActualiza.ExecuteNonQuery();	
						}						
					}
					#endregion Graba Ingreso Gastos

					#region Graba DetalleMaestro
					if((int) this.uGridDetalle.Rows.Count > 0)
					{						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
						{
							string  sSQLDet = string.Format("Exec GuardaDetalleGastosViaje {0}, {1}, '{2}', {3}, '{4}', '{5}', {6}, {7}, '{8}','{9}', '{10}', '{11}', '{12}', {13}" ,
								Convert.ToInt32 ( dr.Cells["idDetalleGViaje"].Value), 
								(int) this.txtidGastosviajes.Value, 
								Convert.ToDateTime(dr.Cells["Fecha"].Value).ToString("yyyyMMdd"),
								Convert.ToInt32 ( dr.Cells["idCuenta"].Value),
								dr.Cells["Detalle"].Value.ToString(),
								Convert.ToDecimal ( dr.Cells["Valor"].Value),
								Convert.ToInt32 ( dr.Cells["idProyecto"].Value),
								Convert.ToInt32 ( dr.Cells["idProveedor"].Value),
								dr.Cells["Ruc"].Value,
								dr.Cells["RazonSocial"].Value,
								dr.Cells["Serie"].Value,
								dr.Cells["Numero"].Value,
								dr.Cells["Autorizacion"].Value,
								Convert.ToBoolean ( dr.Cells["SF"].Value)
								);
							oCmdActualiza.CommandText = sSQLDet;
							oCmdActualiza.ExecuteNonQuery();	
						}						
					}
					#endregion Graba DetalleMaestro
					
					#region Autorizacion de pago
					if ((int)this.txtEstado.Value == 2)
					{
						string sSQLAut = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, {5}, '1 DE 1', '{6}', {7}, 0", 
							10, (int)this.txtidGastosviajes.Value, 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
							Convert.ToDecimal(this.txtAcreditado.Value), 7, this.txtObservaciones.Text.ToString(), (int)this.cmbCuenta.Value);			
						oCmdActualiza.CommandText = sSQLAut;
						oCmdActualiza.ExecuteNonQuery();		
					}
					#endregion Autorizacion de pago

					#region Asiento Liquidacion
					if ((int)this.txtEstado.Value == 6)
					{
						string sSQLAsi = string.Format("Exec AsientoLiquidacionGastosViaje {0}, {1}, {2} ", 
							(int)this.txtidGastosviajes.Value, (int) this.txtidAsiento.Value, dS );			
						oCmdActualiza.CommandText = sSQLAsi;
						oCmdActualiza.ExecuteNonQuery();

						#region Generar Autorizacion de Reposicion
						if (dS < 0.00m)
						{		
							dS = 0.00m;
							this.txtTotal.Value = 0.00m;

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
								this.txtTotal.Value = Convert.ToDecimal(dr.Cells["Valor"].Value) + Convert.ToDecimal(this.txtTotal.Value);

							dS = Convert.ToDecimal(this.txtAcreditado.Value) - Convert.ToDecimal(this.txtTotal.Value);

							string sSQLAut = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, {5}, '1 DE 1', '{6}', {7}, 0", 
								21, (int)this.txtidGastosviajes.Value, 
								Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), 
								Math.Abs(dS), 7, this.txtObservaciones.Text.ToString(), 12);			
							oCmdActualiza.CommandText = sSQLAut;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLAsiEs = string.Format("update GastosDeViaje set Estado = 7 where idGastosViaje = {0}", 
								(int)this.txtidGastosviajes.Value);			
							oCmdActualiza.CommandText = sSQLAsiEs;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Generar Autorizacion de Reposicion

						
					}
					#endregion Asiento Liquidacion

					#region nube 
					if(txtDirectorioarchivo.Length > 10 )
					{
						NuevoGoogle (txtDirectorioarchivo);
					}
					if (sUrlNube.Length > 10 )
					{
						string sSQLGoogle = string.Format("Exec GrabaUrlNubeGastosViaje {0}, '{1}', '{2}'", (int)this.txtidGastosviajes.Value, sCarpeta, sUrlNube );
						oCmdActualiza.CommandText = sSQLGoogle;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion nube 
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
					this.txtCedula.Enabled = false;
					this.dtFechaSalida.Enabled = false;	
					this.dtFechaRetorna.Enabled = false;
					this.dtInicia.Enabled = false;	
					this.dtRetorna.Enabled = false;	
					this.txtValorS.Enabled = false;
					this.txtLDestino.Enabled = false;
					this.txtOrigen.Enabled = false;
					this.optTransporte.Enabled = false;
					this.btnEmpresa.Enabled = false;
					this.txtObservaciones.Enabled = false;

					this.cmbFormaPago.Enabled = false;
					this.cmbCuenta.Enabled = false;
					this.btnRevisar.Enabled = false;				

					FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);

					if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = false;

					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.Consulta((int)this.txtidGastosviajes.Value);

					if (sUrlNube.Length > 10 )
					{
						this.linkLabel1.Visible = true;
					}
					bNuevo = false;
					bEdicion = false;
					#endregion Controles			
				}
				catch(Exception ex)
				{
					try
					{ 
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
					}
				}
				finally
				{
					oConexion.Close();
				}
			}
				
			#endregion Guardar
		}

		public void  NuevoGoogle (string sDirectorio)
		{
			string carpeta_1="";
			string carpeta_2="";
			string carpeta_3="";
			string carpeta ="";
			//string txtDirectorioarchivo ="";
			string txtURL ="";
			string SArhivoNuevo =""; 
			Random random = new Random();
			int randomNumber = random.Next(1, 10001);	

			#region Editar google 
			carpeta_1 = @"\\192.168.1.10\Proyectos\Gastos_De_Viaje";
			carpeta_2 = @"\\192.168.1.10\Proyectos\Gastos_De_Viaje\"+MenuLatinium.stDirFacturacion;

			carpeta = @"\\192.168.1.10\Proyectos\Gastos_De_Viaje\"+MenuLatinium.stDirFacturacion+@"\" + this.txtNumero.Text.Replace(" ", "_") +"_"+this.txtNumero.Text+"_"+ Convert.ToString (this.txtidGastosviajes.Value);
			carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + this.txtCedula.Text +"_"+this.txtNumero.Text+"_"+ Convert.ToString (this.txtidGastosviajes.Value);
			SArhivoNuevo =  Convert.ToString (this.txtidGastosviajes.Value)+"_"+Convert.ToString(randomNumber);
			try
			{
				Directory.GetDirectories(@"\\192.168.1.10\Proyectos");
				Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
				if (!Directory.Exists(carpeta_1))
				{
					Directory.CreateDirectory(carpeta);
					
				}
				if (!Directory.Exists(carpeta_2))
				{
					Directory.CreateDirectory(carpeta);
					
				}

				if (!Directory.Exists(carpeta))
				{
					Directory.CreateDirectory(carpeta);
					
				}
				

				//	Directory.CreateDirectory(carpeta);
				txtDirectorioarchivo = sDirectorio;
				FileInfo fi= new FileInfo(txtDirectorioarchivo);
				string nombre = fi.Name;//nombre.pdf
				string Directorio = txtDirectorioarchivo.Replace(nombre, "")+SArhivoNuevo+".pdf";//archivo nuevo
				File.Delete(Directorio); // First delete the new file if it exists, because it wont be overwritten by Move
				File.Move(txtDirectorioarchivo, Directorio);


			
				txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",Directorio,@"\"));
					
				carpeta  = (carpeta + @"\" + txtURL.Replace(" ", "_"));
          
				System.IO.File.Copy (  Directorio ,carpeta, false);
				sCarpeta = carpeta;
				sUrlNube = carpeta_3.Replace(" ", "_")+ @"\" + txtURL.Replace(" ", "_");
				MessageBox.Show(string.Format(" Archivo '{0}' subido . ", txtURL.Replace(" ", "_")), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
										
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);

			}
			#endregion Editar google 
		}
		
		private void txtLDestino_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.txtObservaciones.Focus();
		}

		private void uGridDetalle_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetalleGViaje"].Value = 0;
				e.Row.Cells["idGastosViaje"].Value = 0;
				e.Row.Cells["idCuenta"].Value = 0;
				e.Row.Cells["Valor"].Value = 0.00 ;	
				e.Row.Cells["Fecha"].Value = DateTime.Today  ;	
				e.Row.Cells["SF"].Value = false;

				e.Row.Cells["idProveedor"].Value = 0;
				e.Row.Cells["Ruc"].Value = "";
				e.Row.Cells["RazonSocial"].Value = "";
				e.Row.Cells["Serie"].Value = "";	
				e.Row.Cells["Numero"].Value = "";
				e.Row.Cells["Autorizacion"].Value = "";
			}



			

		}

		private void uGridDetalle_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
			if (e.Cell.Column.ToString() == "Ruc")
			{
				#region Busca Proveedor
				if (e.Cell.Row.Cells["Ruc"].Value.ToString().Length > 0)
				{
					string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", e.Cell.Row.Cells["Ruc"].Value.ToString(), 4, 0, false);
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
					drC.Read();
					if (drC.HasRows)
					{
						e.Cell.Row.Cells["idProveedor"].Value = (int) drC.GetValue(0);
						e.Cell.Row.Cells["RazonSocial"].Value = drC.GetString(2);	
			
						drC.Close();

						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaGrupoProveedor '{0}'", e.Cell.Row.Cells["Ruc"].Value.ToString())) == 0)
						{
							MessageBox.Show("En la Ficha Proveedores Seleccione el Grupo al cual pertenece el Proveedor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cell.Row.Cells["idProveedor"].Value = System.DBNull.Value;
							e.Cell.Row.Cells["Ruc"].Value = "";
							e.Cell.Row.Cells["RazonSocial"].Value = "";
							return;
						}
					}
					else
					{
						drC.Close();
						MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["Ruc"].Value = "";
						return;
					}
					drC.Close();
				}
				#endregion Busca Proveedor
			}

			#region Autorización
			if (e.Cell.Column.ToString() == "Autorizacion")
			{
				if (e.Cell.Row.Cells["Autorizacion"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["Autorizacion"].Value.ToString().Length != 10)
					{
						if (e.Cell.Row.Cells["Autorizacion"].Value.ToString().Length != 37)
						{
							MessageBox.Show("La Autorización debe ser de 10 o 37 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
							e.Cell.Row.Cells["Autorizacion"].Value = "";
							return;
						}
					}
				}
			}
			#endregion Autorización

			#region Serie
			if (e.Cell.Column.ToString() == "Serie")
			{
				if (e.Cell.Row.Cells["Serie"].Value.ToString().Length > 0)
				{
					if (e.Cell.Row.Cells["Serie"].Value.ToString().Length < 6)
					{
						MessageBox.Show("La Serie debe ser de 6 dígitos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["Serie"].Value = "";
						return;
					}
				}
			}
			#endregion Serie

			#region Juan 
			//			if(bEdicion)
			//			{
			//				#region Codigo
			//				if (e.Cell.Column.ToString() == "Codigo")
			//				{
			//					#region Busca Cuenta
			//					string sSQLContador = string.Format("Exec [AsientoBuscaCTAGastos] '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
			//					int iContador = Funcion.iEscalarSQL(cdsSeteoF, sSQLContador);
			//
			//					if (iContador == 1)
			//					{
			//						string sSQLInfo = string.Format("Exec [AsientoInfoCTAGastos] '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
			//						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLInfo, true);
			//						dr.Read();
			//						if (dr.HasRows)
			//						{
			//							e.Cell.Row.Cells["idCuenta"].Value = dr.GetInt32(0);
			//							e.Cell.Row.Cells["Detalle"].Value = dr.GetString(2);
			//						}
			//						dr.Close();
			//
			//					}
			//					else
			//					{
			//						using (BuscaCuentas Busqueda = new BuscaCuentas(e.Cell.Row.Cells["Codigo"].Value.ToString(), 5))
			//						{
			//							if (DialogResult.OK == Busqueda.ShowDialog())
			//							{
			//								e.Cell.Row.Cells["Codigo"].Value = Busqueda.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
			//							}
			//						}
			//					}
			//					#endregion Busca Cuenta	
			//				}
			//				#endregion Codigo}
			//
			//				#region Presupuesto
			//				if (e.Cell.Column.ToString() == "Total")
			//				{
			//
			//					if (e.Cell.Row.Cells["idCuenta"].Value == System.DBNull.Value)
			//					{	
			//						if (e.Cell.Row.Cells["Total"].Value != System.DBNull.Value)
			//						{	
			//							if ((decimal)e.Cell.Row.Cells["Total"].Value > 0.00m)
			//							{
			//								e.Cell.Value = 0.00m;
			//								MessageBox.Show("Seleccione una Cuenta antes de ingresar el valor del Debe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//							}
			//						}
			//					}
			//
			//					TotalPresupuesto();
			//				}
			//				#endregion Presupuesto
			//      
			//				//				}
			//			}
			#endregion Juan 
		}
		
		private void uGridDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos
			sNotasAnulacion ="";

			this.txtidGastosviajes.Value = 0;
			this.txtidEgreso.Value = 0;
			this.txtidAsiento.Value = 0;
			this.txtidEgresoReposicion.Value = 0;
			this.txtidDeposito.Value = 0;

			this.txtEstado.Value = 0;
			this.lblEstado.Text = "";

			this.txtidPersonal.Value = 0;			
			this.txtCedula.Text = "";	
			this.txtNombre.Text =  "";
			this.txtNumero.Text ="";
			//			this.dtFecha.Value = System.DBNull.Value; 
			this.dtFechaSalida.Value = System.DBNull.Value; 
			this.dtFechaRetorna.Value = System.DBNull.Value; 
			this.dtInicia.Value = System.DBNull.Value; 
			this.dtRetorna.Value = System.DBNull.Value; 
			this.txtValorS.Value =  0.00m;
			this.txtLDestino.Text = "";
			this.txtOrigen.Text = "";
			this.optTransporte.Value = 1;
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbCuenta.Value = System.DBNull.Value;
			this.txtAcreditado.Value =  0.00m;			
			this.txtObservaciones.Text = "";
						
			this.txtTotal.Value = 0.00m;

			this.uGridDetalle.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaGeneralDetalleSocilitudGastosViaje  0");
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleIngresoReembolsoViaticos  0,0,0");
			#endregion Vacia Campos

			#region Bloquea Controles
			this.txtCedula.Enabled = false;
			this.dtFechaSalida.Enabled = false;	
			this.dtFechaRetorna.Enabled = false;
			this.dtInicia.Enabled = false;
			this.dtRetorna.Enabled = false;
			this.txtValorS.Enabled = false;
			this.txtLDestino.Enabled = false;
			this.txtOrigen.Enabled = false;
			this.optTransporte.Enabled = false;
			this.txtObservaciones.Enabled = false;
			this.btnEmpresa.Enabled = false;

			this.cmbFormaPago.Enabled = false;
			this.cmbCuenta.Enabled = false;
			this.btnRevisar.Enabled = false;
			

			FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);

			#endregion Bloquea Controles

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;
			txtDirectorioarchivo ="";
			sCarpeta = "";
			sUrlNube = "";
			this.linkLabel1.Visible = false;

						
		
		}

		private void uGridDetalle_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (uGridDetalle.ActiveCell == null) return;
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.uGridDetalle);
			try
			{				
				switch(e.KeyValue)
				{	
						#region Proveedores
					case (int) Keys.F3:
						if (this.uGridDetalle.ActiveCell.Column.Key.ToString() == "Ruc" && !(bool)this.uGridDetalle.ActiveRow.Cells["SF"].Value )
						{
							using (BuscaClientes miBusqueda = new BuscaClientes(this.uGridDetalle.ActiveRow.Cells["Ruc"].Value.ToString(), 4, 0, 1, 0, false))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{	
									this.uGridDetalle.ActiveRow.Cells["idProveedor"].Value = miBusqueda.idCliente;
									this.uGridDetalle.ActiveRow.Cells["Ruc"].Value = miBusqueda.Ruc;
								}
							}
						}
						
						break;
						#endregion Proveedores
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology");
			}
		}

		private void txtValorS_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtAcreditado_ValueChanged(object sender, System.EventArgs e)
		{
			//			dAcreditado = Convert.ToDecimal(this.txtAcreditado.Value);
			//			decimal Result = Math.Round((VSolicitado - dAcreditado ),2);
			//
			//			if(Result > 0.00m )
			//			{
			//				this.lblSobrante.Text = "Faltante";
			//				this.lblSobrante.Text += " " + Result.ToString("#,##0.00");
			//			}
			//			else if(Result < 0.00m)  
			//			{
			//				this.lblSobrante.Text = "Sobrante";
			//				this.lblSobrante.Text += "  " + Result.ToString("#,##0.00");
			//			}
			//			else
			//			{
			//				this.lblSobrante.Text = "";
			//			}
		}

		private void dtFechaRetorna_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtRetorna.Focus();
		}

		private void txtValorS_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtLDestino.Focus();
		}

		private void txtLDestino_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaGeneralSolicitudViajes Busqueda = new frmConsultaGeneralSolicitudViajes(Exportar))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridCreditos.ActiveRow.Cells["idGastosViaje"].Value);
				}
			}
		}

		private void Consulta (int Id)
		{			
			string sSQL = string.Format("Exec ConsultaIndividualSocilitudGastosViaje {0}", Id);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidGastosviajes.Value = Id;
				if (dr.GetValue(1) != System.DBNull.Value) this.dtFecha.Value = dr.GetDateTime(1);
				if (dr.GetValue(2) != System.DBNull.Value) this.txtidPersonal.Value = dr.GetInt32(2);
				if (dr.GetValue(3) != System.DBNull.Value) this.txtNombre.Text = dr.GetString(3); 
				if (dr.GetValue(4) != System.DBNull.Value) this.txtCedula.Text= dr.GetString(4); 
				if (dr.GetValue(5) != System.DBNull.Value) this.dtFechaSalida.Value = dr.GetDateTime(5);
				if (dr.GetValue(6) != System.DBNull.Value) this.dtFechaRetorna.Value = dr.GetDateTime(6);
				if (dr.GetValue(7) != System.DBNull.Value) this.txtLDestino.Text = dr.GetString(7); 
				if (dr.GetValue(8) != System.DBNull.Value) this.txtObservaciones.Text = dr.GetString(8); 
				if (dr.GetValue(9) != System.DBNull.Value) this.txtValorS.Value = dr.GetDecimal(9);
				if (dr.GetValue(10) != System.DBNull.Value) this.txtAcreditado.Value = dr.GetDecimal(10);
				if (dr.GetValue(11) != System.DBNull.Value) this.cmbFormaPago.Value = dr.GetInt32(11);
				if (dr.GetValue(12) != System.DBNull.Value) this.cmbCuenta.Value  = dr.GetInt32(12);
				if (dr.GetValue(13) != System.DBNull.Value) this.txtidEgreso.Value = dr.GetInt32(13);
				if (dr.GetValue(14) != System.DBNull.Value) this.txtidDeposito.Value = dr.GetInt32(14);
				if (dr.GetValue(15) != System.DBNull.Value) this.txtEstado.Value = dr.GetInt32(15);
				if (dr.GetValue(16) != System.DBNull.Value) this.txtidAsiento.Value = dr.GetInt32(16);
				if (dr.GetValue(17) != System.DBNull.Value) this.txtidEgresoReposicion.Value = dr.GetInt32(17);
				if (dr.GetValue(18) != System.DBNull.Value) this.txtNumero.Text = dr.GetString(18); 
				if (dr.GetValue(19) != System.DBNull.Value) this.dtInicia.Value = dr.GetDateTime(19);
				if (dr.GetValue(20) != System.DBNull.Value) this.dtRetorna.Value = dr.GetDateTime(20);
				if (dr.GetValue(21) != System.DBNull.Value) this.optTransporte.Value = dr.GetInt32(21);
				if (dr.GetValue(22) != System.DBNull.Value) this.txtOrigen.Value = dr.GetString(22);
				if (dr.GetValue(23) != System.DBNull.Value) sCarpeta = dr.GetString(23).ToString();
				if (dr.GetValue(24) != System.DBNull.Value) sUrlNube = dr.GetString(24).ToString();

				if (sUrlNube.Length > 10)
				{
					this.linkLabel1.Visible = true;					
				}
			}
			dr.Close();

			int idDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" SELECT DATEDIFF(DAY, '{0}', '{1}') +1", Convert.ToDateTime(this.dtFechaSalida.Value).ToString("yyyyMMdd"),Convert.ToDateTime(this.dtFechaRetorna.Value).ToString("yyyyMMdd") ));		
			int iGerAdm = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" EXEC RetornaGerAdmin {0}  ", (int) this.txtidPersonal.Value));		

			string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", (int)this.txtidGastosviajes.Value, idDias, iGerAdm);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultraGrid1);

			string sSQLPar = string.Format("Exec ConsultaGeneralDetalleSocilitudGastosViaje {0}", (int)this.txtidGastosviajes.Value);
			FuncionesProcedimientos.dsGeneral(sSQLPar, this.uGridDetalle);
			
			if( (int) this.uGridDetalle.Rows.Count > 0)
			{
				this.txtTotal.Value = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drg in this.uGridDetalle.Rows.All)
				{
					this.txtTotal.Value = Convert.ToDecimal(drg.Cells["Valor"].Value) +  Convert.ToDecimal(this.txtTotal.Value) ;
				}
			}

			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
				if ((int)this.txtEstado.Value == 1 || (int)this.txtEstado.Value == 4 || (int)this.txtEstado.Value == 5) 
					this.btnEditar.Enabled = true;

			if ((int)this.txtEstado.Value == 1 || (int)this.txtEstado.Value == 2) // PENDIENTE - AUTORIZADO
			{
				if (miAcceso.Anular)this.btnAnular.Enabled = true;
			}

			if((int) this.txtEstado.Value == 6)
			{
				this.btnDeposito.Enabled = true;
			}
			if((int) this.txtEstado.Value > 2)
			{
				this.btnAnular.Enabled = false;
			}
			if((int)this.txtEstado.Value >= 6) this.btnImprimir.Visible = true;
			if((int)this.txtEstado.Value <= 5) this.btnImprimir.Visible = false;

			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "POR AUTORIZAR";
			if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "AUTORIZADO";
			if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "RECHAZADO";
			if ((int)this.txtEstado.Value == 4) this.lblEstado.Text = "EGRESO GENERADO";
			if ((int)this.txtEstado.Value == 5) this.lblEstado.Text = "GASTOS REGISTRADOS";
			if ((int)this.txtEstado.Value == 6) this.lblEstado.Text = "LIQUIDACION GENERADA";
			if ((int)this.txtEstado.Value == 7) this.lblEstado.Text = "VERIFICADO";

			this.btnEmpresa.Enabled = true;
		}

	
		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consulta ((int)this.txtidGastosviajes.Value);
			
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 1)
			{
				#region Habilta Autorizacion
				if (miAcceso.CambiarFecha)
				{
					this.cmbFormaPago.Enabled = true;
					this.cmbCuenta.Enabled = true;
					this.btnRevisar.Enabled = true;
					this.txtAcreditado.Value = 0;
					this.txtObservaciones.Enabled = true;
					this.cmbFormaPago.Focus();

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						dr.Cells["Administrativo"].Value = Convert.ToDecimal( dr.Cells["GerentesSolicitado"].Value)+ Convert.ToDecimal( dr.Cells["AdministrativoSolicitado"].Value);
						dr.Cells["SubTotalAcreditado"].Value = Convert.ToDecimal( dr.Cells["SubTotal"].Value);
					}
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						this.txtAcreditado.Value = Convert.ToDecimal(this.txtAcreditado.Value)+ Convert.ToDecimal( dr.Cells["SubTotal"].Value);
					}
				}
				else
				{					
					this.dtFechaSalida.Enabled = true;
					this.dtFechaRetorna.Enabled = true;
					this.dtInicia.Enabled = true;
					this.dtRetorna.Enabled = true;
					this.txtOrigen.Enabled = true;
					this.optTransporte.Enabled = true;
					this.txtValorS.Enabled = false;					
					this.txtLDestino.Enabled = true;
					this.btnEmpresa.Enabled = true;

					this.dtFechaSalida.Focus();
				}
				#endregion habilta Autorizacion
			}
			
			if ((int)this.txtEstado.Value == 4 || (int)this.txtEstado.Value == 5) // REGISTRO DE GASTOS - VERIFICACION DE GASTOS
			{
				FuncionesProcedimientos.EstadoGrids(true, this.uGridDetalle);
			}
      
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
			this.btnAnular.Enabled = false;


			bNuevo = false;
			bEdicion = true;
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCuenta.Focus();
		}

		private void txtAcreditado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				
		}

		private void btnRechazar_Click(object sender, System.EventArgs e)
		{
			this.Consulta ((int)this.txtidGastosviajes.Value);

			if((int) this.txtEstado.Value == 3)
			{
				MessageBox.Show("El documento ya esta RECHAZADO", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if((int) this.txtEstado.Value > 3)
			{
				MessageBox.Show("No puede Anular el Documento", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			using (frmNotas Busqueda = new frmNotas())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					sNotasAnulacion = Busqueda.txtNotas.Text;
					this.txtEstado.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec CambiaEstadoGestionesViajes {0}, 3 ", (int) this.txtidGastosviajes.Value));
					this.btnGuardar_Click (sender, e);
					this.btnAnular.Enabled = false;
				}
			}
		}

		private void uGridDetalle_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ( uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "VALOR" )
			{
				Funcion.CamposNumericos(sender, e);
			}
			if (this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACION" ) Funcion.CamposNumericos(sender, e);
			
			if ((bool)this.uGridDetalle.ActiveRow.Cells["SF"].Value && (this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "RUC" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "RAZONSOCIAL" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "AUTORIZACION" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "SERIE" 
				|| this.uGridDetalle.ActiveCell.Column.Key.ToString().ToUpper() == "NUMERO")) e.Handled = true;

		}

		private void cmbCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtAcreditado.Focus();
		}

	

		private void btnIngresoGastos_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void frmGastosViajes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e); 																																																										 
				if (e.KeyCode == Keys.F5)	this.Consulta((int)this.txtidGastosviajes.Value);
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);								
			}
		}

		private void btnDeposito_Click(object sender, System.EventArgs e)
		{		
			DateTime DFecha;
			int ICuenta = 0;
			string sNumero = "";
			decimal dValor = 0;
			this.txtTotal.Value = 0;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDetalle.Rows.All)
				this.txtTotal.Value = Convert.ToDecimal(dr.Cells["Valor"].Value) + Convert.ToDecimal(this.txtTotal.Value);
			
			decimal dDiferencia =  Convert.ToDecimal (this.txtAcreditado.Value) -  Convert.ToDecimal (this.txtTotal.Value);
			
			if (dDiferencia > 0) 
			{
				using (frmDepositoGastosDeViaje Depositos = new frmDepositoGastosDeViaje( Convert.ToDecimal(  dDiferencia), (DateTime) this.dtFechaRetorna.Value))
				{
					if (DialogResult.OK == Depositos.ShowDialog())
					{
						DFecha = (DateTime) Depositos.dtFecha.Value;
						ICuenta = (int) Depositos.cmbCuenta.Value;
						sNumero = (string) Depositos.txtnDeposito.Text;
						dValor = (decimal)Depositos.txtValor.Value;

						string sSQLAsiento = string.Format("Exec AsientoDepositosGastosDeViaje {0},'{1}', {2}, '{3}', {4}", 
							(int)this.txtidGastosviajes.Value,Convert.ToDateTime(DFecha).ToString("yyyyMMdd") ,ICuenta, sNumero, dValor);
						this.txtidDeposito.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
															
						this.btnDeposito.Enabled = false;

						this.btnGuardar_Click (sender, e);
					}
				}
			}		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}

		private void uGridDetalle_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			#region Elimina
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar Los Gastos?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					if ((int)e.Rows[i].Cells["idGastosViaje"].Value > 0)
					{
						string sSQL = string.Format("Delete From DetalleGastosViaje where idDetalleGViaje = {0} ", (int)e.Rows[i].Cells["idDetalleGViaje"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

						e.DisplayPromptMsg = false;
					}
					else e.DisplayPromptMsg = false;
				}					
			}
			else e.Cancel = true;
			#endregion Elimina
		}

		private void uGridDetalle_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			this.uGridDetalle.UpdateData();

			if (e.Cell.Column.ToString() == "SF")
			{
				if (this.uGridDetalle.ActiveRow.Cells["Ruc"].Value.ToString().Length > 0 || this.uGridDetalle.ActiveRow.Cells["RazonSocial"].Value.ToString().Length > 0 
					|| this.uGridDetalle.ActiveRow.Cells["Autorizacion"].Value.ToString().Length > 0 || this.uGridDetalle.ActiveRow.Cells["Serie"].Value.ToString().Length > 0 
					|| this.uGridDetalle.ActiveRow.Cells["Numero"].Value.ToString().Length > 0)
				{
					MessageBox.Show("No puede crear un gasto sin Factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					e.Cancel= true;
					return;
				}
			}
			
		}

		private void btnEmpresa_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio (this.txtidPersonal,"Ingrese la Cedula del Personal",1,10000)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtFechaSalida.Value), Convert.ToDateTime(this.dtFechaRetorna.Value), this.dtFechaSalida, "La fecha de salida no puede ser posterior a la fecha de retorna", false)) return;


			using (frmReembolsoViaticos Viaticos = new frmReembolsoViaticos(false, (int)this.txtidPersonal.Value, (DateTime)this.dtFechaSalida.Value, (DateTime)this.dtFechaRetorna.Value , (int) this.txtidGastosviajes.Value, (int) this.optTransporte.Value))
			{
				if (DialogResult.OK == Viaticos.ShowDialog( ))
				{
       
					if (Viaticos.UgridDetalle.Rows.Count > 0)
					{
						string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", 0, 0, 0);
						FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultraGrid1);

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Viaticos.UgridDetalle.Rows.All)
						{
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idIngresoReembolsoViaticos"].Value = (int) dr.Cells["idIngresoReembolsoViaticos"].Value;
							this.ultraGrid1.ActiveRow.Cells["idReembolsoViaticos"].Value = (int) dr.Cells["idReembolsoViaticos"].Value;
							this.ultraGrid1.ActiveRow.Cells["Concepto"].Value = Convert.ToString ( dr.Cells["Concepto"].Value);
							this.ultraGrid1.ActiveRow.Cells["AdministrativoSolicitado"].Value = Convert.ToDecimal( dr.Cells["AdministrativoSolicitado"].Value);
							this.ultraGrid1.ActiveRow.Cells["GerentesSolicitado"].Value = Convert.ToDecimal( dr.Cells["GerentesSolicitado"].Value);
							this.ultraGrid1.ActiveRow.Cells["Administrativo"].Value = Convert.ToDecimal( dr.Cells["Administrativo"].Value);
							this.ultraGrid1.ActiveRow.Cells["Gerentes"].Value = Convert.ToDecimal( dr.Cells["Gerentes"].Value);
							this.ultraGrid1.ActiveRow.Cells["SubTotal"].Value = Convert.ToDecimal( dr.Cells["SubTotal"].Value);
							this.ultraGrid1.ActiveRow.Cells["SubTotalAcreditado"].Value = Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value);
							this.ultraGrid1.ActiveRow.Cells["idGastosViajes"].Value = (int) dr.Cells["idGastosViajes"].Value;

						}
						this.txtValorS.Value  = 0.00m ;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Viaticos.UgridDetalle.Rows.All)
						{
							
							this.txtValorS.Value = Convert.ToDecimal( dr.Cells["SubTotal"].Value) + (decimal)this.txtValorS.Value ;
						}
					}
				}
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
		
		}

		private void btnRevisar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio (this.txtidPersonal,"Ingrese la Cedula del Personal",1,10000)) return;
			using (frmReembolsoViaticos Viaticos = new frmReembolsoViaticos( true , (int)this.txtidPersonal.Value, (DateTime)this.dtFechaSalida.Value, (DateTime)this.dtFechaRetorna.Value , (int) this.txtidGastosviajes.Value, (int) this.optTransporte.Value))
			{
				if (DialogResult.OK == Viaticos.ShowDialog( ))
				{
       
					if (Viaticos.UgridDetalle.Rows.Count > 0)
					{
						string sSQLPa = string.Format(" EXEC ConsultaDetalleIngresoReembolsoViaticos   {0}, {1}, {2} ", 0, 0, 0);
						FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultraGrid1);

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Viaticos.UgridDetalle.Rows.All)
						{
							this.ultraGrid1.Rows.Band.AddNew();
							this.ultraGrid1.ActiveRow.Cells["idIngresoReembolsoViaticos"].Value = (int) dr.Cells["idIngresoReembolsoViaticos"].Value;
							this.ultraGrid1.ActiveRow.Cells["idReembolsoViaticos"].Value = (int) dr.Cells["idReembolsoViaticos"].Value;
							this.ultraGrid1.ActiveRow.Cells["Concepto"].Value = Convert.ToString ( dr.Cells["Concepto"].Value);
							this.ultraGrid1.ActiveRow.Cells["AdministrativoSolicitado"].Value = Convert.ToDecimal( dr.Cells["AdministrativoSolicitado"].Value);
							this.ultraGrid1.ActiveRow.Cells["GerentesSolicitado"].Value = Convert.ToDecimal( dr.Cells["GerentesSolicitado"].Value);
							this.ultraGrid1.ActiveRow.Cells["Administrativo"].Value = Convert.ToDecimal( dr.Cells["Administrativo"].Value);
							this.ultraGrid1.ActiveRow.Cells["Gerentes"].Value = Convert.ToDecimal( dr.Cells["Gerentes"].Value);
							this.ultraGrid1.ActiveRow.Cells["SubTotal"].Value = Convert.ToDecimal( dr.Cells["SubTotal"].Value);
							this.ultraGrid1.ActiveRow.Cells["SubTotalAcreditado"].Value = Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value);
							this.ultraGrid1.ActiveRow.Cells["idGastosViajes"].Value = (int) dr.Cells["idGastosViajes"].Value;

						}
						this.txtAcreditado.Value  = 0.00m ;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Viaticos.UgridDetalle.Rows.All)
						{
							
							this.txtAcreditado.Value = Convert.ToDecimal( dr.Cells["SubTotalAcreditado"].Value) + (decimal)this.txtAcreditado.Value ;
						}
					}
				}
			}
	
		}

		private void p(object sender, System.EventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			//			if ((int) this.txtEstado.Value == 2)
			//			{
			//				ParameterFields paramFields = new ParameterFields();
			//
			//				ParameterField paramFieldCedula = new ParameterField ();
			//				ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
			//				paramFieldCedula.ParameterFieldName = "@idGastosViaje";
			//				ValCedula.Value = (int)this.txtidGastosviajes.Value;
			//				paramFieldCedula.CurrentValues.Add (ValCedula);
			//				paramFields.Add (paramFieldCedula);
			//
			//				Reporte miRep = new Reporte("ReporteGastosViaje.rpt", "");
			//				miRep.MdiParent = this.MdiParent;
			//				miRep.Titulo("Planificacion De Carpa");
			//				miRep.crVista.ParameterFieldInfo = paramFields;
			//				miRep.Show();
			//
			//			}
			if ((int) this.txtEstado.Value >= 6)
			{
				ParameterFields paramFields = new ParameterFields();

				ParameterField paramFieldCedula = new ParameterField ();
				ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
				paramFieldCedula.ParameterFieldName = "@idGastosViaje";
				ValCedula.Value = (int)this.txtidGastosviajes.Value;
				paramFieldCedula.CurrentValues.Add (ValCedula);
				paramFields.Add (paramFieldCedula);

				Reporte miRep = new Reporte("DetalleGastosDeViajes2.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Planificacion De Carpa");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();

			}
		}

		private void dtInicia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaRetorna.Focus();
		}

		private void dtRetorna_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtValorS.Focus();
		}

		private void optTransporte_ValueChanged(object sender, System.EventArgs e)
		{
			if(bNuevo)
			{
				this.txtValorS.Value = 0;
			}
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if (sUrlNube.Length > 10)
			{
				if (!miAcceso.BVistas) 
				{
					MessageBox.Show("No tiene acceso a visualizar el Informe.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);				
			}
		}

		private void txtCedula_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}