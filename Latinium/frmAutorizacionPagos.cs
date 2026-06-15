using System;
using System.Globalization;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmAutorizacionPagos.
	/// </summary>
	public class frmAutorizacionPagos : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcAutorizacion;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuentaCorr;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinGrid.UltraDropDown cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuentas;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAutorizado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFaltante;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid3;
		private System.Windows.Forms.Label label8;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbFPTemp;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Label lblCobro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPago;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnAutorizar;
		private System.Windows.Forms.Button btnSeleccionar;
		private System.Windows.Forms.Button btnQuitarSeleccion;
		private System.Windows.Forms.Button btnAsignarValor;
		private System.Windows.Forms.Button btnAsignarCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;

		private Acceso miAcceso;

		public frmAutorizacionPagos()
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

		bool bLoad = false;

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Beneficiario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoPago", 0);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Autorizado", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Autorizado", 11, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 10, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 12, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RUC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Beneficiario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizado");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoPago");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Autorizado", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Autorizado", 12, true);
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 13, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RUC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NuevoPago");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta", -1, "cmbCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Autorizado");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NuevoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSol");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UsuarioAut");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaAut");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 11, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 12, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizacionPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RUC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Autorizado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NuevoPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UsuarioAut");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaAut");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmAutorizacionPagos));
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuentaCorr = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFPTemp = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label8 = new System.Windows.Forms.Label();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid3 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtFaltante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAutorizado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.utcAutorizacion = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbCuentas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblCobro = new System.Windows.Forms.Label();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnAutorizar = new System.Windows.Forms.Button();
			this.btnSeleccionar = new System.Windows.Forms.Button();
			this.btnQuitarSeleccion = new System.Windows.Forms.Button();
			this.btnAsignarValor = new System.Windows.Forms.Button();
			this.btnAsignarCuenta = new System.Windows.Forms.Button();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFPTemp)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFaltante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcAutorizacion)).BeginInit();
			this.utcAutorizacion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.cmbFormaPago);
			this.ultraTabPageControl1.Controls.Add(this.cmbCuentaCorr);
			this.ultraTabPageControl1.Controls.Add(this.cmbCuenta);
			this.ultraTabPageControl1.Controls.Add(this.ultraGrid2);
			this.ultraTabPageControl1.Controls.Add(this.cmbFPTemp);
			this.ultraTabPageControl1.Controls.Add(this.label8);
			this.ultraTabPageControl1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1308, 307);
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Hand;
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(512, 160);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(176, 86);
			this.cmbFormaPago.TabIndex = 57;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.Visible = false;
			// 
			// cmbCuentaCorr
			// 
			this.cmbCuentaCorr.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentaCorr.DisplayMember = "NumCuenta";
			this.cmbCuentaCorr.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentaCorr.Location = new System.Drawing.Point(312, 159);
			this.cmbCuentaCorr.Name = "cmbCuentaCorr";
			this.cmbCuentaCorr.Size = new System.Drawing.Size(192, 88);
			this.cmbCuentaCorr.TabIndex = 56;
			this.cmbCuentaCorr.ValueMember = "idCuentaCorriente";
			this.cmbCuentaCorr.Visible = false;
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(128, 160);
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(176, 86);
			this.cmbCuenta.TabIndex = 55;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.Visible = false;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 88;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 90;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.Caption = "Tipo";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 140;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 95;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "R.U.C.";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 100;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 220;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Documento";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 120;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.Caption = "Fecha de Pago";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.Header.Caption = "Forma de Pago";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn9.Width = 100;
			ultraGridColumn10.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn10.Width = 180;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance2;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn11.Width = 75;
			appearance3.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance3.BackColor2 = System.Drawing.Color.LightSteelBlue;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance3;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.MaskInput = "";
			ultraGridColumn12.Width = 75;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.BackColor = System.Drawing.Color.LightBlue;
			appearance4.BackColor2 = System.Drawing.Color.LightBlue;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 75;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance5;
			ultraGridColumn14.Header.Caption = "Número";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 180;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "...";
			ultraGridColumn16.Header.Fixed = true;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn16.Width = 20;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Width = 220;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.Caption = "Usuario Solicita";
			ultraGridColumn18.Header.VisiblePosition = 19;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn19.Header.Caption = "Fecha Solicitud";
			ultraGridColumn19.Header.VisiblePosition = 20;
			ultraGridColumn19.Width = 90;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn20.Header.Caption = "Usuario Autoriza";
			ultraGridColumn20.Header.VisiblePosition = 21;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 111;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn21.Header.Caption = "Fecha Autorizado";
			ultraGridColumn21.Header.VisiblePosition = 22;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 93;
			ultraGridColumn22.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Width = 220;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.DataType = typeof(System.DateTime);
			ultraGridColumn23.DefaultCellValue = new System.DateTime(((long)(0)));
			ultraGridColumn23.Header.Caption = "Pago Saldo";
			ultraGridColumn23.Header.VisiblePosition = 15;
			ultraGridColumn23.Width = 80;
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
			ultraGridBand1.GroupHeaderLines = 4;
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.True;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			ultraGridBand1.Override.GroupByColumnHeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance8;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.Hidden = true;
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.LightBlue;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance9;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid2.DisplayLayout.GroupByBox.Hidden = true;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowGroupBy = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance13.FontData.BoldAsString = "True";
			appearance13.FontData.Name = "Arial";
			appearance13.FontData.SizeInPoints = 8F;
			appearance13.ForeColor = System.Drawing.Color.White;
			appearance13.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance13;
			this.ultraGrid2.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8F);
			this.ultraGrid2.Location = new System.Drawing.Point(6, 5);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1296, 296);
			this.ultraGrid2.TabIndex = 53;
			this.ultraGrid2.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_CellChange);
			this.ultraGrid2.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid2_BeforeCellUpdate);
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid2_KeyDown);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
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
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn21.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn22});
			// 
			// cmbFPTemp
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFPTemp.Appearance = appearance16;
			this.cmbFPTemp.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFPTemp.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFPTemp.DisplayMember = "FormaPago";
			this.cmbFPTemp.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFPTemp.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFPTemp.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFPTemp.Location = new System.Drawing.Point(104, 464);
			this.cmbFPTemp.Name = "cmbFPTemp";
			this.cmbFPTemp.Size = new System.Drawing.Size(256, 21);
			this.cmbFPTemp.TabIndex = 178;
			this.cmbFPTemp.ValueMember = "idFormaPago";
			this.cmbFPTemp.Visible = false;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(16, 464);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 17);
			this.label8.TabIndex = 177;
			this.label8.Text = "Forma de Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label8.Visible = false;
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl2.Enabled = false;
			this.ultraTabPageControl2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1308, 307);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance17.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance17;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn26.Header.Caption = "Tipo";
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Width = 140;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn27.Header.VisiblePosition = 3;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "R.U.C.";
			ultraGridColumn28.Header.VisiblePosition = 4;
			ultraGridColumn28.Width = 100;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 5;
			ultraGridColumn29.Width = 220;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.Caption = "Documento";
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Width = 120;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn31.Header.Caption = "Fecha De Pago";
			ultraGridColumn31.Header.VisiblePosition = 7;
			ultraGridColumn32.Header.Caption = "Forma De Pago";
			ultraGridColumn32.Header.VisiblePosition = 8;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn32.Width = 100;
			ultraGridColumn33.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn33.Header.VisiblePosition = 9;
			ultraGridColumn33.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn33.Width = 180;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance18;
			ultraGridColumn34.Header.Caption = "Número";
			ultraGridColumn34.Header.VisiblePosition = 10;
			ultraGridColumn34.Width = 50;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn35.Header.VisiblePosition = 11;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 75;
			appearance19.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance19.BackColor2 = System.Drawing.Color.LightSteelBlue;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance19;
			ultraGridColumn36.Format = "#,##0.00";
			ultraGridColumn36.Header.VisiblePosition = 12;
			ultraGridColumn36.Width = 75;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance20.BackColor = System.Drawing.Color.LightBlue;
			appearance20.BackColor2 = System.Drawing.Color.LightBlue;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance20;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 13;
			ultraGridColumn37.Width = 75;
			ultraGridColumn38.Header.Caption = "...";
			ultraGridColumn38.Header.VisiblePosition = 15;
			ultraGridColumn38.Width = 20;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn39.Header.VisiblePosition = 16;
			ultraGridColumn39.Width = 220;
			ultraGridColumn40.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn40.Header.VisiblePosition = 17;
			ultraGridColumn40.Width = 220;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn41.Header.Caption = "Usuario Solicita";
			ultraGridColumn41.Header.VisiblePosition = 18;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn42.Header.Caption = "Fecha Solicitud";
			ultraGridColumn42.Header.VisiblePosition = 19;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn43.Header.Caption = "Usuario Autoriza";
			ultraGridColumn43.Header.VisiblePosition = 20;
			ultraGridColumn43.Width = 100;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn44.Header.Caption = "Fecha Autorizado";
			ultraGridColumn44.Header.VisiblePosition = 21;
			ultraGridColumn44.Width = 100;
			ultraGridColumn45.Header.Caption = "Pago Saldo";
			ultraGridColumn45.Header.VisiblePosition = 14;
			ultraGridColumn45.Width = 80;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45});
			appearance21.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance21;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance22.BackColor = System.Drawing.Color.LightBlue;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance22;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings4,
																																															summarySettings5});
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance23.ForeColor = System.Drawing.Color.Black;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance23;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
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
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8F);
			this.ultraGrid1.Location = new System.Drawing.Point(6, 5);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1296, 296);
			this.ultraGrid1.TabIndex = 54;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(System.DateTime);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn35.DataType = typeof(System.Decimal);
			ultraDataColumn36.DataType = typeof(System.Decimal);
			ultraDataColumn37.DataType = typeof(bool);
			ultraDataColumn37.DefaultValue = false;
			ultraDataColumn41.DataType = typeof(System.DateTime);
			ultraDataColumn43.DataType = typeof(System.DateTime);
			ultraDataColumn44.DataType = typeof(System.DateTime);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn44});
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.ultraGrid3);
			this.ultraTabPageControl3.Enabled = false;
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(1308, 307);
			// 
			// ultraGrid3
			// 
			this.ultraGrid3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid3.DataSource = this.ultraDataSource1;
			appearance29.BackColor = System.Drawing.Color.White;
			this.ultraGrid3.DisplayLayout.Appearance = appearance29;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 73;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 65;
			ultraGridColumn48.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn48.Header.Caption = "Tipo";
			ultraGridColumn48.Header.VisiblePosition = 2;
			ultraGridColumn48.Width = 140;
			ultraGridColumn49.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn49.Header.VisiblePosition = 3;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 65;
			ultraGridColumn50.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn50.Header.Caption = "R.U.C.";
			ultraGridColumn50.Header.VisiblePosition = 4;
			ultraGridColumn50.Width = 100;
			ultraGridColumn51.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn51.Header.VisiblePosition = 5;
			ultraGridColumn51.Width = 250;
			ultraGridColumn52.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn52.Header.Caption = "Documento";
			ultraGridColumn52.Header.VisiblePosition = 6;
			ultraGridColumn52.Width = 120;
			ultraGridColumn53.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn53.Header.Caption = "Fecha De Pago";
			ultraGridColumn53.Header.VisiblePosition = 7;
			ultraGridColumn53.Width = 85;
			ultraGridColumn54.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn54.Header.Caption = "Forma De Pago";
			ultraGridColumn54.Header.VisiblePosition = 8;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn54.Width = 100;
			ultraGridColumn55.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn55.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn55.Header.VisiblePosition = 9;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn55.Width = 180;
			ultraGridColumn56.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn56.CellAppearance = appearance30;
			ultraGridColumn56.Format = "#,##0.00";
			ultraGridColumn56.Header.VisiblePosition = 10;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 80;
			ultraGridColumn57.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn57.CellAppearance = appearance31;
			ultraGridColumn57.Format = "#,##0.00";
			ultraGridColumn57.Header.VisiblePosition = 11;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 80;
			ultraGridColumn58.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance32.BackColor = System.Drawing.Color.LightBlue;
			appearance32.BackColor2 = System.Drawing.Color.LightBlue;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn58.CellAppearance = appearance32;
			ultraGridColumn58.Format = "#,##0.00";
			ultraGridColumn58.Header.VisiblePosition = 12;
			ultraGridColumn58.Width = 80;
			ultraGridColumn59.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn59.CellAppearance = appearance33;
			ultraGridColumn59.Header.Caption = "Numero";
			ultraGridColumn59.Header.VisiblePosition = 13;
			ultraGridColumn59.Width = 50;
			ultraGridColumn60.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn60.Header.VisiblePosition = 15;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 32;
			ultraGridColumn61.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn61.Header.VisiblePosition = 14;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 39;
			ultraGridColumn62.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn62.Header.VisiblePosition = 16;
			ultraGridColumn62.Width = 250;
			ultraGridColumn63.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn63.Header.VisiblePosition = 17;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 220;
			ultraGridColumn64.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn64.Header.Caption = "Usuario Solicita";
			ultraGridColumn64.Header.VisiblePosition = 18;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 65;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn65.Header.Caption = "Fecha Solicitud";
			ultraGridColumn65.Header.VisiblePosition = 19;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 60;
			ultraGridColumn66.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn66.Header.Caption = "Usuario Autoriza";
			ultraGridColumn66.Header.VisiblePosition = 20;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 90;
			ultraGridColumn67.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn67.Header.Caption = "Fecha Autorizado";
			ultraGridColumn67.Header.VisiblePosition = 21;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 70;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67});
			appearance34.BackColor = System.Drawing.Color.LightBlue;
			appearance34.BackColor2 = System.Drawing.Color.LightBlue;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance34;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.Hidden = true;
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.BackColor = System.Drawing.Color.LightBlue;
			appearance35.BackColor2 = System.Drawing.Color.LightBlue;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance35;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6,
																																															summarySettings7});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ultraGrid3.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance36.ForeColor = System.Drawing.Color.Black;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid3.DisplayLayout.Override.ActiveRowAppearance = appearance36;
			this.ultraGrid3.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid3.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid3.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid3.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid3.DisplayLayout.Override.CellAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 8F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid3.DisplayLayout.Override.HeaderAppearance = appearance39;
			this.ultraGrid3.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid3.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.ultraGrid3.Font = new System.Drawing.Font("Tahoma", 8F);
			this.ultraGrid3.Location = new System.Drawing.Point(6, 5);
			this.ultraGrid3.Name = "ultraGrid3";
			this.ultraGrid3.Size = new System.Drawing.Size(1296, 296);
			this.ultraGrid3.TabIndex = 55;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(System.DateTime);
			ultraDataColumn53.DataType = typeof(int);
			ultraDataColumn54.DataType = typeof(int);
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
			ultraDataColumn59.DataType = typeof(bool);
			ultraDataColumn60.DataType = typeof(System.DateTime);
			ultraDataColumn64.DataType = typeof(System.DateTime);
			ultraDataColumn66.DataType = typeof(System.DateTime);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn66});
			// 
			// txtFaltante
			// 
			this.txtFaltante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance42.BackColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance42.BackColorDisabled2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance42.ForeColor = System.Drawing.Color.Firebrick;
			appearance42.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtFaltante.Appearance = appearance42;
			this.txtFaltante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFaltante.Enabled = false;
			this.txtFaltante.FormatString = "#,##0.00";
			this.txtFaltante.Location = new System.Drawing.Point(1226, 72);
			this.txtFaltante.Name = "txtFaltante";
			this.txtFaltante.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFaltante.PromptChar = ' ';
			this.txtFaltante.Size = new System.Drawing.Size(86, 22);
			this.txtFaltante.TabIndex = 134;
			// 
			// txtAutorizado
			// 
			this.txtAutorizado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance43.BackColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance43.BackColorDisabled2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance43.ForeColor = System.Drawing.Color.Firebrick;
			appearance43.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtAutorizado.Appearance = appearance43;
			this.txtAutorizado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAutorizado.Enabled = false;
			this.txtAutorizado.FormatString = "#,##0.00";
			this.txtAutorizado.Location = new System.Drawing.Point(1226, 40);
			this.txtAutorizado.Name = "txtAutorizado";
			this.txtAutorizado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAutorizado.PromptChar = ' ';
			this.txtAutorizado.Size = new System.Drawing.Size(86, 22);
			this.txtAutorizado.TabIndex = 133;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance44.BackColorDisabled = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance44.BackColorDisabled2 = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance44.ForeColor = System.Drawing.Color.Firebrick;
			appearance44.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtTotal.Appearance = appearance44;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(1226, 7);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(86, 22);
			this.txtTotal.TabIndex = 132;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label7.ForeColor = System.Drawing.Color.Firebrick;
			this.label7.Location = new System.Drawing.Point(1146, 75);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(56, 17);
			this.label7.TabIndex = 131;
			this.label7.Text = "FALTANTE";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label6.ForeColor = System.Drawing.Color.Firebrick;
			this.label6.Location = new System.Drawing.Point(1146, 43);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(73, 17);
			this.label6.TabIndex = 130;
			this.label6.Text = "AUTORIZADO";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(1146, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(38, 17);
			this.label5.TabIndex = 129;
			this.label5.Text = "TOTAL";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 54;
			this.label1.Text = "Fecha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 53;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1308, 307);
			// 
			// utcAutorizacion
			// 
			this.utcAutorizacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcAutorizacion.Appearance = appearance45;
			this.utcAutorizacion.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.utcAutorizacion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcAutorizacion.Controls.Add(this.ultraTabPageControl1);
			this.utcAutorizacion.Controls.Add(this.ultraTabPageControl2);
			this.utcAutorizacion.Controls.Add(this.ultraTabPageControl3);
			this.utcAutorizacion.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.utcAutorizacion.Location = new System.Drawing.Point(5, 120);
			this.utcAutorizacion.Name = "utcAutorizacion";
			this.utcAutorizacion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcAutorizacion.Size = new System.Drawing.Size(1310, 328);
			this.utcAutorizacion.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcAutorizacion.TabIndex = 55;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Pendientes";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Autorizados";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "Pagados";
			this.utcAutorizacion.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															ultraTab1,
																																															ultraTab2,
																																															ultraTab3});
			this.utcAutorizacion.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.utcAutorizacion.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcAutorizacion_SelectedTabChanged);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(8, 73);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 126;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtNombre
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance46;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Location = new System.Drawing.Point(64, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(220, 22);
			this.txtNombre.TabIndex = 127;
			this.txtNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombre_KeyPress);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 128;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(690, 43);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 129;
			this.label4.Text = "Cuenta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCuentas
			// 
			this.cmbCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCuentas.Appearance = appearance47;
			this.cmbCuentas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuentas.DisplayMember = "Descripcion";
			this.cmbCuentas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuentas.Enabled = false;
			this.cmbCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuentas.Location = new System.Drawing.Point(746, 41);
			this.cmbCuentas.Name = "cmbCuentas";
			this.cmbCuentas.Size = new System.Drawing.Size(256, 21);
			this.cmbCuentas.TabIndex = 170;
			this.cmbCuentas.ValueMember = "idCuenta";
			// 
			// lblCobro
			// 
			this.lblCobro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblCobro.AutoSize = true;
			this.lblCobro.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCobro.Location = new System.Drawing.Point(690, 10);
			this.lblCobro.Name = "lblCobro";
			this.lblCobro.Size = new System.Drawing.Size(31, 16);
			this.lblCobro.TabIndex = 633;
			this.lblCobro.Text = "Pago";
			this.lblCobro.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPago
			// 
			this.txtPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance48.ForeColorDisabled = System.Drawing.Color.Red;
			this.txtPago.Appearance = appearance48;
			this.txtPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPago.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtPago.FormatString = "#,##0.00";
			this.txtPago.Location = new System.Drawing.Point(746, 7);
			this.txtPago.Name = "txtPago";
			this.txtPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPago.PromptChar = ' ';
			this.txtPago.Size = new System.Drawing.Size(112, 22);
			this.txtPago.TabIndex = 634;
			this.txtPago.Click += new System.EventHandler(this.txtPago_Click);
			this.txtPago.Validated += new System.EventHandler(this.txtPago_Validated);
			this.txtPago.Enter += new System.EventHandler(this.txtPago_Enter);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1320, 8);
			this.groupBox1.TabIndex = 637;
			this.groupBox1.TabStop = false;
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(304, 39);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 640;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(544, 72);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(112, 23);
			this.btnExportar.TabIndex = 641;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnAutorizar
			// 
			this.btnAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAutorizar.CausesValidation = false;
			this.btnAutorizar.Image = ((System.Drawing.Image)(resources.GetObject("btnAutorizar.Image")));
			this.btnAutorizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAutorizar.Location = new System.Drawing.Point(1018, 72);
			this.btnAutorizar.Name = "btnAutorizar";
			this.btnAutorizar.Size = new System.Drawing.Size(104, 23);
			this.btnAutorizar.TabIndex = 642;
			this.btnAutorizar.Text = "&Autorizar";
			this.btnAutorizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAutorizar.Click += new System.EventHandler(this.btnAutorizar_Click);
			// 
			// btnSeleccionar
			// 
			this.btnSeleccionar.CausesValidation = false;
			this.btnSeleccionar.Image = ((System.Drawing.Image)(resources.GetObject("btnSeleccionar.Image")));
			this.btnSeleccionar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSeleccionar.Location = new System.Drawing.Point(544, 7);
			this.btnSeleccionar.Name = "btnSeleccionar";
			this.btnSeleccionar.Size = new System.Drawing.Size(112, 23);
			this.btnSeleccionar.TabIndex = 643;
			this.btnSeleccionar.Text = "&Seleccionar Todo";
			this.btnSeleccionar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
			// 
			// btnQuitarSeleccion
			// 
			this.btnQuitarSeleccion.CausesValidation = false;
			this.btnQuitarSeleccion.Image = ((System.Drawing.Image)(resources.GetObject("btnQuitarSeleccion.Image")));
			this.btnQuitarSeleccion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnQuitarSeleccion.Location = new System.Drawing.Point(544, 40);
			this.btnQuitarSeleccion.Name = "btnQuitarSeleccion";
			this.btnQuitarSeleccion.Size = new System.Drawing.Size(112, 23);
			this.btnQuitarSeleccion.TabIndex = 644;
			this.btnQuitarSeleccion.Text = "&Quitar Seleccion";
			this.btnQuitarSeleccion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnQuitarSeleccion.Click += new System.EventHandler(this.btnQuitarSeleccion_Click);
			// 
			// btnAsignarValor
			// 
			this.btnAsignarValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsignarValor.CausesValidation = false;
			this.btnAsignarValor.Image = ((System.Drawing.Image)(resources.GetObject("btnAsignarValor.Image")));
			this.btnAsignarValor.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsignarValor.Location = new System.Drawing.Point(1018, 7);
			this.btnAsignarValor.Name = "btnAsignarValor";
			this.btnAsignarValor.Size = new System.Drawing.Size(104, 23);
			this.btnAsignarValor.TabIndex = 645;
			this.btnAsignarValor.Text = "&Asignar Valor";
			this.btnAsignarValor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsignarValor.Click += new System.EventHandler(this.btnAsignarValor_Click);
			// 
			// btnAsignarCuenta
			// 
			this.btnAsignarCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsignarCuenta.CausesValidation = false;
			this.btnAsignarCuenta.Image = ((System.Drawing.Image)(resources.GetObject("btnAsignarCuenta.Image")));
			this.btnAsignarCuenta.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsignarCuenta.Location = new System.Drawing.Point(1018, 40);
			this.btnAsignarCuenta.Name = "btnAsignarCuenta";
			this.btnAsignarCuenta.Size = new System.Drawing.Size(104, 23);
			this.btnAsignarCuenta.TabIndex = 646;
			this.btnAsignarCuenta.Text = "&Asignar Cuenta";
			this.btnAsignarCuenta.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsignarCuenta.Click += new System.EventHandler(this.btnAsignarCuenta_Click);
			// 
			// frmAutorizacionPagos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1320, 454);
			this.Controls.Add(this.btnAsignarCuenta);
			this.Controls.Add(this.btnAsignarValor);
			this.Controls.Add(this.btnQuitarSeleccion);
			this.Controls.Add(this.btnSeleccionar);
			this.Controls.Add(this.btnAutorizar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblCobro);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.cmbCuentas);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.utcAutorizacion);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtFaltante);
			this.Controls.Add(this.txtAutorizado);
			this.Controls.Add(this.txtTotal);
			this.KeyPreview = true;
			this.Name = "frmAutorizacionPagos";
			this.Text = "Autorización de Pagos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmAutorizacionPagos_KeyDown);
			this.Load += new System.EventHandler(this.frmAutorizacionPagos_Load);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentaCorr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFPTemp)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFaltante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAutorizado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcAutorizacion)).EndInit();
			this.utcAutorizacion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		bool bModificaEstado = false;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void utcAutorizacion_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
			if (!bLoad) return;

			this.txtPago.Enabled = false;			
			this.btnAsignarValor.Enabled = false;
			this.cmbCuentas.Enabled = false;
			this.btnAsignarCuenta.Enabled = false;			
			this.btnSeleccionar.Enabled = false;
			this.btnQuitarSeleccion.Enabled = false;
			this.btnAutorizar.Enabled = false;
			this.cmbCuentas.Value = 0;	

			if (this.utcAutorizacion.Tabs[0].Selected) 
			{
				if (miAcceso.BAPAsignaBancos)
				{
					this.cmbCuentas.Enabled = true;
					this.btnAsignarCuenta.Enabled = true;

					this.txtPago.Enabled = true;
					this.btnAsignarValor.Enabled = true;
				}
			
				if (miAcceso.BAPAutorizar)
				{
					this.btnSeleccionar.Enabled = true;
					this.btnQuitarSeleccion.Enabled = true;
					this.btnAutorizar.Enabled = true;
				}				
			}
			
			Consulta();
		}

		private void frmAutorizacionPagos_Load(object sender, System.EventArgs e)
		{				
			Funcion.Resolucion(this, this.Height, this.Width);

			#region Seguridad
			string stTipo = "0409";
			string stPermiso = "51";
							
			miAcceso = new Acceso(cdsSeteoF, stTipo);

			if (miAcceso.BAPAutorizar)
			{
				this.btnAutorizar.Enabled = true;
				this.btnSeleccionar.Enabled = true;
				this.btnQuitarSeleccion.Enabled = true;

				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Estado"].CellActivation = Activation.AllowEdit;
			}
			if (miAcceso.BAPAsignaBancos)
			{
				this.btnAsignarCuenta.Enabled = true;
				
				this.cmbCuentas.Enabled = true;
				
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["idFormaPago"].CellActivation = Activation.AllowEdit;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["idCuenta"].CellActivation = Activation.AllowEdit;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["Autorizado"].CellActivation = Activation.AllowEdit;
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["NuevoPago"].CellActivation = Activation.AllowEdit;
			}

			if (miAcceso.BAPActAutorizados)
			{								
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idFormaPago"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idCuenta"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Autorizado"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NuevoPago"].CellActivation = Activation.AllowEdit;
			}

			if (miAcceso.BAPVerAutorizados) this.utcAutorizacion.Tabs[1].Enabled = true;
			if (miAcceso.BAPVerPagados) this.utcAutorizacion.Tabs[2].Enabled = true;

   		if (miAcceso.CambiarNumeracion) this.ultraGrid2.DisplayLayout.Bands[0].Columns["Beneficiario"].Hidden = false;
			#endregion Seguridad

			bLoad = true;
						
			this.dtFecha.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();				

			#region Listas
			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 7) order by idFormaPago desc");
			this.cmbFormaPago.DisplayLayout.Bands[0].Columns["FormaPago"].Width = 200;
			this.cmbFormaPago.DisplayLayout.Bands[0].Columns["idFormaPago"].Hidden = true;

			this.cmbFPTemp.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 7) order by idFormaPago desc ");
			this.cmbFPTemp.DisplayLayout.Bands[0].Columns["FormaPago"].Width = 200;
			this.cmbFPTemp.DisplayLayout.Bands[0].Columns["idFormaPago"].Hidden = true;

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasAutorizacionPagos");
			this.cmbCuenta.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 200;
			this.cmbCuenta.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;

			this.cmbCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ListaCuentasAutorizacionPagos");
			this.cmbCuentas.DisplayLayout.Bands[0].Columns["Descripcion"].Width = 250;
			this.cmbCuentas.DisplayLayout.Bands[0].Columns["idCuenta"].Hidden = true;
			#endregion Listas
		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Autorizado", 2);
		}

		private void Sumatoria()
		{	
			this.txtTotal.Value = 0;
			this.txtAutorizado.Value = 0;
			this.txtFaltante.Value = 0;

			SqlDataReader drSuma = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec SumaAutorizacionesDePago '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")), true);
			drSuma.Read();
			if (drSuma.HasRows)
			{
				this.txtTotal.Value = Convert.ToDecimal(drSuma.GetValue(0));
				this.txtAutorizado.Value = Convert.ToDecimal(drSuma.GetValue(1));
				this.txtFaltante.Value = Convert.ToDecimal(drSuma.GetValue(2));
			}
			drSuma.Close();
		}
		
		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

			if (e.Cell.Column.ToString() == "idFormaPago")
			{				
				int idFormaPago = 0;

				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value) 
					idFormaPago = (int)e.Cell.Row.Cells["idFormaPago"].Value;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set idFormaPago = {0} Where idAutorizacionPago = {1}", 
					idFormaPago, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "idCuenta")
			{				
				int idCuenta = 0;

				if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value) 
					idCuenta = (int)e.Cell.Row.Cells["idCuenta"].Value;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set idCuenta = {0} Where idAutorizacionPago = {1}", idCuenta, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "Autorizado")
			{ 

				decimal Saldo = (decimal)e.Cell.Row.Cells["Saldo"].Value;
				decimal Autorizado = (decimal)e.Cell.Row.Cells["Autorizado"].Value;

				if (!miAcceso.CambiarPrecio)
				{
					if (Math.Round(Autorizado, 2) > 1000.00m) 
					{
						e.Cell.Row.Cells["Autorizado"].Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show(string.Format("El valor máximo a autorizar no puede ser mayor a {0}.", 1000), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}
				}

				if (Math.Round(Autorizado, 2) > Math.Round(Saldo, 2)) 
				{
					e.Cell.Row.Cells["Autorizado"].Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show(string.Format("El valor máximo a autorizar no puede ser mayor a {0}.", Saldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (Math.Round(Autorizado, 2) < Math.Round(Saldo, 2))
				{
					DateTime FPago = (DateTime)e.Cell.Row.Cells["FechaPago"].Value;
					e.Cell.Row.Cells["NuevoPago"].Value = FPago.AddDays(8);
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Autorizado = {0} Where idAutorizacionPago = {1}", 
					(decimal)e.Cell.Row.Cells["Autorizado"].Value, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);

				Sumatoria();
			}

			if (e.Cell.Column.ToString() == "NuevoPago")
			{ 				
				if (e.Cell.Row.Cells["NuevoPago"].Value != System.DBNull.Value)
				{
					DateTime FPagoAnt = (DateTime)e.Cell.Row.Cells["FechaPago"].Value;				 
					DateTime FPagoNue = (DateTime)e.Cell.Row.Cells["NuevoPago"].Value;

					if (FPagoAnt > FPagoNue)
					{
						e.Cell.Row.Cells["NuevoPago"].Value = (DateTime)e.Cell.OriginalValue;
						MessageBox.Show(string.Format("La fecha para el pago del saldo no puede ser antes de '{0}'.", FPagoAnt), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}

      if (e.Cell.Column.ToString() == "Notas")
			{									
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Notas = '{0}' Where idAutorizacionPago = {1}", 
					e.Cell.Row.Cells["Notas"].Value.ToString(), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}			
		}

		private void txtNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnVer_Click(sender, e);
		}

		private void ultraGrid2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid2);
		}

		private void frmAutorizacionPagos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) Close();			
		}

		private void ultraGrid2_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{			
			if (e.Cell.Column.ToString() == "Autorizado")
			{ 
				if (((int)e.Cell.Row.Cells["idTipo"].Value != 4) && !bModificaEstado)
				{
					MessageBox.Show("Para este tipo de pago no es permitido cambiar el valor a pagar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}
		}		

		private void Suma ()
		{
			this.txtAutorizado.Value = 0;
			this.txtFaltante.Value = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value) this.txtAutorizado.Value = (decimal)this.txtAutorizado.Value + (decimal)dr.Cells["Autorizado"].Value;
			}
			
			this.txtFaltante.Value = (decimal)this.txtTotal.Value - (decimal)this.txtAutorizado.Value;
		}

		private void ultraGrid2_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Estado")
			{				
				this.ultraGrid2.UpdateData();

				int iTemp = 0;

				if ((bool)e.Cell.Row.Cells["Estado"].Value) iTemp = 1;
				
				bModificaEstado = true;
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Sel = {0}, Notas = '{1}' Where idAutorizacionPago = {2}", 
					iTemp, e.Cell.Row.Cells["Notas"].Value.ToString(), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);

				if (iTemp == 1)
				{						
					if ((decimal)e.Cell.Row.Cells["Autorizado"].Value == 0.00m) e.Cell.Row.Cells["Autorizado"].Value = (decimal)e.Cell.Row.Cells["Saldo"].Value;						
					if (this.cmbCuentas.ActiveRow != null) e.Cell.Row.Cells["idCuenta"].Value = (int)this.cmbCuentas.Value;
				}
				else 
				{						
					e.Cell.Row.Cells["Autorizado"].Value = 0.00m;						
					e.Cell.Row.Cells["idCuenta"].Value = System.DBNull.Value;					
				}
				bModificaEstado = false;
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Autorizado", 2);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idFormaPago")
			{				
				int idFormaPago = 0;

				if (e.Cell.Row.Cells["idFormaPago"].Value != System.DBNull.Value) 
					idFormaPago = (int)e.Cell.Row.Cells["idFormaPago"].Value;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set idFormaPago = {0} Where idAutorizacionPago = {1}", 
					idFormaPago, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "idCuenta")
			{				
				int idCuenta = 0;

				if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value) 
					idCuenta = (int)e.Cell.Row.Cells["idCuenta"].Value;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set idCuenta = {0} Where idAutorizacionPago = {1}", 
					idCuenta, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}

			if (e.Cell.Column.ToString() == "Autorizado")
			{ 
				decimal Saldo = (decimal)e.Cell.Row.Cells["Saldo"].Value;
				decimal Autorizado = (decimal)e.Cell.Row.Cells["Autorizado"].Value;

				if (Math.Round(Autorizado, 2) > Math.Round(Saldo, 2)) 
				{
					e.Cell.Row.Cells["Autorizado"].Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show(string.Format("El valor máximo a autorizar no puede ser mayor a {0}.", Saldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (Math.Round(Autorizado, 2) < Math.Round(Saldo, 2))
				{
					DateTime FPago = (DateTime)e.Cell.Row.Cells["FechaPago"].Value;
					e.Cell.Row.Cells["NuevoPago"].Value = FPago.AddDays(8);
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Autorizado = {0} Where idAutorizacionPago = {1}", 
					(decimal)e.Cell.Row.Cells["Autorizado"].Value, (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);

				Sumatoria();
			}

			if (e.Cell.Column.ToString() == "NuevoPago")
			{ 
				DateTime FPagoAnt = (DateTime)e.Cell.Row.Cells["FechaPago"].Value;
				DateTime FPagoNue = (DateTime)e.Cell.Row.Cells["NuevoPago"].Value;
				if (FPagoAnt > FPagoNue)
				{
					e.Cell.Row.Cells["NuevoPago"].Value = (DateTime)e.Cell.OriginalValue;
					MessageBox.Show(string.Format("La fecha para el pago del saldo no puede ser antes de '{0}'.", FPagoAnt), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}

			if (e.Cell.Column.ToString() == "Notas")
			{									
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Notas = '{0}' Where idAutorizacionPago = {1}", 
					e.Cell.Row.Cells["Notas"].Value.ToString(), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);
			}
		}

		private void txtPago_Click(object sender, System.EventArgs e)
		{
			this.txtPago.SelectAll();
		}

		private void txtPago_Enter(object sender, System.EventArgs e)
		{
			this.txtPago.SelectAll();
		}

		private void txtPago_Validated(object sender, System.EventArgs e)
		{
			if (this.txtPago.Value == System.DBNull.Value) this.txtPago.Value = 0;
		}

		private void Consulta()
		{
			int iEstado = 0;

			if (this.utcAutorizacion.Tabs[0].Selected) iEstado = 0;
			if (this.utcAutorizacion.Tabs[1].Selected) iEstado = 1;
			if (this.utcAutorizacion.Tabs[2].Selected) iEstado = 2;
							
//			this.ultraGrid1.DataSource = this.ultraDataSource3;
//			this.ultraGrid2.DataSource = this.ultraDataSource2;
//			this.ultraGrid3.DataSource = this.ultraDataSource1;

			string sSQL = string.Format("Exec ListaDeautorizacionesDePago '{0}', '{1}', {2}", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtNombre.Text.ToString(), iEstado);
			
			if (this.utcAutorizacion.Tabs[0].Selected)
			{				
				FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid2);

				this.lblContador.Text = this.ultraGrid2.Rows.Count + " REGISTROS ENCONTRADOS";

				this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
				this.ultraGrid2.Rows.CollapseAll(true);				
			}
			if (this.utcAutorizacion.Tabs[1].Selected) 
			{				
				FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);
				
				this.lblContador.Text = this.ultraGrid1.Rows.Count + " REGISTROS ENCONTRADOS";
			}			
			if (this.utcAutorizacion.Tabs[2].Selected)
			{				
				FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid3);							

				this.lblContador.Text = this.ultraGrid3.Rows.Count + " REGISTROS ENCONTRADOS";
			}
			
			Sumatoria();
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			if (this.utcAutorizacion.Tabs[0].Selected) FuncionesProcedimientos.ExportaExcel(this.ultraGrid2);

			if (this.utcAutorizacion.Tabs[1].Selected) FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);

			if (this.utcAutorizacion.Tabs[2].Selected) FuncionesProcedimientos.ExportaExcel(this.ultraGrid3);			
		}

		private void btnAutorizar_Click(object sender, System.EventArgs e)
		{
			#region Valida si hay registros en el grid
			if (this.ultraGrid2.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos para autorizar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			#endregion Valida si hay registros en el grid

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, false);

			#region Valida si se ha seleccionado un pago
			int i = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				if ((bool)dr.Cells["Estado"].Value == true) i++;

			if (i == 0)
			{
				MessageBox.Show("Seleccione el o los pagos que va a autorizar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
				return;
			}
			#endregion Valida si se ha seleccionado un pago

			#region Valida Forma de pago - Cuenta - Valor - Fecha Nuevo Pago
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value)
				{
					#region Forma de pago
					if (dr.Cells["idFormaPago"].Value == DBNull.Value)
					{
						MessageBox.Show(string.Format("Seleccione la forma de pago para la autorización de '{0}', Número '{1}'.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
						return;
					}
					if ((int)dr.Cells["idFormaPago"].Value == 0)
					{
						MessageBox.Show(string.Format("Seleccione la forma de pago para la autorización de '{0}', Número '{1}'.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
						return;
					}
					if ((int)dr.Cells["idFormaPago"].Value != 2 && (int)dr.Cells["idFormaPago"].Value != 7 && (int)dr.Cells["idFormaPago"].Value != 16)
					{
						MessageBox.Show(string.Format("La forma de pago para la autorización de '{0}', Número '{1}', debe ser CHEQUE o TRANSFERENCIA.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
						return;
					}
					#endregion Forma de pago

					#region Cuenta Bancaria
					if (dr.Cells["idCuenta"].Value == DBNull.Value)
					{
						MessageBox.Show(string.Format("Seleccione la Cuenta para la autorización de '{0}', Número '{1}'.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
						return;
					}

					if ((int)dr.Cells["idCuenta"].Value == 0)
					{
						MessageBox.Show(string.Format("Seleccione la Cuenta para la autorización de '{0}', Número '{1}'.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
						return;
					}
					#endregion Cuenta Bancaria

					#region Valor autorizado mayor a Cero - Fecha de pago si no se paga el total del pago
					if (dr.Cells["Autorizado"].Value == DBNull.Value || (decimal)dr.Cells["Autorizado"].Value < 0.00m)
					{
						MessageBox.Show(string.Format("Ingrese el valor de la autorización de '{0}', Número '{1}'.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
						return;
					}
			
					if (dr.Cells["Autorizado"].Value == DBNull.Value || (decimal)dr.Cells["Autorizado"].Value == 0.00m)
					{
						MessageBox.Show("No Puede Autorizas Pagos con Valor 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}

					if (Math.Round((decimal)dr.Cells["Autorizado"].Value, 2) < Math.Round((decimal)dr.Cells["Saldo"].Value, 2))
					{
						if (dr.Cells["NuevoPago"].Value == System.DBNull.Value)
						{
							MessageBox.Show(string.Format("Ingrese la fecha para el pago del saldo de la autorización de '{0}', Número '{1}'.", dr.Cells["Nombre"].Value.ToString(), dr.Cells["Numero"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
							return;
						}
					}	
					#endregion Valor autorizado mayor a Cero - Fecha de pago si no se paga el total del pago
				}
			}
			#endregion Valida Forma de pago - Cuenta - Valor - Fecha Nuevo Pago

			#region Valida Saldo Anticipos Notas De Credito
			decimal dSaldoNC = 0.00m;
			decimal dSaldoA = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value)
				{					
					if ((int)dr.Cells["idTipo"].Value == 4)
					{
						dSaldoNC = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(SUM(ISNULL(SaldoAnticipo, 0)), 0) From Compra Where idTipoFactura = 6 And Borrar = 0 And idCliente In (Select idCliente From Cliente Where Ruc = '{0}' And Proveedor = 1)", dr.Cells["RUC"].Value.ToString()));
						dSaldoA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select ISNULL(ROUND(SUM(ISNULL(Saldo, 0)), 2), 0) From Anticipos Where idCliente In (Select idCliente From Cliente Where Ruc = '{0}' And Proveedor = 1) And Estado = 5 And ClienteProveedor = 1", dr.Cells["RUC"].Value.ToString()));

						if (dSaldoNC > 0.00m)
						{
							if (DialogResult.No == MessageBox.Show(string.Format("El Proveedor '{0}',\n\ntiene Notas de Credito con un Saldo de {1} por cruzar\n\n\n¿Desea Continuar con la Autorización de este Pago?", dr.Cells["Nombre"].Value.ToString(), dSaldoNC), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) 
							{							
								this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
								return;
							}
						}

						if (dSaldoA > 0.00m)
						{
							if (DialogResult.No == MessageBox.Show(string.Format("El Proveedor '{0}',\n\ntiene Anticipos con un Saldo de {1} por cruzar\n\n\n¿Desea Continuar con la Autorización de este Pago?", dr.Cells["Nombre"].Value.ToString(), dSaldoA), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2)) 
							{							
								this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
								return;
							}
						}
					}
				}
			}
			#endregion Valida Saldo Anticipos Notas De Credito

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if ((bool)dr.Cells["Estado"].Value)
				{					
					int IdCuenta = 0;
					if (dr.Cells["idCuenta"].Value != System.DBNull.Value) IdCuenta = (int)dr.Cells["idCuenta"].Value;

					#region Actualiza Autorizacion de pago
					string sSQL = string.Format("Exec PagosAutorizados {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '', '{8}'", 
						(int)dr.Cells["idAutorizacionPago"].Value, 
						(int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idTipo"].Value, 
						(decimal)dr.Cells["Autorizado"].Value, 
						Convert.ToDateTime(dr.Cells["FechaPago"].Value).ToString("yyyyMMdd"), (int)dr.Cells["idFormaPago"].Value, 
						IdCuenta, dr.Cells["Notas"].Value.ToString(), dr.Cells["Beneficiario"].Value.ToString());
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
					#endregion Actualiza Autorizacion de pago

					#region Envio Email
					if ((int) dr.Cells["idFormaPago"].Value == 7)
					{
						string sQLenvio = string.Format("Exec PagosAutorizadosEnvioCorreo {0}, '{1}','{2}','{3}','{4}' ",
							(int)dr.Cells["idAutorizacionPago"].Value,
								 dr.Cells["TipoPrestamo"].Value.ToString() ,
								 dr.Cells["RUC"].Value.ToString() ,
							 dr.Cells["Nombre"].Value.ToString() ,
						  dr.Cells["Numero"].Value.ToString() );
						Funcion.EjecutaSQL(cdsSeteoF, sQLenvio, true);
					}
					#endregion Envio Email

					#region Envio Email solidario
					if ( MenuLatinium.stIdDB == 1  )
					{
						if ((int) dr.Cells["idTipo"].Value == 8)
						{
							string sQLenvio = string.Format("Exec [PagosAutorizadosEnvioCorreoSolidario] {0}, '{1}','{2}','{3}','{4}', {5} ",
								(int)dr.Cells["idAutorizacionPago"].Value,
								dr.Cells["TipoPrestamo"].Value.ToString() ,
								dr.Cells["RUC"].Value.ToString() ,
								dr.Cells["Nombre"].Value.ToString() ,
								dr.Cells["Numero"].Value.ToString(),
								(int)dr.Cells["idOrigen"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sQLenvio, true);
						}

					}
					#endregion Envio Email solidario

					#region Actualiza Estado Tipo Prestamo - Anticipo
					string sSQLActorigen = string.Format("Exec ActualizaEstadoOrigenesAutPago {0}, {1}, 1, {2}", 
						(int)dr.Cells["idTipo"].Value, (int)dr.Cells["idOrigen"].Value, Convert.ToDecimal(dr.Cells["Autorizado"].Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLActorigen);
					#endregion Actualiza Estado Tipo Prestamo - Anticipo
				}
			}

			MessageBox.Show("Autorizó con éxito los pagos seleccionados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);

			this.btnVer_Click(sender, e);
		}

		private void btnSeleccionar_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid2.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos para autorizar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, false);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				dr.Cells["Estado"].Value = true;

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
		}

		private void btnQuitarSeleccion_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid2.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos seleccionados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, false);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				dr.Cells["Estado"].Value = false;

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
		}

		private void btnAsignarValor_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid2.Rows.Count == 0)
			{
				MessageBox.Show("No existen filas para asignar pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}	
			
			if (Convert.ToDecimal(this.txtPago.Value) == 0.00m)
			{
				MessageBox.Show("Ingrese un valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtPago.Focus();
				return;
			}			

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, false);

			if (this.ultraGrid2.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

				this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);

				return;
			}	
			
			string sRUC = this.ultraGrid2.ActiveRow.Cells["RUC"].Value.ToString();
			decimal dValor = Convert.ToDecimal(this.txtPago.Value);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				string RUC = dr.Cells["RUC"].Value.ToString();

				if (RUC == sRUC)
				{
					if (dValor > (decimal)dr.Cells["Valor"].Value)
					{
						dr.Cells["Autorizado"].Value = (decimal)dr.Cells["Valor"].Value;
						dr.Cells["Estado"].Value = true;
						dValor = dValor - (decimal)dr.Cells["Valor"].Value;
					}
					else if (dValor < (decimal)dr.Cells["Valor"].Value && dValor > 0)
					{
						dr.Cells["Autorizado"].Value = dValor;
						dr.Cells["Estado"].Value = true;
						dValor = 0;
					}
					else if (dValor == (decimal)dr.Cells["Valor"].Value)
					{
						dr.Cells["Autorizado"].Value = dValor;
						dr.Cells["Estado"].Value = true;
						dValor = 0;
					}				
				}				
			}

			this.txtPago.Value = 0;

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
			this.ultraGrid2.Rows.CollapseAll(false);
		}

		private void btnAsignarCuenta_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid2.Rows.Count == 0)
			{
				MessageBox.Show("No existen pagos para asignarles una cuenta.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbCuentas, "Seleccione una cuenta")) return;

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, false);

			int i = 0;		

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				if ((bool)dr.Cells["Estado"].Value == true) i++;

			if (i == 0)
			{
				MessageBox.Show(string.Format("Seleccione el o los pagos para asignarles la cuenta {0}.", this.cmbCuentas.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				if ((bool)dr.Cells["Estado"].Value)	dr.Cells["idCuenta"].Value = (int)this.cmbCuentas.Value;

			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Nombre", false, true);
		}

		bool bActualiza = false;

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "Estado")
			{				
				this.ultraGrid1.UpdateData();

				int iTemp = 0;				

//				if (  MenuLatinium.stIdDB == 1 )
//				{
//					if ((int) e.Cell.Row.Cells["idTipo"].Value == 8)
//					{
//						if (!(bool) e.Cell.Row.Cells["Estado"].Value )
//						{
//							bActualiza = true;											
//							e.Cell.Row.Cells["Estado"].Value = (bool)e.Cell.OriginalValue;	
//					
//							MessageBox.Show("REPOSICIÓN CARTERA SOLIDARIO no se puede reversar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					
//							bActualiza = false;
//							return;
//						}
//					}
//				}

				if ((bool)e.Cell.Row.Cells["Estado"].Value) iTemp = 1;
				
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update AutorizacionDePagos Set Sel = {0}, Estado = {0}, Notas = '{1}' Where idAutorizacionPago = {2}", 
					iTemp, e.Cell.Row.Cells["Notas"].Value.ToString(), (int)e.Cell.Row.Cells["idAutorizacionPago"].Value), true);

				string sSQLActorigen = string.Format("Exec ActualizaEstadoOrigenesAutPago {0}, {1}, 0, 0", (int)e.Cell.Row.Cells["idTipo"].Value, (int)e.Cell.Row.Cells["idOrigen"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLActorigen);			
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNombre.Focus();
		}
	}
}

