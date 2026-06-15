using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data.SqlClient;
using System.Globalization;
using C1.Data;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmImportacion.
	/// </summary>
	public class frmImportacion : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasViaje;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasDeProduccion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesAduanizacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtLlegadaPuerto;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEmbarque;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFin;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtIngresoBodega;
		private System.Windows.Forms.Button btProgramacion;
		private System.Windows.Forms.Button btFactura;
		private System.Windows.Forms.Button btImportacionTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Button btAsiento;
		private System.Windows.Forms.Button btResumenImp;
		private System.Windows.Forms.Button btSalir;
		private System.Windows.Forms.Button btEditar;
		private System.Windows.Forms.Button btGrabar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSalidaCapitales;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIce;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFodInfa;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAdValorem;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoCalculo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Button btCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCIF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSeguro;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFleteExterior;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFOBEXW;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFOB;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalImportacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdFlete;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblSaldo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIni;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.ContextMenu mnuAsientos;
		private System.Windows.Forms.MenuItem mnuAsientosFactura;
		private System.Windows.Forms.MenuItem mnuAsientosLiquidacion;
		private System.Windows.Forms.MenuItem mnuAsientosIva;
		private System.Windows.Forms.Button btnActualizar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdEgreso;
		private System.Windows.Forms.Label lblPagoIVA;
		private System.Windows.Forms.MenuItem mnuAsientosEgreso;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPallets;
		private System.Windows.Forms.MenuItem mnuAsientosGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidAsiento2;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiferenciaISC;
		private System.ComponentModel.IContainer components;

		public frmImportacion()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmImportacion));
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Flete");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Flete");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Calculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Calculo");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacionDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoAnticipo");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Importacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo", -1, "cmbTipoCalculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Adicional");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalS");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoObligatorio");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionDet");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Iva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoImportacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Importacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Adicional");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TotalS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NoObligatorio");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton7 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			this.label19 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDiasViaje = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiasDeProduccion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.dtDesAduanizacion = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtLlegadaPuerto = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaEmbarque = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtIngresoBodega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btProgramacion = new System.Windows.Forms.Button();
			this.btFactura = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btImportacionTipo = new System.Windows.Forms.Button();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btAsiento = new System.Windows.Forms.Button();
			this.btResumenImp = new System.Windows.Forms.Button();
			this.btSalir = new System.Windows.Forms.Button();
			this.btCancelar = new System.Windows.Forms.Button();
			this.btEditar = new System.Windows.Forms.Button();
			this.btGrabar = new System.Windows.Forms.Button();
			this.txtSalidaCapitales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblSaldo = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtTotalImportacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIce = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFodInfa = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtAdValorem = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCIF = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSeguro = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFleteExterior = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFOBEXW = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFOB = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbTipoCalculo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblNumero = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdFlete = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.mnuAsientos = new System.Windows.Forms.ContextMenu();
			this.mnuAsientosFactura = new System.Windows.Forms.MenuItem();
			this.mnuAsientosLiquidacion = new System.Windows.Forms.MenuItem();
			this.mnuAsientosIva = new System.Windows.Forms.MenuItem();
			this.mnuAsientosEgreso = new System.Windows.Forms.MenuItem();
			this.mnuAsientosGastos = new System.Windows.Forms.MenuItem();
			this.btnActualizar = new System.Windows.Forms.Button();
			this.txtIdEgreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPagoIVA = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtPallets = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtidAsiento2 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.txtDiferenciaISC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasViaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeProduccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesAduanizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLlegadaPuerto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmbarque)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngresoBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalidaCapitales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalImportacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIce)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFodInfa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdValorem)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSeguro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFleteExterior)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFOBEXW)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFOB)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCalculo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFlete)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdEgreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPallets)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferenciaISC)).BeginInit();
			this.SuspendLayout();
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(410, 67);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(79, 16);
			this.label19.TabIndex = 270;
			this.label19.Text = "DIAS DE VIAJE";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(410, 43);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(74, 16);
			this.label17.TabIndex = 269;
			this.label17.Text = "PRODUCCIÓN";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(118, 91);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(105, 16);
			this.label11.TabIndex = 268;
			this.label11.Text = "DESADUANIZACIÓN";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(714, 67);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(106, 16);
			this.label16.TabIndex = 267;
			this.label16.Text = "LLEGADA A PUERTO";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(118, 67);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(62, 16);
			this.label9.TabIndex = 266;
			this.label9.Text = "EMBARQUE";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(714, 43);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(22, 16);
			this.label10.TabIndex = 265;
			this.label10.Text = "FIN";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(410, 91);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(115, 16);
			this.label8.TabIndex = 264;
			this.label8.Text = "INGRESO DE BODEGA";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(118, 43);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(39, 16);
			this.label20.TabIndex = 263;
			this.label20.Text = "INICIO";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.lblEstado.Location = new System.Drawing.Point(872, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 262;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(78, 16);
			this.label1.TabIndex = 261;
			this.label1.Text = "IMPORTACIÓN";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiasViaje
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasViaje.Appearance = appearance1;
			this.txtDiasViaje.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtDiasViaje.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasViaje.Enabled = false;
			this.txtDiasViaje.FormatString = "";
			this.txtDiasViaje.Location = new System.Drawing.Point(552, 64);
			this.txtDiasViaje.Name = "txtDiasViaje";
			this.txtDiasViaje.PromptChar = ' ';
			this.txtDiasViaje.Size = new System.Drawing.Size(80, 20);
			this.txtDiasViaje.TabIndex = 253;
			this.txtDiasViaje.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasViaje_KeyDown);
			this.txtDiasViaje.Click += new System.EventHandler(this.txtDiasViaje_Click);
			this.txtDiasViaje.Validating += new System.ComponentModel.CancelEventHandler(this.txtDiasViaje_Validating);
			this.txtDiasViaje.Validated += new System.EventHandler(this.txtDiasViaje_Validated);
			this.txtDiasViaje.Enter += new System.EventHandler(this.txtDiasViaje_Enter);
			// 
			// txtDiasDeProduccion
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasDeProduccion.Appearance = appearance2;
			this.txtDiasDeProduccion.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtDiasDeProduccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasDeProduccion.Enabled = false;
			this.txtDiasDeProduccion.FormatString = "";
			this.txtDiasDeProduccion.Location = new System.Drawing.Point(552, 40);
			this.txtDiasDeProduccion.Name = "txtDiasDeProduccion";
			this.txtDiasDeProduccion.PromptChar = ' ';
			this.txtDiasDeProduccion.Size = new System.Drawing.Size(80, 20);
			this.txtDiasDeProduccion.TabIndex = 250;
			this.txtDiasDeProduccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiasDeProduccion_KeyDown);
			this.txtDiasDeProduccion.Click += new System.EventHandler(this.txtDiasDeProduccion_Click);
			this.txtDiasDeProduccion.Validated += new System.EventHandler(this.txtDiasDeProduccion_Validated);
			this.txtDiasDeProduccion.Enter += new System.EventHandler(this.txtDiasDeProduccion_Enter);
			// 
			// dtDesAduanizacion
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesAduanizacion.Appearance = appearance3;
			this.dtDesAduanizacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesAduanizacion.DateButtons.Add(dateButton1);
			this.dtDesAduanizacion.Enabled = false;
			this.dtDesAduanizacion.Format = "dd/MM/yyyy";
			this.dtDesAduanizacion.Location = new System.Drawing.Point(246, 88);
			this.dtDesAduanizacion.Name = "dtDesAduanizacion";
			this.dtDesAduanizacion.NonAutoSizeHeight = 23;
			this.dtDesAduanizacion.Size = new System.Drawing.Size(112, 21);
			this.dtDesAduanizacion.SpinButtonsVisible = true;
			this.dtDesAduanizacion.TabIndex = 255;
			this.dtDesAduanizacion.Value = ((object)(resources.GetObject("dtDesAduanizacion.Value")));
			this.dtDesAduanizacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesAduanizacion_KeyDown);
			this.dtDesAduanizacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtDesAduanizacion_KeyPress);
			this.dtDesAduanizacion.Validating += new System.ComponentModel.CancelEventHandler(this.dtDesAduanizacion_Validating);
			this.dtDesAduanizacion.Validated += new System.EventHandler(this.dtDesAduanizacion_Validated);
			// 
			// dtLlegadaPuerto
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtLlegadaPuerto.Appearance = appearance4;
			this.dtLlegadaPuerto.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtLlegadaPuerto.DateButtons.Add(dateButton2);
			this.dtLlegadaPuerto.Enabled = false;
			this.dtLlegadaPuerto.Format = "dd/MM/yyyy";
			this.dtLlegadaPuerto.Location = new System.Drawing.Point(842, 64);
			this.dtLlegadaPuerto.Name = "dtLlegadaPuerto";
			this.dtLlegadaPuerto.NonAutoSizeHeight = 23;
			this.dtLlegadaPuerto.Size = new System.Drawing.Size(112, 21);
			this.dtLlegadaPuerto.SpinButtonsVisible = true;
			this.dtLlegadaPuerto.TabIndex = 254;
			this.dtLlegadaPuerto.Value = ((object)(resources.GetObject("dtLlegadaPuerto.Value")));
			// 
			// dtFechaEmbarque
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEmbarque.Appearance = appearance5;
			this.dtFechaEmbarque.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaEmbarque.DateButtons.Add(dateButton3);
			this.dtFechaEmbarque.Enabled = false;
			this.dtFechaEmbarque.Format = "dd/MM/yyyy";
			this.dtFechaEmbarque.Location = new System.Drawing.Point(246, 64);
			this.dtFechaEmbarque.Name = "dtFechaEmbarque";
			this.dtFechaEmbarque.NonAutoSizeHeight = 23;
			this.dtFechaEmbarque.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEmbarque.SpinButtonsVisible = true;
			this.dtFechaEmbarque.TabIndex = 252;
			this.dtFechaEmbarque.Value = ((object)(resources.GetObject("dtFechaEmbarque.Value")));
			this.dtFechaEmbarque.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaEmbarque_KeyDown);
			this.dtFechaEmbarque.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaEmbarque_KeyPress);
			this.dtFechaEmbarque.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaEmbarque_Validating);
			this.dtFechaEmbarque.Validated += new System.EventHandler(this.dtFechaEmbarque_Validated);
			// 
			// dtFechaFin
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaFin.Appearance = appearance6;
			this.dtFechaFin.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtFechaFin.DateButtons.Add(dateButton4);
			this.dtFechaFin.Enabled = false;
			this.dtFechaFin.Format = "dd/MM/yyyy";
			this.dtFechaFin.Location = new System.Drawing.Point(842, 40);
			this.dtFechaFin.Name = "dtFechaFin";
			this.dtFechaFin.NonAutoSizeHeight = 23;
			this.dtFechaFin.Size = new System.Drawing.Size(112, 21);
			this.dtFechaFin.SpinButtonsVisible = true;
			this.dtFechaFin.TabIndex = 251;
			this.dtFechaFin.Value = ((object)(resources.GetObject("dtFechaFin.Value")));
			// 
			// dtIngresoBodega
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtIngresoBodega.Appearance = appearance7;
			this.dtIngresoBodega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtIngresoBodega.DateButtons.Add(dateButton5);
			this.dtIngresoBodega.Enabled = false;
			this.dtIngresoBodega.Format = "dd/MM/yyyy";
			this.dtIngresoBodega.Location = new System.Drawing.Point(552, 88);
			this.dtIngresoBodega.Name = "dtIngresoBodega";
			this.dtIngresoBodega.NonAutoSizeHeight = 23;
			this.dtIngresoBodega.Size = new System.Drawing.Size(112, 21);
			this.dtIngresoBodega.SpinButtonsVisible = true;
			this.dtIngresoBodega.TabIndex = 256;
			this.dtIngresoBodega.Value = ((object)(resources.GetObject("dtIngresoBodega.Value")));
			this.dtIngresoBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtIngresoBodega_KeyDown);
			this.dtIngresoBodega.Validating += new System.ComponentModel.CancelEventHandler(this.dtIngresoBodega_Validating);
			// 
			// dtFechaIni
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIni.Appearance = appearance8;
			this.dtFechaIni.AutoSelect = false;
			this.dtFechaIni.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtFechaIni.DateButtons.Add(dateButton6);
			this.dtFechaIni.Enabled = false;
			this.dtFechaIni.Format = "dd/MM/yyyy";
			this.dtFechaIni.Location = new System.Drawing.Point(246, 40);
			this.dtFechaIni.Name = "dtFechaIni";
			this.dtFechaIni.NonAutoSizeHeight = 23;
			this.dtFechaIni.Size = new System.Drawing.Size(112, 21);
			this.dtFechaIni.SpinButtonsVisible = true;
			this.dtFechaIni.TabIndex = 249;
			this.dtFechaIni.Value = ((object)(resources.GetObject("dtFechaIni.Value")));
			this.dtFechaIni.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaIni_KeyDown);
			this.dtFechaIni.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaIni_KeyPress);
			this.dtFechaIni.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaIni_Validating);
			this.dtFechaIni.Validated += new System.EventHandler(this.dtFechaIni_Validated);
			this.dtFechaIni.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaIni_BeforeDropDown);
			// 
			// btProgramacion
			// 
			this.btProgramacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btProgramacion.Image = ((System.Drawing.Image)(resources.GetObject("btProgramacion.Image")));
			this.btProgramacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btProgramacion.Location = new System.Drawing.Point(493, 624);
			this.btProgramacion.Name = "btProgramacion";
			this.btProgramacion.Size = new System.Drawing.Size(91, 24);
			this.btProgramacion.TabIndex = 257;
			this.btProgramacion.Text = "Programación";
			this.btProgramacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btProgramacion.Click += new System.EventHandler(this.btProgramacion_Click);
			// 
			// btFactura
			// 
			this.btFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btFactura.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btFactura.ImageIndex = 18;
			this.btFactura.ImageList = this.imageList1;
			this.btFactura.Location = new System.Drawing.Point(723, 624);
			this.btFactura.Name = "btFactura";
			this.btFactura.Size = new System.Drawing.Size(72, 24);
			this.btFactura.TabIndex = 259;
			this.btFactura.Text = "Factura";
			this.btFactura.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btFactura.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btImportacionTipo
			// 
			this.btImportacionTipo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImportacionTipo.Image = ((System.Drawing.Image)(resources.GetObject("btImportacionTipo.Image")));
			this.btImportacionTipo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btImportacionTipo.Location = new System.Drawing.Point(592, 624);
			this.btImportacionTipo.Name = "btImportacionTipo";
			this.btImportacionTipo.Size = new System.Drawing.Size(123, 24);
			this.btImportacionTipo.TabIndex = 258;
			this.btImportacionTipo.Text = "Tipo de Importación";
			this.btImportacionTipo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btImportacionTipo.Click += new System.EventHandler(this.btImportacionTipo_Click);
			// 
			// cmbProyecto
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbProyecto.Appearance = appearance9;
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.ultraDataSource3;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 300;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 150;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(104, 8);
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(320, 21);
			this.cmbProyecto.TabIndex = 248;
			this.cmbProyecto.ValueMember = "idImportacionNombre";
			this.cmbProyecto.ValueChanged += new System.EventHandler(this.cmbProyecto_ValueChanged);
			this.cmbProyecto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProyecto_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btAsiento.ImageIndex = 36;
			this.btAsiento.ImageList = this.imageList1;
			this.btAsiento.Location = new System.Drawing.Point(803, 624);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(72, 24);
			this.btAsiento.TabIndex = 260;
			this.btAsiento.Text = "Asientos";
			this.btAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btResumenImp
			// 
			this.btResumenImp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btResumenImp.Image = ((System.Drawing.Image)(resources.GetObject("btResumenImp.Image")));
			this.btResumenImp.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btResumenImp.Location = new System.Drawing.Point(413, 624);
			this.btResumenImp.Name = "btResumenImp";
			this.btResumenImp.Size = new System.Drawing.Size(72, 24);
			this.btResumenImp.TabIndex = 300;
			this.btResumenImp.Text = "&Resumén";
			this.btResumenImp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btResumenImp.Click += new System.EventHandler(this.btResumenImp_Click);
			// 
			// btSalir
			// 
			this.btSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSalir.CausesValidation = false;
			this.btSalir.Image = ((System.Drawing.Image)(resources.GetObject("btSalir.Image")));
			this.btSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btSalir.Location = new System.Drawing.Point(966, 624);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(72, 24);
			this.btSalir.TabIndex = 279;
			this.btSalir.Text = "&Salir";
			this.btSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.CausesValidation = false;
			this.btCancelar.Enabled = false;
			this.btCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btCancelar.Image")));
			this.btCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btCancelar.Location = new System.Drawing.Point(333, 624);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(72, 24);
			this.btCancelar.TabIndex = 278;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btEditar
			// 
			this.btEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btEditar.Image = ((System.Drawing.Image)(resources.GetObject("btEditar.Image")));
			this.btEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btEditar.Location = new System.Drawing.Point(173, 624);
			this.btEditar.Name = "btEditar";
			this.btEditar.Size = new System.Drawing.Size(72, 24);
			this.btEditar.TabIndex = 271;
			this.btEditar.Text = "&Editar";
			this.btEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btEditar.Click += new System.EventHandler(this.btEditar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Enabled = false;
			this.btGrabar.Image = ((System.Drawing.Image)(resources.GetObject("btGrabar.Image")));
			this.btGrabar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btGrabar.Location = new System.Drawing.Point(253, 624);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(72, 24);
			this.btGrabar.TabIndex = 277;
			this.btGrabar.Text = "&Grabar";
			this.btGrabar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// txtSalidaCapitales
			// 
			this.txtSalidaCapitales.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSalidaCapitales.Appearance = appearance10;
			this.txtSalidaCapitales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSalidaCapitales.Enabled = false;
			this.txtSalidaCapitales.FormatString = "#,##0.00";
			this.txtSalidaCapitales.Location = new System.Drawing.Point(813, 496);
			this.txtSalidaCapitales.Name = "txtSalidaCapitales";
			this.txtSalidaCapitales.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSalidaCapitales.PromptChar = ' ';
			this.txtSalidaCapitales.Size = new System.Drawing.Size(96, 22);
			this.txtSalidaCapitales.TabIndex = 276;
			// 
			// label22
			// 
			this.label22.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(669, 499);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(122, 16);
			this.label22.TabIndex = 299;
			this.label22.Text = "5% SALIDA CAPITALES";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label21
			// 
			this.label21.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(8, 504);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(116, 16);
			this.label21.TabIndex = 297;
			this.label21.Text = "TOTAL IMPORTACIÓN";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label21.Visible = false;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(437, 523);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(22, 16);
			this.label5.TabIndex = 296;
			this.label5.Text = "ICE";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSaldo
			// 
			this.lblSaldo.AutoSize = true;
			this.lblSaldo.BackColor = System.Drawing.Color.Transparent;
			this.lblSaldo.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSaldo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.lblSaldo.Location = new System.Drawing.Point(714, 88);
			this.lblSaldo.Name = "lblSaldo";
			this.lblSaldo.Size = new System.Drawing.Size(0, 23);
			this.lblSaldo.TabIndex = 295;
			this.lblSaldo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(163, 547);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(47, 16);
			this.label15.TabIndex = 288;
			this.label15.Text = "SEGURO";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(163, 523);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(114, 16);
			this.label14.TabIndex = 287;
			this.label14.Text = "FLETE DEL EXTERIOR";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(163, 499);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(51, 16);
			this.label13.TabIndex = 286;
			this.label13.Text = "FOB EXW";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(163, 475);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(25, 16);
			this.label12.TabIndex = 285;
			this.label12.Text = "FOB";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label7.Location = new System.Drawing.Point(669, 523);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(91, 17);
			this.label7.TabIndex = 284;
			this.label7.Text = "TOTAL GASTOS";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(437, 499);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(50, 16);
			this.label4.TabIndex = 282;
			this.label4.Text = "FODINFA";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(437, 475);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(68, 16);
			this.label3.TabIndex = 281;
			this.label3.Text = "ADVALOREM";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label2.Location = new System.Drawing.Point(163, 571);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(24, 17);
			this.label2.TabIndex = 280;
			this.label2.Text = "CIF";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(437, 547);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(22, 16);
			this.label6.TabIndex = 283;
			this.label6.Text = "IVA";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalImportacion
			// 
			this.txtTotalImportacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalImportacion.Appearance = appearance11;
			this.txtTotalImportacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalImportacion.Enabled = false;
			this.txtTotalImportacion.FormatString = "#,##0.00";
			this.txtTotalImportacion.Location = new System.Drawing.Point(8, 520);
			this.txtTotalImportacion.Name = "txtTotalImportacion";
			this.txtTotalImportacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalImportacion.PromptChar = ' ';
			this.txtTotalImportacion.Size = new System.Drawing.Size(96, 22);
			this.txtTotalImportacion.TabIndex = 298;
			this.txtTotalImportacion.Visible = false;
			// 
			// txtIce
			// 
			this.txtIce.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIce.Appearance = appearance12;
			this.txtIce.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIce.Enabled = false;
			this.txtIce.FormatString = "#,##0.00";
			this.txtIce.Location = new System.Drawing.Point(525, 520);
			this.txtIce.Name = "txtIce";
			this.txtIce.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIce.PromptChar = ' ';
			this.txtIce.Size = new System.Drawing.Size(96, 22);
			this.txtIce.TabIndex = 274;
			this.txtIce.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIce_KeyDown);
			this.txtIce.Click += new System.EventHandler(this.txtIce_Click);
			this.txtIce.Validated += new System.EventHandler(this.txtIce_Validated);
			this.txtIce.Enter += new System.EventHandler(this.txtIce_Enter);
			// 
			// txtTotalGastos
			// 
			this.txtTotalGastos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGastos.Appearance = appearance13;
			this.txtTotalGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGastos.Enabled = false;
			this.txtTotalGastos.FormatString = "#,##0.00";
			this.txtTotalGastos.Location = new System.Drawing.Point(813, 520);
			this.txtTotalGastos.Name = "txtTotalGastos";
			this.txtTotalGastos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGastos.PromptChar = ' ';
			this.txtTotalGastos.Size = new System.Drawing.Size(96, 22);
			this.txtTotalGastos.TabIndex = 294;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance14;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(525, 544);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(96, 22);
			this.txtIva.TabIndex = 275;
			this.txtIva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIva_KeyDown);
			this.txtIva.Click += new System.EventHandler(this.txtIva_Click);
			this.txtIva.Validated += new System.EventHandler(this.txtIva_Validated);
			this.txtIva.ValueChanged += new System.EventHandler(this.txtIva_ValueChanged);
			this.txtIva.Enter += new System.EventHandler(this.txtIva_Enter);
			// 
			// txtFodInfa
			// 
			this.txtFodInfa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFodInfa.Appearance = appearance15;
			this.txtFodInfa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFodInfa.Enabled = false;
			this.txtFodInfa.FormatString = "#,##0.00";
			this.txtFodInfa.Location = new System.Drawing.Point(525, 496);
			this.txtFodInfa.Name = "txtFodInfa";
			this.txtFodInfa.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFodInfa.PromptChar = ' ';
			this.txtFodInfa.Size = new System.Drawing.Size(96, 22);
			this.txtFodInfa.TabIndex = 273;
			this.txtFodInfa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtFodInfa_KeyDown);
			this.txtFodInfa.Click += new System.EventHandler(this.txtFodInfa_Click);
			this.txtFodInfa.Validated += new System.EventHandler(this.txtFodInfa_Validated);
			this.txtFodInfa.Enter += new System.EventHandler(this.txtFodInfa_Enter);
			// 
			// txtAdValorem
			// 
			this.txtAdValorem.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAdValorem.Appearance = appearance16;
			this.txtAdValorem.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAdValorem.Enabled = false;
			this.txtAdValorem.FormatString = "#,##0.00";
			this.txtAdValorem.Location = new System.Drawing.Point(525, 472);
			this.txtAdValorem.Name = "txtAdValorem";
			this.txtAdValorem.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAdValorem.PromptChar = ' ';
			this.txtAdValorem.Size = new System.Drawing.Size(96, 22);
			this.txtAdValorem.TabIndex = 272;
			this.txtAdValorem.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAdValorem_KeyDown);
			this.txtAdValorem.Click += new System.EventHandler(this.txtAdValorem_Click);
			this.txtAdValorem.Validated += new System.EventHandler(this.txtAdValorem_Validated);
			this.txtAdValorem.ValueChanged += new System.EventHandler(this.txtAdValorem_ValueChanged);
			this.txtAdValorem.Enter += new System.EventHandler(this.txtAdValorem_Enter);
			// 
			// txtCIF
			// 
			this.txtCIF.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCIF.Appearance = appearance17;
			this.txtCIF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCIF.Enabled = false;
			this.txtCIF.FormatString = "#,##0.00";
			this.txtCIF.Location = new System.Drawing.Point(293, 568);
			this.txtCIF.Name = "txtCIF";
			this.txtCIF.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCIF.PromptChar = ' ';
			this.txtCIF.Size = new System.Drawing.Size(96, 22);
			this.txtCIF.TabIndex = 293;
			// 
			// txtSeguro
			// 
			this.txtSeguro.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSeguro.Appearance = appearance18;
			this.txtSeguro.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSeguro.Enabled = false;
			this.txtSeguro.FormatString = "#,##0.00";
			this.txtSeguro.Location = new System.Drawing.Point(293, 544);
			this.txtSeguro.Name = "txtSeguro";
			this.txtSeguro.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSeguro.PromptChar = ' ';
			this.txtSeguro.Size = new System.Drawing.Size(96, 22);
			this.txtSeguro.TabIndex = 292;
			// 
			// txtFleteExterior
			// 
			this.txtFleteExterior.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFleteExterior.Appearance = appearance19;
			this.txtFleteExterior.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFleteExterior.Enabled = false;
			this.txtFleteExterior.FormatString = "#,##0.00";
			this.txtFleteExterior.Location = new System.Drawing.Point(293, 520);
			this.txtFleteExterior.Name = "txtFleteExterior";
			this.txtFleteExterior.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFleteExterior.PromptChar = ' ';
			this.txtFleteExterior.Size = new System.Drawing.Size(96, 22);
			this.txtFleteExterior.TabIndex = 291;
			// 
			// txtFOBEXW
			// 
			this.txtFOBEXW.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFOBEXW.Appearance = appearance20;
			this.txtFOBEXW.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFOBEXW.Enabled = false;
			this.txtFOBEXW.FormatString = "#,##0.00";
			this.txtFOBEXW.Location = new System.Drawing.Point(293, 496);
			this.txtFOBEXW.Name = "txtFOBEXW";
			this.txtFOBEXW.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFOBEXW.PromptChar = ' ';
			this.txtFOBEXW.Size = new System.Drawing.Size(96, 22);
			this.txtFOBEXW.TabIndex = 290;
			// 
			// txtFOB
			// 
			this.txtFOB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFOB.Appearance = appearance21;
			this.txtFOB.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFOB.Enabled = false;
			this.txtFOB.FormatString = "#,##0.00";
			this.txtFOB.Location = new System.Drawing.Point(293, 472);
			this.txtFOB.Name = "txtFOB";
			this.txtFOB.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFOB.PromptChar = ' ';
			this.txtFOB.Size = new System.Drawing.Size(96, 22);
			this.txtFOB.TabIndex = 289;
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(8, 624);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 24);
			this.btnConsultar.TabIndex = 301;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 608);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1048, 8);
			this.groupBox1.TabIndex = 302;
			this.groupBox1.TabStop = false;
			// 
			// cmbTipoCalculo
			// 
			this.cmbTipoCalculo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoCalculo.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipoCalculo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbTipoCalculo.DisplayMember = "Calculo";
			this.cmbTipoCalculo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoCalculo.Location = new System.Drawing.Point(528, 328);
			this.cmbTipoCalculo.Name = "cmbTipoCalculo";
			this.cmbTipoCalculo.Size = new System.Drawing.Size(312, 104);
			this.cmbTipoCalculo.TabIndex = 305;
			this.cmbTipoCalculo.ValueMember = "idTipoCalculo";
			this.cmbTipoCalculo.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance22.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance22;
			ultraGridBand3.AddButtonCaption = "_ImportacionDet";
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 57;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 47;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 36;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 35;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 4;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Width = 85;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn12.Header.Caption = "N. Asiento";
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 85;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn13.Header.Caption = "Proveedor";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Width = 280;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance23;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "IVA";
			ultraGridColumn14.Header.VisiblePosition = 8;
			ultraGridColumn14.Width = 36;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance24;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Width = 85;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance25;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Saldo";
			ultraGridColumn16.Header.VisiblePosition = 10;
			ultraGridColumn16.Width = 85;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn17.Header.VisiblePosition = 11;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn18.Header.VisiblePosition = 12;
			ultraGridColumn18.Width = 180;
			ultraGridColumn19.Header.Caption = "Calculo";
			ultraGridColumn19.Header.VisiblePosition = 13;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn19.Width = 89;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance26;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 14;
			ultraGridColumn20.Width = 85;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance27;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Subtotal";
			ultraGridColumn21.Header.VisiblePosition = 15;
			ultraGridColumn21.Width = 85;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 16;
			ultraGridColumn22.Width = 25;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn22});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance28.ForeColor = System.Drawing.Color.Black;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance28;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance29.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance29;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance31.FontData.BoldAsString = "True";
			appearance31.FontData.Name = "Arial";
			appearance31.FontData.SizeInPoints = 10F;
			appearance31.ForeColor = System.Drawing.Color.White;
			appearance31.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance31;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance32.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance32.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance32;
			appearance33.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance33.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance33;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 136);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1032, 328);
			this.ultraGrid1.TabIndex = 303;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
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
			ultraDataColumn17.DataType = typeof(int);
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
			ultraDataColumn22.DataType = typeof(bool);
			ultraDataColumn22.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(448, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(49, 16);
			this.lblNumero.TabIndex = 307;
			this.lblNumero.Text = "NÚMERO";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance34;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(528, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(114, 22);
			this.txtNumero.TabIndex = 306;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumero_Validating);
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
			// txtIdFlete
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdFlete.Appearance = appearance35;
			this.txtIdFlete.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtIdFlete.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdFlete.Enabled = false;
			this.txtIdFlete.FormatString = "";
			this.txtIdFlete.Location = new System.Drawing.Point(8, 472);
			this.txtIdFlete.Name = "txtIdFlete";
			this.txtIdFlete.PromptChar = ' ';
			this.txtIdFlete.Size = new System.Drawing.Size(16, 20);
			this.txtIdFlete.TabIndex = 309;
			this.txtIdFlete.Visible = false;
			// 
			// txtIdAsiento1
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdAsiento1.Appearance = appearance36;
			this.txtIdAsiento1.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtIdAsiento1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento1.Enabled = false;
			this.txtIdAsiento1.FormatString = "";
			this.txtIdAsiento1.Location = new System.Drawing.Point(56, 472);
			this.txtIdAsiento1.Name = "txtIdAsiento1";
			this.txtIdAsiento1.PromptChar = ' ';
			this.txtIdAsiento1.Size = new System.Drawing.Size(16, 20);
			this.txtIdAsiento1.TabIndex = 311;
			this.txtIdAsiento1.Visible = false;
			// 
			// txtIdAsiento
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdAsiento.Appearance = appearance37;
			this.txtIdAsiento.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.FormatString = "";
			this.txtIdAsiento.Location = new System.Drawing.Point(32, 472);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 20);
			this.txtIdAsiento.TabIndex = 310;
			this.txtIdAsiento.Visible = false;
			// 
			// txtEstado
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstado.Appearance = appearance38;
			this.txtEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.FormatString = "";
			this.txtEstado.Location = new System.Drawing.Point(80, 472);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 20);
			this.txtEstado.TabIndex = 312;
			this.txtEstado.Visible = false;
			// 
			// mnuAsientos
			// 
			this.mnuAsientos.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAsientosFactura,
																																								this.mnuAsientosLiquidacion,
																																								this.mnuAsientosIva,
																																								this.mnuAsientosEgreso,
																																								this.mnuAsientosGastos});
			// 
			// mnuAsientosFactura
			// 
			this.mnuAsientosFactura.Index = 0;
			this.mnuAsientosFactura.Text = "De Factura Seleccionada";
			this.mnuAsientosFactura.Click += new System.EventHandler(this.mnuAsientosFactura_Click);
			// 
			// mnuAsientosLiquidacion
			// 
			this.mnuAsientosLiquidacion.Index = 1;
			this.mnuAsientosLiquidacion.Text = "De Liquidación";
			this.mnuAsientosLiquidacion.Click += new System.EventHandler(this.mnuAsientosLiquidacion_Click);
			// 
			// mnuAsientosIva
			// 
			this.mnuAsientosIva.Index = 2;
			this.mnuAsientosIva.Text = "De Liquidación Aduana";
			this.mnuAsientosIva.Click += new System.EventHandler(this.mnuAsientosIva_Click);
			// 
			// mnuAsientosEgreso
			// 
			this.mnuAsientosEgreso.Index = 3;
			this.mnuAsientosEgreso.Text = "De Egreso IVA Importación";
			this.mnuAsientosEgreso.Click += new System.EventHandler(this.mnuAsientosEgreso_Click);
			// 
			// mnuAsientosGastos
			// 
			this.mnuAsientosGastos.Index = 4;
			this.mnuAsientosGastos.Text = "De Gastos de Importación";
			this.mnuAsientosGastos.Click += new System.EventHandler(this.mnuAsientosGastos_Click);
			// 
			// btnActualizar
			// 
			this.btnActualizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnActualizar.CausesValidation = false;
			this.btnActualizar.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizar.Image")));
			this.btnActualizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizar.Location = new System.Drawing.Point(883, 624);
			this.btnActualizar.Name = "btnActualizar";
			this.btnActualizar.Size = new System.Drawing.Size(75, 24);
			this.btnActualizar.TabIndex = 313;
			this.btnActualizar.Text = "&Actualizar";
			this.btnActualizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnActualizar.Click += new System.EventHandler(this.btnActualizar_Click);
			// 
			// txtIdEgreso
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdEgreso.Appearance = appearance39;
			this.txtIdEgreso.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtIdEgreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdEgreso.Enabled = false;
			this.txtIdEgreso.FormatString = "";
			this.txtIdEgreso.Location = new System.Drawing.Point(104, 472);
			this.txtIdEgreso.Name = "txtIdEgreso";
			this.txtIdEgreso.PromptChar = ' ';
			this.txtIdEgreso.Size = new System.Drawing.Size(16, 20);
			this.txtIdEgreso.TabIndex = 314;
			this.txtIdEgreso.Visible = false;
			// 
			// lblPagoIVA
			// 
			this.lblPagoIVA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblPagoIVA.AutoSize = true;
			this.lblPagoIVA.BackColor = System.Drawing.Color.Transparent;
			this.lblPagoIVA.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPagoIVA.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(0)), ((System.Byte)(0)));
			this.lblPagoIVA.Location = new System.Drawing.Point(669, 568);
			this.lblPagoIVA.Name = "lblPagoIVA";
			this.lblPagoIVA.Size = new System.Drawing.Size(0, 23);
			this.lblPagoIVA.TabIndex = 315;
			this.lblPagoIVA.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(672, 11);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(38, 16);
			this.label18.TabIndex = 317;
			this.label18.Text = "FECHA";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance40;
			this.dtFecha.AutoSelect = false;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton7.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton7);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(736, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 316;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(88, 624);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 318;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 120);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1048, 8);
			this.groupBox2.TabIndex = 319;
			this.groupBox2.TabStop = false;
			// 
			// txtPallets
			// 
			this.txtPallets.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPallets.Appearance = appearance41;
			this.txtPallets.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPallets.Enabled = false;
			this.txtPallets.FormatString = "#,##0.00";
			this.txtPallets.Location = new System.Drawing.Point(525, 568);
			this.txtPallets.Name = "txtPallets";
			this.txtPallets.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPallets.PromptChar = ' ';
			this.txtPallets.Size = new System.Drawing.Size(96, 22);
			this.txtPallets.TabIndex = 320;
			this.txtPallets.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPallets_KeyDown);
			this.txtPallets.Click += new System.EventHandler(this.txtPallets_Click);
			this.txtPallets.Validated += new System.EventHandler(this.txtPallets_Validated);
			this.txtPallets.Enter += new System.EventHandler(this.txtPallets_Enter);
			// 
			// label23
			// 
			this.label23.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(437, 568);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(36, 16);
			this.label23.TabIndex = 321;
			this.label23.Text = "Pallets";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidAsiento2
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidAsiento2.Appearance = appearance42;
			this.txtidAsiento2.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtidAsiento2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidAsiento2.Enabled = false;
			this.txtidAsiento2.FormatString = "";
			this.txtidAsiento2.Location = new System.Drawing.Point(128, 472);
			this.txtidAsiento2.Name = "txtidAsiento2";
			this.txtidAsiento2.PromptChar = ' ';
			this.txtidAsiento2.Size = new System.Drawing.Size(16, 20);
			this.txtidAsiento2.TabIndex = 322;
			this.txtidAsiento2.Visible = false;
			// 
			// label24
			// 
			this.label24.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(669, 475);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(88, 16);
			this.label24.TabIndex = 323;
			this.label24.Text = "DIFERENCIA ISC";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiferenciaISC
			// 
			this.txtDiferenciaISC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiferenciaISC.Appearance = appearance43;
			this.txtDiferenciaISC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiferenciaISC.Enabled = false;
			this.txtDiferenciaISC.FormatString = "#,##0.00";
			this.txtDiferenciaISC.Location = new System.Drawing.Point(813, 472);
			this.txtDiferenciaISC.Name = "txtDiferenciaISC";
			this.txtDiferenciaISC.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiferenciaISC.PromptChar = ' ';
			this.txtDiferenciaISC.Size = new System.Drawing.Size(96, 22);
			this.txtDiferenciaISC.TabIndex = 324;
			this.txtDiferenciaISC.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiferenciaISC_KeyDown);
			this.txtDiferenciaISC.Click += new System.EventHandler(this.txtDiferenciaISC_Click);
			this.txtDiferenciaISC.Enter += new System.EventHandler(this.txtDiferenciaISC_Enter);
			// 
			// frmImportacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1048, 654);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.txtDiferenciaISC);
			this.Controls.Add(this.txtidAsiento2);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtPallets);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblPagoIVA);
			this.Controls.Add(this.txtIdEgreso);
			this.Controls.Add(this.btnActualizar);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdAsiento1);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.txtIdFlete);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.cmbTipoCalculo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btResumenImp);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btEditar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.txtSalidaCapitales);
			this.Controls.Add(this.lblSaldo);
			this.Controls.Add(this.txtTotalImportacion);
			this.Controls.Add(this.txtIce);
			this.Controls.Add(this.txtTotalGastos);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtFodInfa);
			this.Controls.Add(this.txtAdValorem);
			this.Controls.Add(this.txtCIF);
			this.Controls.Add(this.txtSeguro);
			this.Controls.Add(this.txtFleteExterior);
			this.Controls.Add(this.txtFOBEXW);
			this.Controls.Add(this.txtFOB);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtDiasViaje);
			this.Controls.Add(this.txtDiasDeProduccion);
			this.Controls.Add(this.dtDesAduanizacion);
			this.Controls.Add(this.dtLlegadaPuerto);
			this.Controls.Add(this.dtFechaEmbarque);
			this.Controls.Add(this.dtFechaFin);
			this.Controls.Add(this.dtIngresoBodega);
			this.Controls.Add(this.dtFechaIni);
			this.Controls.Add(this.btProgramacion);
			this.Controls.Add(this.btFactura);
			this.Controls.Add(this.btImportacionTipo);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.btAsiento);
			this.Name = "frmImportacion";
			this.Text = "Liquidación de Importaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmImportacion_KeyDown);
			this.Load += new System.EventHandler(this.frmImportacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDiasViaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeProduccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesAduanizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtLlegadaPuerto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEmbarque)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtIngresoBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSalidaCapitales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalImportacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIce)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFodInfa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAdValorem)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCIF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSeguro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFleteExterior)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFOBEXW)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFOB)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoCalculo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFlete)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdEgreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPallets)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidAsiento2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferenciaISC)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		int iTipoCliente = 0;
		bool bEdicion = false;
		private C1DataRow drSeteoF;
		CultureInfo us = new CultureInfo("en-US");

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idTipoCalculo"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Adicional"].CellActivation = Activation.AllowEdit;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NoObligatorio"].CellActivation = Activation.AllowEdit;
			}
			else
			{
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["idTipoCalculo"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["Adicional"].CellActivation = Activation.Disabled;
				this.ultraGrid1.DisplayLayout.Bands[0].Columns["NoObligatorio"].CellActivation = Activation.Disabled;
			}
		}

		private void frmImportacion_Load(object sender, System.EventArgs e)
		{
			Funcion.Resolucion(this, this.Height, this.Width);
			this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select * From ListaDeImportaciones() Order By Nombre");
			this.cmbTipoCalculo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idTipoCalculo, Calculo From TiposDeCalculo");
			this.cmbProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, "Select Top 1 idImportacionNombre From ImportacionNombre Order By idImportacionNombre Desc");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Calculo(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			e.Cell.Row.Cells["TotalS"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Adicional"].Value) + Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value);
				
			this.txtFOB.Value = 0.00m;
			this.txtFOBEXW.Value = 0.00m;
			this.txtFleteExterior.Value = 0.00m;
			this.txtSeguro.Value = 0.00m;
			this.txtSalidaCapitales.Value = 0.00m;
			this.txtTotalGastos.Value = 0.00m;
				
			#region Suma Grid
			string sSQL = "";
			decimal dImpSalCap = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idTipoImportacion"].Value == 1)
				{
					//if (Convert.ToDecimal(dr.Cells["Valores"].Value) == 0) 
					this.txtFOB.Value = Convert.ToDecimal(this.txtFOB.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);		 
					//else this.txtFOB.Value = Convert.ToDecimal(this.txtFOB.Value) + Convert.ToDecimal(dr.Cells["Valores"].Value);
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 2)
				{
					this.txtFOBEXW.Value = Convert.ToDecimal(this.txtFOBEXW.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}
					
				if ((int)dr.Cells["idTipoImportacion"].Value == 3)
				{
					this.txtFleteExterior.Value = Convert.ToDecimal(this.txtFleteExterior.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 4)
				{
					this.txtSeguro.Value = Convert.ToDecimal(this.txtSeguro.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}
				
				sSQL = string.Format("Select Exterior From ImportacionTipo Where idImportacionTipo = {0}", (int)dr.Cells["idTipoImportacion"].Value);
				bool bExterior = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);
				if (bExterior)
				{
					dImpSalCap = 0;
					dImpSalCap = Convert.ToDecimal(dr.Cells["TotalS"].Value) * 0.05m;
					this.txtSalidaCapitales.Value = Convert.ToDecimal(this.txtSalidaCapitales.Value) + dImpSalCap;
				}
			}
			#endregion Suma Grid

			#region CIF
			this.txtCIF.Value = Convert.ToDecimal(this.txtFOB.Value)
				+ Convert.ToDecimal(this.txtFOBEXW.Value)
				+ Convert.ToDecimal(this.txtFleteExterior.Value)
				+ Convert.ToDecimal(this.txtSeguro.Value);
			#endregion CIF

			#region Total Importacion
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{					
				this.txtTotalGastos.Value = Convert.ToDecimal(this.txtTotalGastos.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
			}

			this.txtTotalGastos.Value = Convert.ToDecimal(this.txtTotalGastos.Value)
				+ Convert.ToDecimal(this.txtAdValorem.Value) 
				+ Convert.ToDecimal(this.txtFodInfa.Value)
				+ Convert.ToDecimal(this.txtIce.Value) 
				+ Convert.ToDecimal(this.txtIva.Value)
				+ Convert.ToDecimal(this.txtSalidaCapitales.Value);
			#endregion Total Importacion

			#region Saldos
			decimal dSaldo = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{					
				if (dr.Cells["SaldoAnticipo"].Value != System.DBNull.Value) dSaldo = dSaldo + Convert.ToDecimal(dr.Cells["SaldoAnticipo"].Value);
			}

			this.lblSaldo.Text = "SALDO : " + dSaldo.ToString("#,##0.00");
			#endregion Saldos
		}

		private void Calculos()
		{
			this.txtFOB.Value = 0.00m;
			this.txtFOBEXW.Value = 0.00m;
			this.txtFleteExterior.Value = 0.00m;
			this.txtSeguro.Value = 0.00m;
			this.txtSalidaCapitales.Value = 0.00m;
			this.txtTotalGastos.Value = 0.00m;
							
			#region Suma Grid
			string sSQL = "";
			decimal dImpSalCap = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idTipoImportacion"].Value == 1)
				{
					this.txtFOB.Value = Convert.ToDecimal(this.txtFOB.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 2)
				{
					this.txtFOBEXW.Value = Convert.ToDecimal(this.txtFOBEXW.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}
					
				if ((int)dr.Cells["idTipoImportacion"].Value == 3)
				{
					this.txtFleteExterior.Value = Convert.ToDecimal(this.txtFleteExterior.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}

				if ((int)dr.Cells["idTipoImportacion"].Value == 4)
				{
					this.txtSeguro.Value = Convert.ToDecimal(this.txtSeguro.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
				}
				
				sSQL = string.Format("Select Exterior From ImportacionTipo Where idImportacionTipo = {0}", (int)dr.Cells["idTipoImportacion"].Value);
				bool bExterior = Funcion.bEscalarSQL(cdsSeteoF, sSQL, true);
				if (bExterior)
				{
					dImpSalCap = 0;
					dImpSalCap = Convert.ToDecimal(dr.Cells["TotalS"].Value) * 0.05m;
					this.txtSalidaCapitales.Value = Convert.ToDecimal(this.txtSalidaCapitales.Value) + dImpSalCap;
				}
			}
			#endregion Suma Grid

			#region CIF
			this.txtCIF.Value = Convert.ToDecimal(this.txtFOB.Value)
				+ Convert.ToDecimal(this.txtFOBEXW.Value)
				+ Convert.ToDecimal(this.txtFleteExterior.Value)
				+ Convert.ToDecimal(this.txtSeguro.Value);
			#endregion CIF

			#region Total Importacion
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{					
				this.txtTotalGastos.Value = Convert.ToDecimal(this.txtTotalGastos.Value) + Convert.ToDecimal(dr.Cells["TotalS"].Value);
			}

			this.txtTotalGastos.Value = Convert.ToDecimal(this.txtTotalGastos.Value)
				+ Convert.ToDecimal(this.txtAdValorem.Value) 
				+ Convert.ToDecimal(this.txtFodInfa.Value)
				+ Convert.ToDecimal(this.txtIce.Value) 
				+ Convert.ToDecimal(this.txtIva.Value)
				+ Convert.ToDecimal(this.txtSalidaCapitales.Value);
			#endregion Total Importacion

			#region Saldos
			decimal dSaldo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{					
				if (dr.Cells["SaldoAnticipo"].Value != System.DBNull.Value) dSaldo = dSaldo + Convert.ToDecimal(dr.Cells["SaldoAnticipo"].Value);
			}

			this.lblSaldo.Text = "SALDO : " + dSaldo.ToString("#,##0.00");
			#endregion Saldos
		}

    private void cmbProyecto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtAdValorem_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbProyecto_ValueChanged(object sender, System.EventArgs e)
		{
			bEdicion = false;

			ModificaGrilla(false);

			this.txtIdFlete.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtIdAsiento1.Value = 0;
			this.txtidAsiento2.Value = 0;
			this.txtIdEgreso.Value = 0;
			this.txtEstado.Value = 0;
			this.txtNumero.Text = "";
			this.dtFechaIni.Value = System.DBNull.Value;
			this.txtDiasDeProduccion.Value = 0;
			this.dtFechaFin.Value = System.DBNull.Value;
			this.dtFechaEmbarque.Value = System.DBNull.Value;
			this.txtDiasViaje.Value = 0;
			this.dtLlegadaPuerto.Value = System.DBNull.Value;
			this.dtDesAduanizacion.Value = System.DBNull.Value;
			this.dtIngresoBodega.Value = System.DBNull.Value;

			this.txtFOB.Value = 0.00m;
			this.txtFOBEXW.Value = 0.00m;
			this.txtFleteExterior.Value = 0.00m;
			this.txtSeguro.Value = 0.00m;
			this.txtCIF.Value = 0.00m;
			
			this.txtAdValorem.Value = 0.00m;
			this.txtFodInfa.Value = 0.00m;
			this.txtIce.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtPallets.Value = 0.00m;
			this.txtDiferenciaISC.Value = 0.00m;

			this.txtTotalImportacion.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtTotalImportacion.Value = 0.00m;

			this.lblEstado.Text = "";

			if (this.cmbProyecto.ActiveRow == null) return;

			Cursor = Cursors.WaitCursor;
			string sSQLMaestro = string.Format("Exec ConsultaImportacion {0}", (int)this.cmbProyecto.Value);
			SqlDataReader drMaestro = Funcion.rEscalarSQL(cdsSeteoF, sSQLMaestro, true);
			drMaestro.Read();
			if (drMaestro.HasRows)
			{
				this.txtIdFlete.Value = drMaestro.GetInt32(1);
				if (drMaestro.GetValue(2) != System.DBNull.Value) this.txtIdAsiento.Value = drMaestro.GetInt32(2);
				if (drMaestro.GetValue(3) != System.DBNull.Value) this.txtIdAsiento1.Value = drMaestro.GetInt32(3);
				if (drMaestro.GetValue(26) != System.DBNull.Value) this.txtIdEgreso.Value = drMaestro.GetInt32(26);
				this.txtEstado.Value = drMaestro.GetInt32(4);
				if (drMaestro.GetValue(5) != System.DBNull.Value) this.txtNumero.Text = drMaestro.GetString(5);
				if (drMaestro.GetValue(27) != System.DBNull.Value) this.dtFecha.Value = drMaestro.GetDateTime(27);
				if (drMaestro.GetValue(6) != System.DBNull.Value) this.dtFechaIni.Value = drMaestro.GetDateTime(6);
				if (drMaestro.GetValue(7) != System.DBNull.Value) this.txtDiasDeProduccion.Value = drMaestro.GetInt32(7);
				if (drMaestro.GetValue(8) != System.DBNull.Value) this.dtFechaFin.Value = drMaestro.GetDateTime(8);
				if (drMaestro.GetValue(9) != System.DBNull.Value) this.dtFechaEmbarque.Value = drMaestro.GetDateTime(9);
				if (drMaestro.GetValue(10) != System.DBNull.Value) this.txtDiasViaje.Value = drMaestro.GetInt32(10);
				if (drMaestro.GetValue(11) != System.DBNull.Value) this.dtLlegadaPuerto.Value = drMaestro.GetDateTime(11);
				if (drMaestro.GetValue(12) != System.DBNull.Value) this.dtDesAduanizacion.Value = drMaestro.GetDateTime(12);
				if (drMaestro.GetValue(13) != System.DBNull.Value) this.dtIngresoBodega.Value = drMaestro.GetDateTime(13);

				if ((int)this.txtEstado.Value == 1)
				{
					this.txtFOB.Value = drMaestro.GetDecimal(14);
					this.txtFOBEXW.Value = drMaestro.GetDecimal(15);
					this.txtFleteExterior.Value = drMaestro.GetDecimal(16);
					this.txtSeguro.Value = drMaestro.GetDecimal(17);
					this.txtCIF.Value = drMaestro.GetDecimal(18);
				}

				if (drMaestro.GetValue(19) != System.DBNull.Value) this.txtAdValorem.Value = drMaestro.GetDecimal(19);
				if (drMaestro.GetValue(20) != System.DBNull.Value) this.txtFodInfa.Value = drMaestro.GetDecimal(20);
				if (drMaestro.GetValue(21) != System.DBNull.Value) this.txtIce.Value = drMaestro.GetDecimal(21);
				if (drMaestro.GetValue(22) != System.DBNull.Value) this.txtIva.Value = drMaestro.GetDecimal(22);

				if ((int)this.txtEstado.Value == 1)
				{
					this.txtTotalImportacion.Value = drMaestro.GetDecimal(23);
					this.txtSalidaCapitales.Value = drMaestro.GetDecimal(24);
					this.txtTotalGastos.Value = drMaestro.GetDecimal(25);
				}

				if (drMaestro.GetValue(28) != System.DBNull.Value) this.txtPallets.Value = drMaestro.GetDecimal(28);
				if (drMaestro.GetValue(29) != System.DBNull.Value) this.txtidAsiento2.Value = drMaestro.GetInt32(29);
				if (drMaestro.GetValue(30) != System.DBNull.Value) this.txtDiferenciaISC.Value = drMaestro.GetDecimal(30);
				if (drMaestro.GetValue(31) != System.DBNull.Value) this.lblPagoIVA.Text = drMaestro.GetString(31);
			}
			drMaestro.Close();

			if ((int)this.txtEstado.Value == 0)
			{	
				this.ultraGrid1.DataSource = ultraDataSource1; 
				string sSQL = string.Format("Exec ConsultaDetalleImportacionPorLiquidar {0}", (int)this.cmbProyecto.Value);
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				Calculos();

				this.lblEstado.Text = "POR LIQUIDAR";

				this.btEditar.Enabled = true;
			}
			else if ((int)this.txtEstado.Value == 1)
			{
				this.ultraGrid1.DataSource = ultraDataSource1; 

				string sSQL = string.Format("Exec ConsultaDetalleImportacionLiquidada {0}", (int)this.cmbProyecto.Value);
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.lblEstado.Text = "LIQUIDADO";

				this.btEditar.Enabled = false;
			}

			this.btEditar.Enabled = true;

			if ((int)this.txtIdEgreso.Value == 0) this.lblPagoIVA.Text = "PAGO DE IVA PENDIENTE";
			else if ((int)this.txtIdEgreso.Value > 0) this.lblPagoIVA.Text = "IVA PAGADO";

			string sSQLGrupo = string.Format("Select dbo.GrupoClienteImportacion ({0})", (int)this.cmbProyecto.Value);
			iTipoCliente = Funcion.iEscalarSQL(cdsSeteoF, sSQLGrupo);

			if (iTipoCliente == 2) this.btProgramacion.Enabled = false;
			else if (iTipoCliente == 3) this.btProgramacion.Enabled = true;

			Cursor = Cursors.Default;
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			bEdicion = false;

			ModificaGrilla(false);

			this.txtIdFlete.Enabled = false;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento1.Enabled = false;
			this.txtEstado.Enabled = false;
			this.txtNumero.Enabled = false;
			this.dtFecha.Enabled = false;
			this.dtFechaIni.Enabled = false;
			this.txtDiasDeProduccion.Enabled = false;
			this.dtFechaFin.Enabled = false;
			this.dtFechaEmbarque.Enabled = false;
			this.txtDiasViaje.Enabled = false;
			this.dtLlegadaPuerto.Enabled = false;
			this.dtDesAduanizacion.Enabled = false;
			this.dtIngresoBodega.Enabled = false;
			
			this.txtAdValorem.Enabled = false;
			this.txtFodInfa.Enabled = false;
			this.txtIce.Enabled = false;
			this.txtIva.Enabled = false;
			this.txtPallets.Enabled = false;
			this.txtDiferenciaISC.Enabled = false;

			this.btnConsultar.Enabled = true;
			this.btnImprimir.Enabled = true;
			this.btEditar.Enabled = true;
			this.btGrabar.Enabled = false;
			this.btCancelar.Enabled = false;
			this.btResumenImp.Enabled = true;
			this.btProgramacion.Enabled = true;
			this.btAsiento.Enabled = true;

			int idImportacion = (int)this.cmbProyecto.Value;
			this.cmbProyecto.Value = System.DBNull.Value;
			this.cmbProyecto.Value = idImportacion;
			this.cmbProyecto.Enabled = true;
			this.cmbProyecto.Focus();
		}

		private void btEditar_Click(object sender, System.EventArgs e)
		{
			ModificaGrilla(true);

			this.cmbProyecto.Enabled = false;

			this.txtNumero.Enabled = true;
			this.dtFecha.Enabled = true;
			this.dtFechaIni.Enabled = true;
			if (iTipoCliente == 3) this.txtDiasDeProduccion.Enabled = true;
			this.dtFechaEmbarque.Enabled = true;
			this.txtDiasViaje.Enabled = true;
			this.dtDesAduanizacion.Enabled = true;
			this.dtIngresoBodega.Enabled = true;
						
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Asiento Where Borrar = 0 And idAsiento = {0}", (int)this.txtIdEgreso.Value)) == 0) 
			{
				this.txtAdValorem.Enabled = true;
				this.txtFodInfa.Enabled = true;
				this.txtIce.Enabled = true;
				this.txtIva.Enabled = true;
			}

			this.txtPallets.Enabled = true;
			this.txtDiferenciaISC.Enabled = true;
			
			this.btnConsultar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btEditar.Enabled = false;
			this.btGrabar.Enabled = true;
			this.btCancelar.Enabled = true;
			this.btResumenImp.Enabled = false;
			this.btProgramacion.Enabled = false;
			this.btAsiento.Enabled = false;

			bEdicion = true;

			this.txtNumero.Focus();
		}

		private void btResumenImp_Click(object sender, System.EventArgs e)
		{
			frmReporteImportacion RepImp = new frmReporteImportacion((int)this.cmbProyecto.Value);
			RepImp.ShowDialog();
		}

		private void btProgramacion_Click(object sender, System.EventArgs e)
		{
			if (this.cmbProyecto.ActiveRow == null)
			{
				MessageBox.Show("No existen Importaciones.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (iTipoCliente != 3) 
			{
				MessageBox.Show("Programación de pagos solo para Proveedores del Asia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.dtFechaIni.Value == System.DBNull.Value)	
			{
				MessageBox.Show("Ingrese las Fechas de la Importación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			using (ImportacionProgramacion miImportacionProgramacion = new ImportacionProgramacion())
			{
				miImportacionProgramacion.ShowDialog();
			}
		}

		private void btImportacionTipo_Click(object sender, System.EventArgs e)
		{
			using (ImportacionTipo ImportacionTipo = new ImportacionTipo())
			{
				if (DialogResult.OK == ImportacionTipo.ShowDialog())
				{
					if (this.cmbProyecto.ActiveRow == null) return;
									
					if ((int)this.txtEstado.Value == 0)			
					{
						this.ultraGrid1.DataSource = ultraDataSource1; 

						string sSQL = string.Format("Exec ConsultaDetalleImportacionPorLiquidar {0}", (int)this.cmbProyecto.Value);
						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
					}

					Calculos();
				}
			}
		}

		private void txtAdValorem_Validated(object sender, System.EventArgs e)
		{
			if (this.txtAdValorem.Value == System.DBNull.Value) this.txtAdValorem.Value = 0;

			Calculos();
		}

		private void txtFodInfa_Validated(object sender, System.EventArgs e)
		{
			if (this.txtFodInfa.Value == System.DBNull.Value) this.txtFodInfa.Value = 0;

			Calculos();
		}

		private void txtIce_Validated(object sender, System.EventArgs e)
		{
			if (this.txtIce.Value == System.DBNull.Value) this.txtIce.Value = 0;

			Calculos();
		}

		private void txtIva_Validated(object sender, System.EventArgs e)
		{
			if (this.txtIva.Value == System.DBNull.Value) this.txtIva.Value = 0;

			Calculos();
		}

		private void txtAdValorem_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtFodInfa.Focus();
		}

		private void txtFodInfa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtIce.Focus();
		}

		private void txtIce_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtIva.Focus();
		}

		private void txtIva_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtPallets.Focus();
		}

		private void txtAdValorem_Click(object sender, System.EventArgs e)
		{
			this.txtAdValorem.SelectAll();
		}

		private void txtAdValorem_Enter(object sender, System.EventArgs e)
		{
			this.txtAdValorem.SelectAll();
		}

		private void txtFodInfa_Click(object sender, System.EventArgs e)
		{
			this.txtFodInfa.SelectAll();
		}

		private void txtFodInfa_Enter(object sender, System.EventArgs e)
		{
			this.txtFodInfa.SelectAll();
		}

		private void txtIce_Click(object sender, System.EventArgs e)
		{
			this.txtIce.SelectAll();
		}

		private void txtIce_Enter(object sender, System.EventArgs e)
		{
			this.txtIce.SelectAll();
		}

		private void txtIva_Click(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void txtIva_Enter(object sender, System.EventArgs e)
		{
			this.txtIva.SelectAll();
		}

		private void frmImportacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void dtFechaIni_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if (this.dtFechaIni.Value == System.DBNull.Value) return;
			
			if (iTipoCliente == 2)
			{				
				DateTime FInicio = (DateTime)this.dtFechaIni.Value;
				this.dtFechaFin.Value = FInicio.AddDays((int)this.txtDiasDeProduccion.Value);

				DateTime FFin = (DateTime)this.dtFechaFin.Value;
				this.dtFechaEmbarque.CalendarInfo.MinDate = FFin;
				this.dtFechaEmbarque.Value = FFin.AddDays(2);

				if ((int)this.txtIdFlete.Value == 1) this.txtDiasViaje.Value = 30;
				else if ((int)this.txtIdFlete.Value == 2) this.txtDiasViaje.Value = 8;

				DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
				this.dtLlegadaPuerto.Value = FEmb.AddDays((int)this.txtDiasViaje.Value);

				DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
				this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;
				this.dtDesAduanizacion.Value = LlegPto.AddDays(5);

				DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
				this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
				this.dtIngresoBodega.Value = DesAd.AddDays(2);
			}
			else if (iTipoCliente == 3)
			{
				DateTime FInicio = (DateTime)this.dtFechaIni.Value;
				this.dtFechaFin.Value = FInicio.AddDays((int)this.txtDiasDeProduccion.Value);

				DateTime FFin = (DateTime)this.dtFechaFin.Value;
				this.dtFechaEmbarque.CalendarInfo.MinDate = FFin;
				this.dtFechaEmbarque.Value = FFin.AddDays(2);

				if ((int)this.txtIdFlete.Value == 1) this.txtDiasViaje.Value = 30;
				else if ((int)this.txtIdFlete.Value == 2) this.txtDiasViaje.Value = 8;

				DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
				this.dtLlegadaPuerto.Value = FEmb.AddDays((int)this.txtDiasViaje.Value);

				DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
				this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;
				this.dtDesAduanizacion.Value = LlegPto.AddDays(5);

				DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
				this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
				this.dtIngresoBodega.Value = DesAd.AddDays(2);
			}			
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}
			if (this.dtFechaIni.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaIni.Focus();
				return;
			}

			if (this.dtFechaEmbarque.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Embarque.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaEmbarque.Focus();
				return;
			}

			if (this.dtDesAduanizacion.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Desaduanización", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesAduanizacion.Focus();
				return;
			}

			if (this.dtIngresoBodega.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Ingreso a Bodega de la Mercaderia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIngresoBodega.Focus();
				return;
			}

//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
//			{	
//				if ((int)dr.Cells["idCompra"].Value == 0 && (bool)dr.Cells["NoObligatorio"].Value && Convert.ToDecimal(dr.Cells["Adicional"].Value > 0)
//				{
//					if ((int)dr.Cells["idTipoCalculo"].Value == 0)
//					{
//						MessageBox.Show("Seleccione la Forma de Calculo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.ultraGrid1.ActiveRow = dr;
//						this.ultraGrid1.ActiveRow.Selected = true;
//						return;
//					}
//				}
//			}
			#endregion Validacion

			Calculos();

			#region Estado
			int iCont = 0;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{	
				if ((int)dr.Cells["idCompra"].Value == 0 && (bool)dr.Cells["NoObligatorio"].Value) iCont++;
			}

			if (iCont == 0) 
			{
				if ((int)this.txtIdEgreso.Value == 0)
				{
					MessageBox.Show("No puede Terminar esta Liquidación si no se ha Pagado el I.V.A.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtEstado.Value = 0;
				}
				else 
				{
					if (DialogResult.No == MessageBox.Show("Desea Terminar esta Liquidación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) this.txtEstado.Value = 0;
					else this.txtEstado.Value = 1;
				}
			}
			#endregion Estado

			#region Maestro
			DateTime Fecha = (DateTime)this.dtFecha.Value;
			DateTime Inicio = (DateTime)this.dtFechaIni.Value;
			DateTime Fin = (DateTime)this.dtFechaFin.Value;
			DateTime Embarque = (DateTime)this.dtFechaEmbarque.Value;
			DateTime LlegadaPuerto = (DateTime)this.dtLlegadaPuerto.Value;
			DateTime Desaduanizacion = (DateTime)this.dtDesAduanizacion.Value;
			DateTime IngresoBodega = (DateTime)this.dtIngresoBodega.Value;

			string sSQLGuarda = string.Format("Exec GuardaImportacion {0}, {1}, '{2}', '{3}', {4}, '{5}', '{6}', {7}, '{8}', '{9}', '{10}', {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, '{23}', {24}, {25}", 
				(int)this.cmbProyecto.Value, (int)this.txtEstado.Value, this.txtNumero.Text.ToString(),
				Inicio.ToString("yyyyMMdd"), (int)this.txtDiasDeProduccion.Value, Fin.ToString("yyyyMMdd"),
				Embarque.ToString("yyyyMMdd"), (int)this.txtDiasViaje.Value, LlegadaPuerto.ToString("yyyyMMdd"),
				Desaduanizacion.ToString("yyyyMMdd"), IngresoBodega.ToString("yyyyMMdd"),
				Convert.ToDecimal(this.txtFOB.Value), Convert.ToDecimal(this.txtFOBEXW.Value), Convert.ToDecimal(this.txtFleteExterior.Value), Convert.ToDecimal(this.txtSeguro.Value), Convert.ToDecimal(this.txtCIF.Value),
				Convert.ToDecimal(this.txtAdValorem.Value), Convert.ToDecimal(this.txtFodInfa.Value), Convert.ToDecimal(this.txtIce.Value), Convert.ToDecimal(this.txtIva.Value), 
				Convert.ToDecimal(this.txtTotalImportacion.Value), Convert.ToDecimal(this.txtSalidaCapitales.Value), 
				Convert.ToDecimal(this.txtTotalGastos.Value), Fecha.ToString("yyyyMMdd"), Convert.ToDecimal(this.txtPallets.Value), Convert.ToDecimal(this.txtDiferenciaISC.Value));
			Funcion.EjecutaSQL(cdsSeteoF, sSQLGuarda);  
			#endregion Maestro

			#region Detalle
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From ImportacionDet Where idImportacion = {0}", (int)this.cmbProyecto.Value));

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{					
				string sSQLDetalle = string.Format("Exec GuardaDetalleImportacion {0}, {1}, {2}, {3}, {4}, {5}, {6}",
					(int)this.cmbProyecto.Value, (int)dr.Cells["idCompra"].Value, (int)dr.Cells["idTipoImportacion"].Value, (int)dr.Cells["idTipoCalculo"].Value, 
					Convert.ToDecimal(dr.Cells["Adicional"].Value), Convert.ToDecimal(dr.Cells["TotalS"].Value), (bool)dr.Cells["NoObligatorio"].Value);				
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}
//			string sSQL = string.Format("Exec ConsultaDetalleImportacionLiquidada {0}", (int)this.cmbProyecto.Value);
//			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			#endregion Detalle

			string sSQLReporte = string.Format("Exec CreaResumenDeImportaciones {0}", (int)this.cmbProyecto.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQLReporte);
			
			#region Actualiza Fecha de llegada en Orden de Compra
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((int)dr.Cells["idTipoImportacion"].Value == 1)
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set FechaTentativaLlegada = '{0}' Where idTipoFactura = 2 And idCompra = {1}", 
						Convert.ToDateTime(this.dtIngresoBodega.Value).ToString("yyyyMMdd"), (int)dr.Cells["idCompra"].Value));				
			}
			#endregion Actualiza Fecha de llegada en Orden de Compra
			
			#region Asientos
//			string sSQLAsiento = string.Format("Exec AsientoLiquidacionImportacion {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, '{10}'",
//				(int)this.txtIdAsiento.Value, (int)this.cmbProyecto.Value, this.cmbProyecto.Text.ToString().Trim(), this.txtNumero.Text.ToString(), 19, 
//				Convert.ToDecimal(this.txtTotalGastos.Value, Convert.ToDecimal(this.txtFodInfa.Value, Convert.ToDecimal(this.txtAdValorem.Value, Convert.ToDecimal(this.txtSalidaCapitales.Value, Convert.ToDecimal(this.txtIva.Value, Fecha.ToString("yyyyMMdd"));
//			this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);
			
			if (Convert.ToDecimal(this.txtSalidaCapitales.Value) > 0.00m)
			{
				string sSQLAsiento = string.Format("Exec AsientoCostoDeVentasImportacion {0}, {1}, '{2}', '{3}', '{4}', {5}",
					(int)this.txtIdAsiento.Value, (int)this.cmbProyecto.Value, this.txtNumero.Text.ToString(), Fecha.ToString("yyyyMMdd"), this.cmbProyecto.Text.ToString().Trim(), Convert.ToDecimal(this.txtSalidaCapitales.Value));
				this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);			
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From Asiento Where Borrar = 0 And idAsiento = {0}", (int)this.txtIdEgreso.Value)) == 0)
			{
				if ((Math.Round(Convert.ToDecimal(this.txtAdValorem.Value), 2) + Math.Round(Convert.ToDecimal(this.txtFodInfa.Value), 2) + Math.Round(Convert.ToDecimal(this.txtIva.Value), 2)) > 0.00m)
				{
					string sSQLAsientoImpuestos = string.Format("Exec AsientoImpuestosImportacion {0}, {1}, '{2}', '{3}', {4}, {5}, {6}, {7}, '{8}', {9}", 
						(int)this.txtIdAsiento1.Value, (int)this.cmbProyecto.Value, this.cmbProyecto.Text.ToString().Trim(), this.txtNumero.Text.ToString(), 19, 
						Convert.ToDecimal(this.txtAdValorem.Value), Convert.ToDecimal(this.txtFodInfa.Value), Convert.ToDecimal(this.txtIva.Value), Fecha.ToString("yyyyMMdd"), Convert.ToDecimal(this.txtIce.Value));
					this.txtIdAsiento1.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsientoImpuestos);				
				}
			
				#region Autorización de Pagos
				decimal dValor = Math.Round(Convert.ToDecimal(this.txtAdValorem.Value), 2) + 
					Math.Round(Convert.ToDecimal(this.txtFodInfa.Value), 2) + 
					Math.Round(Convert.ToDecimal(this.txtIce.Value), 2) + 
					Math.Round(Convert.ToDecimal(this.txtIva.Value), 2);

				string sSQLAut = string.Format("Exec CreaAutorizacionesDePago 11, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'LIQUIDACION DE IVA DE IMPORTACION {4} NUMERO : {5}', 0", 
					(int)this.cmbProyecto.Value, Fecha.ToString("yyyyMMdd"), Fecha.ToString("yyyyMMdd"), dValor, this.cmbProyecto.Text.ToString(), this.txtNumero.Text.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAut, true);
				#endregion Autorización de Pagos			
			}

			if ((int)this.txtEstado.Value == 1)
			{
				string sSQLGI = string.Format("Exec AsientoGastosImportacion {0}", (int)this.cmbProyecto.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLGI, true);
			}
			#endregion Asientos

			this.btCancelar_Click(sender, e);
			
			Cursor = Cursors.Default;
		}

		private void dtFechaIni_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter)
			{
				if (iTipoCliente == 2) this.dtFechaEmbarque.Focus();	
				else if (iTipoCliente == 3) this.txtDiasDeProduccion.Focus();			
			}
		}

		private void dtFechaIni_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion)return;

			if (this.dtFechaIni.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Inicio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaIni.Focus();
				e.Cancel = true;
			}
		}

		private void txtDiasDeProduccion_Click(object sender, System.EventArgs e)
		{
			this.txtDiasDeProduccion.SelectAll();
		}

		private void txtDiasDeProduccion_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasDeProduccion.SelectAll();
		}

		private void txtDiasDeProduccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)	this.dtFechaEmbarque.Focus();
		}

		private void txtDiasDeProduccion_Validated(object sender, System.EventArgs e)
		{
			if (this.txtDiasDeProduccion.Value == System.DBNull.Value) this.txtDiasDeProduccion.Value = 0;
			this.dtFechaIni_Validated(sender, e);
		}

		private void dtFechaEmbarque_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiasViaje.Focus();
		}

		private void dtFechaEmbarque_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if (this.dtFechaEmbarque.Value == System.DBNull.Value) return;

			if ((int)this.txtIdFlete.Value == 1) this.txtDiasViaje.Value = 30;
			else if ((int)this.txtIdFlete.Value == 2) this.txtDiasViaje.Value = 8;
						
			DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;						
			this.dtLlegadaPuerto.Value = FEmb.AddDays((int)this.txtDiasViaje.Value);

			DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
			
			this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;
			this.dtDesAduanizacion.Value = LlegPto.AddDays(5);
			
			DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
			this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
			this.dtIngresoBodega.Value = DesAd.AddDays(2);
		}

		private void dtFechaEmbarque_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (this.dtFechaEmbarque.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Embarque.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaEmbarque.Focus();
				e.Cancel = true;
			}
		}

		private void txtDiasViaje_Click(object sender, System.EventArgs e)
		{
			this.txtDiasViaje.SelectAll();
		}

		private void txtDiasViaje_Enter(object sender, System.EventArgs e)
		{
			this.txtDiasViaje.SelectAll();
		}

		private void txtDiasViaje_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtDesAduanizacion.Focus();
		}

		private void txtDiasViaje_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			DateTime FEmb = (DateTime)this.dtFechaEmbarque.Value;
			this.dtLlegadaPuerto.Value = FEmb.AddDays((int)this.txtDiasViaje.Value);

			DateTime LlegPto = (DateTime)this.dtLlegadaPuerto.Value;
			this.dtDesAduanizacion.CalendarInfo.MinDate = LlegPto;	
			this.dtDesAduanizacion.Value = LlegPto.AddDays(5);

			DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
			this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
			this.dtIngresoBodega.Value = DesAd.AddDays(2);
		}

		private void txtDiasViaje_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (this.txtDiasViaje.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese los Dias de Viaje", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDiasViaje.Focus();
				e.Cancel = true;
			}
		}

		private void dtDesAduanizacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.dtIngresoBodega.Focus();
		}

		private void dtDesAduanizacion_Validated(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if (this.dtDesAduanizacion.Value == System.DBNull.Value) return;

			DateTime DesAd = (DateTime)this.dtDesAduanizacion.Value;
			this.dtIngresoBodega.CalendarInfo.MinDate = DesAd;
			this.dtIngresoBodega.Value = DesAd.AddDays(2);		
		}

		private void dtDesAduanizacion_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (this.dtDesAduanizacion.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Desaduanización", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtDesAduanizacion.Focus();
				e.Cancel = true;
			}
		}

		private void dtIngresoBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (!bEdicion) return;

			if (this.dtIngresoBodega.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha de Ingreso a Bodega de la Mercaderia.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtIngresoBodega.Focus();
				e.Cancel = true;
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Adicional") Calculo(e);
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
			#endregion Formato numerico

			#region Decimales
			e.Layout.Bands[0].Columns["Adicional"].Format = stFormato;
			e.Layout.Bands[0].Columns["Adicional"].MaskInput = stInput;
			#endregion Decimales
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;

			try
			{				
				switch(e.KeyValue)
				{
						#region Flechas
					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
						#endregion Flechas
						#region Enter
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error del Sistema");
			}
		}

		private void dtIngresoBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtAdValorem.Focus();		
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!(bool)this.ultraGrid1.ActiveRow.Cells["NoObligatorio"].Value)
			{
				MessageBox.Show("La Fila seleccionada no esta seleccionada para esta Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion

			#region Abre Pantalla de Compra Existente o Nueva u Orden de Compra 
			int IdCompra = 0;

			if (this.ultraGrid1.ActiveRow.Cells["idCompra"].Value != System.DBNull.Value) 
			{
				IdCompra = (int)this.ultraGrid1.ActiveRow.Cells["idCompra"].Value;
				int IdTipoFactura = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTipoFactura From Compra Where idCompra = {0}", IdCompra));

				if (IdTipoFactura == 2)
				{
					frmOrdenesDeCompra OrdenDeCompra = new frmOrdenesDeCompra(IdCompra);
					OrdenDeCompra.MdiParent = MdiParent;
					OrdenDeCompra.Show();
				}
				else if (IdTipoFactura == 4)
				{
					frmCompra Compra = new frmCompra(IdCompra, true);
					Compra.MdiParent = MdiParent;
					Compra.Show();
				}
			}
			else
			{
				
			}
			#endregion Abre Pantalla de Compra Existente o Nueva u Orden de Compra
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmBuscaLiqImportaciones Busqueda = new frmBuscaLiqImportaciones())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					try
					{							
						this.cmbProyecto.Value = System.DBNull.Value;
						this.cmbProyecto.Value = (int)Busqueda.ultraGrid1.ActiveRow.Cells["idImportacion"].Value;
					}
					catch(Exception Exc)
					{
						MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}		
		}

		private void mnuAsientosFactura_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione una Fila", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);			
			}
			else if ((int)this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value == 0)
			{
				MessageBox.Show("No existe Asiento del Documento Seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(((int)this.ultraGrid1.ActiveRow.Cells["idAsiento"].Value));
				Asiento.MdiParent = MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;		
			}
		}

		private void mnuAsientosLiquidacion_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento.Value == 0)
			{
				MessageBox.Show("No existe Asiento de esta Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(((int)this.txtIdAsiento.Value));
				Asiento.MdiParent = MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;		
			}
		}

		private void mnuAsientosIva_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdAsiento1.Value == 0)
			{
				MessageBox.Show("No existe Asiento de I.V.A. de esta Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(((int)this.txtIdAsiento1.Value));
				Asiento.MdiParent = MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;		
			}
		}

		private void btnActualizar_Click(object sender, System.EventArgs e)
		{
			try
			{	
				int IdImportacion = (int)this.cmbProyecto.Value;
				this.cmbProyecto.Value = System.DBNull.Value;
				this.cmbProyecto.Value = IdImportacion;
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Actualizar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			this.mnuAsientos.Show(this.btAsiento, new Point(1, 25));			
		}

		private void mnuAsientosEgreso_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtIdEgreso.Value == 0)
			{
				MessageBox.Show("No existe Egreso de I.V.A. de esta Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(((int)this.txtIdEgreso.Value));
				Asiento.MdiParent = MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;		
			}
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.dtFechaIni.Focus();
		}

		private void txtNumero_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtNumero.Text.Length == 0)
				{
					MessageBox.Show("Escriba el numero de liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}

				string sSQL = string.Format("Select Count(*) From Importacion Where Numero = '{1}' And idImportacion <> {0}", (int)this.cmbProyecto.Value, this.txtNumero.Text.ToString());
				int iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCont > 0)
				{
					MessageBox.Show("El numero de liquidación " + this.txtNumero.Text.ToString() + " ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumero.Focus();
					return;
				}
			}
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha de Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					return;
				}
			}
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldDesde = new ParameterField ();
			ParameterDiscreteValue discreteValDesde = new ParameterDiscreteValue ();
			paramFieldDesde.ParameterFieldName = "@idImportacion";
			discreteValDesde.Value = (int)this.cmbProyecto.Value;
			paramFieldDesde.CurrentValues.Add (discreteValDesde);
			paramFields.Add (paramFieldDesde);

			Reporte Reporte = new Reporte("LiquidacionImportaciones", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void txtIva_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPallets_Validated(object sender, System.EventArgs e)
		{
			if (this.txtPallets.Value == System.DBNull.Value) this.txtPallets.Value = 0;
		}

		private void txtPallets_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiferenciaISC.Focus();
		}

		private void dtFechaIni_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void mnuAsientosGastos_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtidAsiento2.Value == 0)
			{
				MessageBox.Show("No existe Asiento de Gastos de esta Liquidación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(((int)this.txtidAsiento2.Value));
				Asiento.MdiParent = MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;		
			}		
		}

		private void txtDiferenciaISC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btGrabar.Focus();
		}

		private void txtPallets_Click(object sender, System.EventArgs e)
		{
			this.txtPallets.SelectAll();
		}

		private void txtPallets_Enter(object sender, System.EventArgs e)
		{
			this.txtPallets.SelectAll();
		}

		private void txtDiferenciaISC_Enter(object sender, System.EventArgs e)
		{
			this.txtDiferenciaISC.SelectAll();
		}

		private void txtDiferenciaISC_Click(object sender, System.EventArgs e)
		{
			this.txtDiferenciaISC.SelectAll();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFechaIni_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFechaEmbarque_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtDesAduanizacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}
	}
}

