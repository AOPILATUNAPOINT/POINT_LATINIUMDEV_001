using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmConciliacionBancaria.
	/// </summary>
	public class frmConciliacionBancaria : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEntrada;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSistemaEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSistemaSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConciliadosEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConciliadosSalidas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNoConciliadosSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNoConciliadosEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiferencias;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoBancario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoConciliado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoNoConciliado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoEnLibros;
		private System.Windows.Forms.Button btnNuevo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdConciliacionBancaria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCuenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBanco;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuenta;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnBloquear;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.MenuItem mnuExportarEntradas;
		private System.Windows.Forms.MenuItem mnuExportarSalidas;
		private System.Windows.Forms.ContextMenu mnuExportar;
		private System.Windows.Forms.ContextMenu mnuMovimiento;
		private System.Windows.Forms.MenuItem mnuMovimientoEntrada;
		private System.Windows.Forms.MenuItem mnuMovimientoSalida;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSelIngresos;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConciliacionBancaria()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int IdCuenta = 0;
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		int idConciliacionBancaria = 0;
		int idBloqueaTransacciones = 0;

		public frmConciliacionBancaria(int IdConciliacionBancaria)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			
			idConciliacionBancaria = IdConciliacionBancaria;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

//		public frmConciliacionBancaria(int idCuenta)
//		{
//			//
//			// Required for Windows Form Designer support
//			//
//			InitializeComponent();
//			
//			IdCuenta = idCuenta;
//			//
//			// TODO: Add any constructor code after InitializeComponent call
//			//
//		}

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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConciliacionBancaria));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConciliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaConciliacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConciliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Origen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cheque");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaConciliacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Confirmado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.lblNumero = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.uGridEntrada = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSalida = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSistemaEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtSistemaSalida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConciliadosSalidas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConciliadosEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNoConciliadosSalida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNoConciliadosEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.txtDiferencias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoBancario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoConciliado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoNoConciliado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtSaldoEnLibros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.txtIdConciliacionBancaria = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCuenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBanco = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnBloquear = new System.Windows.Forms.Button();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.mnuExportar = new System.Windows.Forms.ContextMenu();
			this.mnuExportarEntradas = new System.Windows.Forms.MenuItem();
			this.mnuExportarSalidas = new System.Windows.Forms.MenuItem();
			this.mnuMovimiento = new System.Windows.Forms.ContextMenu();
			this.mnuMovimientoEntrada = new System.Windows.Forms.MenuItem();
			this.mnuMovimientoSalida = new System.Windows.Forms.MenuItem();
			this.chkSelIngresos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBancario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoConciliado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNoConciliado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoEnLibros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConciliacionBancaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(255, 576);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 42;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(526, 576);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 44;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(810, 576);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 45;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(889, 576);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 46;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 91);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 16);
			this.label3.TabIndex = 256;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblNumero
			// 
			this.lblNumero.AutoSize = true;
			this.lblNumero.BackColor = System.Drawing.Color.Transparent;
			this.lblNumero.Location = new System.Drawing.Point(840, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 254;
			this.lblNumero.Text = "Número";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 43);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(20, 16);
			this.label6.TabIndex = 253;
			this.label6.Text = "Del";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(904, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 245;
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 40);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 249;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// txtBuscar
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance3;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Enabled = false;
			this.txtBuscar.Location = new System.Drawing.Point(64, 88);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(248, 22);
			this.txtBuscar.TabIndex = 247;
			this.txtBuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscar_KeyPress);
			this.txtBuscar.TextChanged += new System.EventHandler(this.txtBuscar_TextChanged);
			this.txtBuscar.ValueChanged += new System.EventHandler(this.txtBuscar_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(208, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(14, 16);
			this.label1.TabIndex = 258;
			this.label1.Text = "Al";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label1.Click += new System.EventHandler(this.label1_Click);
			// 
			// dtHasta
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance4;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 257;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1024, 8);
			this.groupBox1.TabIndex = 259;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 560);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1024, 8);
			this.groupBox2.TabIndex = 260;
			this.groupBox2.TabStop = false;
			// 
			// uGridEntrada
			// 
			this.uGridEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEntrada.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntrada.DataSource = this.ultraDataSource1;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridEntrada.DisplayLayout.Appearance = appearance5;
			this.uGridEntrada.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 95;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 115;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 70;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 50;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Tipo";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 50;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Número";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Concepto";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 280;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.Format = "dd/MM/yyyy HH:MM";
			ultraGridColumn9.Header.Caption = "Conciliado";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 92;
			ultraGridColumn10.Header.Caption = "...";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn10.Width = 25;
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
																										 ultraGridColumn10});
			ultraGridBand1.Header.Caption = "Movimientos de Entrada";
			ultraGridBand1.HeaderVisible = true;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridEntrada.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntrada.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.uGridEntrada.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridEntrada.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntrada.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8.25F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntrada.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.uGridEntrada.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.uGridEntrada.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntrada.Location = new System.Drawing.Point(8, 128);
			this.uGridEntrada.Name = "uGridEntrada";
			this.uGridEntrada.Size = new System.Drawing.Size(768, 208);
			this.uGridEntrada.TabIndex = 607;
			this.uGridEntrada.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridEntrada.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntrada_CellChange);
			this.uGridEntrada.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntrada_AfterCellUpdate);
			this.uGridEntrada.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridEntrada_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// uGridSalida
			// 
			this.uGridSalida.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSalida.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSalida.DataSource = this.ultraDataSource3;
			appearance14.BackColor = System.Drawing.Color.White;
			this.uGridSalida.DisplayLayout.Appearance = appearance14;
			this.uGridSalida.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 95;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 115;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 70;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Tipo";
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Width = 50;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Número";
			ultraGridColumn16.Header.VisiblePosition = 5;
			ultraGridColumn16.Width = 90;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Concepto";
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Width = 280;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance15;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.Format = "dd/MM/yyyy HH:MM";
			ultraGridColumn19.Header.Caption = "Conciliado";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Width = 92;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Width = 25;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			ultraGridBand2.Header.Caption = "Movimientos de Salida";
			ultraGridBand2.HeaderVisible = true;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance16;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridSalida.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSalida.DisplayLayout.Override.ActiveRowAppearance = appearance17;
			this.uGridSalida.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSalida.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridSalida.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.uGridSalida.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 8.25F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSalida.DisplayLayout.Override.HeaderAppearance = appearance19;
			this.uGridSalida.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalida.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalida.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSalida.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridSalida.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSalida.Location = new System.Drawing.Point(8, 344);
			this.uGridSalida.Name = "uGridSalida";
			this.uGridSalida.Size = new System.Drawing.Size(768, 208);
			this.uGridSalida.TabIndex = 608;
			this.uGridSalida.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			this.uGridSalida.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalida_CellChange);
			this.uGridSalida.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSalida_AfterCellUpdate);
			this.uGridSalida.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSalida_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.DateTime);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.DateTime);
			ultraDataColumn20.DataType = typeof(bool);
			ultraDataColumn20.DefaultValue = false;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(840, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 643;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Location = new System.Drawing.Point(784, 112);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 645;
			this.label2.Text = "Entradas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSistemaEntrada
			// 
			this.txtSistemaEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSistemaEntrada.Appearance = appearance23;
			this.txtSistemaEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSistemaEntrada.Enabled = false;
			this.txtSistemaEntrada.FormatString = "#,##0.00";
			this.txtSistemaEntrada.Location = new System.Drawing.Point(904, 112);
			this.txtSistemaEntrada.Name = "txtSistemaEntrada";
			this.txtSistemaEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSistemaEntrada.PromptChar = ' ';
			this.txtSistemaEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtSistemaEntrada.TabIndex = 644;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Location = new System.Drawing.Point(784, 136);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 16);
			this.label4.TabIndex = 647;
			this.label4.Text = "Salidas";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSistemaSalida
			// 
			this.txtSistemaSalida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSistemaSalida.Appearance = appearance24;
			this.txtSistemaSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSistemaSalida.Enabled = false;
			this.txtSistemaSalida.FormatString = "#,##0.00";
			this.txtSistemaSalida.Location = new System.Drawing.Point(904, 136);
			this.txtSistemaSalida.Name = "txtSistemaSalida";
			this.txtSistemaSalida.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSistemaSalida.PromptChar = ' ';
			this.txtSistemaSalida.Size = new System.Drawing.Size(112, 22);
			this.txtSistemaSalida.TabIndex = 646;
			// 
			// txtConciliadosSalidas
			// 
			this.txtConciliadosSalidas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConciliadosSalidas.Appearance = appearance25;
			this.txtConciliadosSalidas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConciliadosSalidas.Enabled = false;
			this.txtConciliadosSalidas.FormatString = "#,##0.00";
			this.txtConciliadosSalidas.Location = new System.Drawing.Point(904, 224);
			this.txtConciliadosSalidas.Name = "txtConciliadosSalidas";
			this.txtConciliadosSalidas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtConciliadosSalidas.PromptChar = ' ';
			this.txtConciliadosSalidas.Size = new System.Drawing.Size(112, 22);
			this.txtConciliadosSalidas.TabIndex = 650;
			// 
			// txtConciliadosEntrada
			// 
			this.txtConciliadosEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConciliadosEntrada.Appearance = appearance26;
			this.txtConciliadosEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConciliadosEntrada.Enabled = false;
			this.txtConciliadosEntrada.FormatString = "#,##0.00";
			this.txtConciliadosEntrada.Location = new System.Drawing.Point(904, 200);
			this.txtConciliadosEntrada.Name = "txtConciliadosEntrada";
			this.txtConciliadosEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtConciliadosEntrada.PromptChar = ' ';
			this.txtConciliadosEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtConciliadosEntrada.TabIndex = 648;
			// 
			// txtNoConciliadosSalida
			// 
			this.txtNoConciliadosSalida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNoConciliadosSalida.Appearance = appearance27;
			this.txtNoConciliadosSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNoConciliadosSalida.Enabled = false;
			this.txtNoConciliadosSalida.FormatString = "#,##0.00";
			this.txtNoConciliadosSalida.Location = new System.Drawing.Point(904, 312);
			this.txtNoConciliadosSalida.Name = "txtNoConciliadosSalida";
			this.txtNoConciliadosSalida.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNoConciliadosSalida.PromptChar = ' ';
			this.txtNoConciliadosSalida.Size = new System.Drawing.Size(112, 22);
			this.txtNoConciliadosSalida.TabIndex = 654;
			// 
			// txtNoConciliadosEntrada
			// 
			this.txtNoConciliadosEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNoConciliadosEntrada.Appearance = appearance28;
			this.txtNoConciliadosEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNoConciliadosEntrada.Enabled = false;
			this.txtNoConciliadosEntrada.FormatString = "#,##0.00";
			this.txtNoConciliadosEntrada.Location = new System.Drawing.Point(904, 288);
			this.txtNoConciliadosEntrada.Name = "txtNoConciliadosEntrada";
			this.txtNoConciliadosEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNoConciliadosEntrada.PromptChar = ' ';
			this.txtNoConciliadosEntrada.Size = new System.Drawing.Size(112, 22);
			this.txtNoConciliadosEntrada.TabIndex = 652;
			// 
			// label10
			// 
			this.label10.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label10.Location = new System.Drawing.Point(784, 88);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(56, 17);
			this.label10.TabIndex = 656;
			this.label10.Text = "SISTEMA";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label11.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Location = new System.Drawing.Point(784, 168);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(86, 17);
			this.label11.TabIndex = 657;
			this.label11.Text = "CONCILIADOS";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label12.Location = new System.Drawing.Point(784, 256);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(107, 17);
			this.label12.TabIndex = 658;
			this.label12.Text = "NO CONCILIADOS";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Location = new System.Drawing.Point(784, 224);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(39, 16);
			this.label5.TabIndex = 660;
			this.label5.Text = "Salidas";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label7.Location = new System.Drawing.Point(784, 200);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(48, 16);
			this.label7.TabIndex = 659;
			this.label7.Text = "Entradas";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label8.Location = new System.Drawing.Point(784, 312);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 16);
			this.label8.TabIndex = 662;
			this.label8.Text = "Salidas";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Location = new System.Drawing.Point(784, 288);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(48, 16);
			this.label9.TabIndex = 661;
			this.label9.Text = "Entradas";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Location = new System.Drawing.Point(784, 488);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(59, 16);
			this.label13.TabIndex = 674;
			this.label13.Text = "Diferencias";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Location = new System.Drawing.Point(784, 464);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(78, 16);
			this.label14.TabIndex = 673;
			this.label14.Text = "Saldo Bancario";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Location = new System.Drawing.Point(784, 440);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(86, 16);
			this.label15.TabIndex = 672;
			this.label15.Text = "Saldo Conciliado";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label16
			// 
			this.label16.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Location = new System.Drawing.Point(784, 416);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(103, 16);
			this.label16.TabIndex = 671;
			this.label16.Text = "Saldo No Conciliado";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label18.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label18.Location = new System.Drawing.Point(784, 360);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(92, 17);
			this.label18.TabIndex = 669;
			this.label18.Text = "CONCILIACIÓN";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiferencias
			// 
			this.txtDiferencias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiferencias.Appearance = appearance29;
			this.txtDiferencias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiferencias.Enabled = false;
			this.txtDiferencias.FormatString = "#,##0.00";
			this.txtDiferencias.Location = new System.Drawing.Point(904, 488);
			this.txtDiferencias.Name = "txtDiferencias";
			this.txtDiferencias.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiferencias.PromptChar = ' ';
			this.txtDiferencias.Size = new System.Drawing.Size(112, 22);
			this.txtDiferencias.TabIndex = 668;
			// 
			// txtSaldoBancario
			// 
			this.txtSaldoBancario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoBancario.Appearance = appearance30;
			this.txtSaldoBancario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoBancario.Enabled = false;
			this.txtSaldoBancario.FormatString = "#,##0.00";
			this.txtSaldoBancario.Location = new System.Drawing.Point(904, 464);
			this.txtSaldoBancario.Name = "txtSaldoBancario";
			this.txtSaldoBancario.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoBancario.PromptChar = ' ';
			this.txtSaldoBancario.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoBancario.TabIndex = 667;
			this.txtSaldoBancario.Validated += new System.EventHandler(this.txtSaldoBancario_Validated);
			this.txtSaldoBancario.ValueChanged += new System.EventHandler(this.txtSaldoBancario_ValueChanged);
			// 
			// txtSaldoConciliado
			// 
			this.txtSaldoConciliado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoConciliado.Appearance = appearance31;
			this.txtSaldoConciliado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoConciliado.Enabled = false;
			this.txtSaldoConciliado.FormatString = "#,##0.00";
			this.txtSaldoConciliado.Location = new System.Drawing.Point(904, 440);
			this.txtSaldoConciliado.Name = "txtSaldoConciliado";
			this.txtSaldoConciliado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoConciliado.PromptChar = ' ';
			this.txtSaldoConciliado.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoConciliado.TabIndex = 666;
			// 
			// txtSaldoNoConciliado
			// 
			this.txtSaldoNoConciliado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoNoConciliado.Appearance = appearance32;
			this.txtSaldoNoConciliado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoNoConciliado.Enabled = false;
			this.txtSaldoNoConciliado.FormatString = "#,##0.00";
			this.txtSaldoNoConciliado.Location = new System.Drawing.Point(904, 416);
			this.txtSaldoNoConciliado.Name = "txtSaldoNoConciliado";
			this.txtSaldoNoConciliado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoNoConciliado.PromptChar = ' ';
			this.txtSaldoNoConciliado.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoNoConciliado.TabIndex = 665;
			// 
			// label19
			// 
			this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label19.Location = new System.Drawing.Point(784, 392);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(81, 16);
			this.label19.TabIndex = 664;
			this.label19.Text = "Saldo en Libros";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldoEnLibros
			// 
			this.txtSaldoEnLibros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoEnLibros.Appearance = appearance33;
			this.txtSaldoEnLibros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoEnLibros.Enabled = false;
			this.txtSaldoEnLibros.FormatString = "#,##0.00";
			this.txtSaldoEnLibros.Location = new System.Drawing.Point(904, 392);
			this.txtSaldoEnLibros.Name = "txtSaldoEnLibros";
			this.txtSaldoEnLibros.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoEnLibros.PromptChar = ' ';
			this.txtSaldoEnLibros.Size = new System.Drawing.Size(112, 22);
			this.txtSaldoEnLibros.TabIndex = 663;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 576);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 40;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// txtIdConciliacionBancaria
			// 
			this.txtIdConciliacionBancaria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdConciliacionBancaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdConciliacionBancaria.Enabled = false;
			this.txtIdConciliacionBancaria.Location = new System.Drawing.Point(968, 576);
			this.txtIdConciliacionBancaria.Name = "txtIdConciliacionBancaria";
			this.txtIdConciliacionBancaria.PromptChar = ' ';
			this.txtIdConciliacionBancaria.Size = new System.Drawing.Size(16, 22);
			this.txtIdConciliacionBancaria.TabIndex = 677;
			this.txtIdConciliacionBancaria.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(1000, 576);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 676;
			this.txtEstado.Visible = false;
			// 
			// txtIdCuenta
			// 
			this.txtIdCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCuenta.Enabled = false;
			this.txtIdCuenta.Location = new System.Drawing.Point(984, 576);
			this.txtIdCuenta.Name = "txtIdCuenta";
			this.txtIdCuenta.PromptChar = ' ';
			this.txtIdCuenta.Size = new System.Drawing.Size(16, 22);
			this.txtIdCuenta.TabIndex = 675;
			this.txtIdCuenta.Visible = false;
			// 
			// txtBanco
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBanco.Appearance = appearance34;
			this.txtBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBanco.Enabled = false;
			this.txtBanco.Location = new System.Drawing.Point(64, 8);
			this.txtBanco.Name = "txtBanco";
			this.txtBanco.Size = new System.Drawing.Size(360, 22);
			this.txtBanco.TabIndex = 679;
			// 
			// txtCuenta
			// 
			this.txtCuenta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuenta.Appearance = appearance35;
			this.txtCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuenta.Enabled = false;
			this.txtCuenta.Location = new System.Drawing.Point(520, 8);
			this.txtCuenta.Name = "txtCuenta";
			this.txtCuenta.Size = new System.Drawing.Size(224, 22);
			this.txtCuenta.TabIndex = 680;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(8, 11);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 681;
			this.label17.Text = "Banco";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(456, 11);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(39, 16);
			this.label20.TabIndex = 682;
			this.label20.Text = "Cuenta";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(87, 576);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 683;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnBloquear
			// 
			this.btnBloquear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBloquear.CausesValidation = false;
			this.btnBloquear.Enabled = false;
			this.btnBloquear.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBloquear.Location = new System.Drawing.Point(609, 576);
			this.btnBloquear.Name = "btnBloquear";
			this.btnBloquear.Size = new System.Drawing.Size(112, 23);
			this.btnBloquear.TabIndex = 684;
			this.btnBloquear.Text = "Cerrar Conciliación";
			this.btnBloquear.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBloquear.Click += new System.EventHandler(this.btnBloquear_Click);
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(336, 88);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(75, 22);
			this.btnBuscar.TabIndex = 685;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.Location = new System.Drawing.Point(422, 576);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(97, 23);
			this.btnAsiento.TabIndex = 687;
			this.btnAsiento.Text = "&Ver Movimiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(339, 576);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 686;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(172, 576);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 688;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(728, 576);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 689;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(784, 528);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 17);
			this.lblUsuario.TabIndex = 690;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// mnuExportar
			// 
			this.mnuExportar.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuExportarEntradas,
																																								this.mnuExportarSalidas});
			// 
			// mnuExportarEntradas
			// 
			this.mnuExportarEntradas.Index = 0;
			this.mnuExportarEntradas.Text = "Entradas";
			this.mnuExportarEntradas.Click += new System.EventHandler(this.mnuExportarEntradas_Click);
			// 
			// mnuExportarSalidas
			// 
			this.mnuExportarSalidas.Index = 1;
			this.mnuExportarSalidas.Text = "Salidas";
			this.mnuExportarSalidas.Click += new System.EventHandler(this.mnuExportarSalidas_Click);
			// 
			// mnuMovimiento
			// 
			this.mnuMovimiento.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																									this.mnuMovimientoEntrada,
																																									this.mnuMovimientoSalida});
			// 
			// mnuMovimientoEntrada
			// 
			this.mnuMovimientoEntrada.Index = 0;
			this.mnuMovimientoEntrada.Text = "Entrada";
			this.mnuMovimientoEntrada.Click += new System.EventHandler(this.mnuMovimientoEntrada_Click);
			// 
			// mnuMovimientoSalida
			// 
			this.mnuMovimientoSalida.Index = 1;
			this.mnuMovimientoSalida.Text = "Salida";
			this.mnuMovimientoSalida.Click += new System.EventHandler(this.mnuMovimientoSalida_Click);
			// 
			// chkSelIngresos
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSelIngresos.Appearance = appearance36;
			this.chkSelIngresos.BackColor = System.Drawing.Color.Transparent;
			this.chkSelIngresos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkSelIngresos.Enabled = false;
			this.chkSelIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.chkSelIngresos.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.chkSelIngresos.Location = new System.Drawing.Point(730, 104);
			this.chkSelIngresos.Name = "chkSelIngresos";
			this.chkSelIngresos.Size = new System.Drawing.Size(24, 22);
			this.chkSelIngresos.TabIndex = 691;
			this.chkSelIngresos.CheckedChanged += new System.EventHandler(this.chkSelIngresos_CheckedChanged);
			// 
			// frmConciliacionBancaria
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1024, 606);
			this.Controls.Add(this.chkSelIngresos);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.btnBloquear);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtCuenta);
			this.Controls.Add(this.txtBanco);
			this.Controls.Add(this.txtIdConciliacionBancaria);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCuenta);
			this.Controls.Add(this.txtDiferencias);
			this.Controls.Add(this.txtSaldoBancario);
			this.Controls.Add(this.txtSaldoConciliado);
			this.Controls.Add(this.txtSaldoNoConciliado);
			this.Controls.Add(this.txtSaldoEnLibros);
			this.Controls.Add(this.txtNoConciliadosSalida);
			this.Controls.Add(this.txtNoConciliadosEntrada);
			this.Controls.Add(this.txtConciliadosSalidas);
			this.Controls.Add(this.txtConciliadosEntrada);
			this.Controls.Add(this.txtSistemaSalida);
			this.Controls.Add(this.txtSistemaEntrada);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.uGridSalida);
			this.Controls.Add(this.uGridEntrada);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.KeyPreview = true;
			this.Name = "frmConciliacionBancaria";
			this.Text = "Conciliación Bancaria";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConciliacionBancaria_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmConciliacionBancaria_Closing);
			this.Load += new System.EventHandler(this.frmConciliacionBancaria_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBancario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoConciliado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNoConciliado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoEnLibros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConciliacionBancaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void label1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void EntradasSalidas()
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
      			
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec ConciliacionBancariaMovimientos {0}, '{1}', '{2}', '{3}', {4}, 0", (int)this.txtIdCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtBuscar.Text.ToString(), (int)this.txtIdConciliacionBancaria.Value), this.uGridEntrada);
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec ConciliacionBancariaMovimientos {0}, '{1}', '{2}', '{3}', {4}, 1", (int)this.txtIdCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtBuscar.Text.ToString(), (int)this.txtIdConciliacionBancaria.Value), this.uGridSalida);      

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec SumaValoresConciliacionBancaria {0}, '{1}', '{2}'", (int)this.txtIdCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd")), true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.txtSistemaEntrada.Value = Convert.ToDecimal(dr.GetValue(0));
				this.txtSistemaSalida.Value = Convert.ToDecimal(dr.GetValue(1));
			}
			dr.Close();

			if ((int)this.txtEstado.Value != 2) Resumen();
		}

		private void ModificaGrillaEntrada(bool bModifica)
		{
			this.uGridEntrada.Enabled = true;
			if (bModifica)
			{
				this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;			
			}
			else
			{
				this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void ModificaGrillaSalida(bool bModifica)
		{
			this.uGridSalida.Enabled = true;
			if (bModifica)
			{
				this.uGridSalida.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;				
			}
			else
			{
				this.uGridSalida.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;				
			}
		}

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConciliacionBancaria_Load(object sender, System.EventArgs e)
		{						
			miAcceso = new Acceso(cdsSeteoF, "040807");			

			if (!Funcion.Permiso("58", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Conciliacion Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}
			else
			{
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

				ModificaGrillaEntrada(false);
				ModificaGrillaSalida(false);

				if (idConciliacionBancaria > 0) this.Consulta(idConciliacionBancaria);	
			}
		}

		private void txtBuscar_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConciliacionNueva Crear = new frmConciliacionNueva())
			{				
				if (DialogResult.OK == Crear.ShowDialog())
				{
					this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, "Exec NumeroConciliacion"); // Crear.txtSecuencial.Text.ToString();

					this.txtIdCuenta.Value = (int)Crear.cmbCuenta.Value;					
					this.txtBanco.Text = Crear.txtBanco.Text.ToString();
					this.txtCuenta.Text = Crear.txtCuenta.Text.ToString();
					this.dtDesde.Value = (DateTime)Crear.dtDesde.Value;
					this.dtHasta.Value = (DateTime)Crear.dtHasta.Value;					
					this.txtSaldoBancario.Value = Convert.ToDecimal(Crear.txtSaldo.Value);					

					DateTime dtDesde = (DateTime) this.dtDesde.Value;
					DateTime dtHasta = (DateTime) this.dtHasta.Value;

					string sSQL = string.Format("Exec GuardaConciliacionBancaria {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}",
						(int)this.txtIdConciliacionBancaria.Value, (int)this.txtIdCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtNumero.Text.ToString(), (int)this.txtEstado.Value,
						Convert.ToDecimal(this.txtSaldoEnLibros.Value), Convert.ToDecimal(this.txtSaldoConciliado.Value), Convert.ToDecimal(this.txtSaldoNoConciliado.Value), Convert.ToDecimal(this.txtSaldoBancario.Value), Convert.ToDecimal(this.txtDiferencias.Value));
					this.txtIdConciliacionBancaria.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

					this.txtBuscar.Enabled = true;
					this.btnBuscar.Enabled = true;
					
					if (miAcceso.BModificarSaldoCB) this.txtSaldoBancario.Enabled = true;					
				
					bNuevo = true;
					bEdicion = true;
					
					ModificaGrillaEntrada(true);
					ModificaGrillaSalida(true);

					EntradasSalidas();

					this.chkSelIngresos.Enabled = true;
					this.btnNuevo.Enabled = false;	
					this.btnConsultar.Enabled = false;	
					this.btnBloquear.Enabled = false;	
					this.btnImprimir.Enabled = false;
					this.btnCancelar.Enabled = true;
					this.btnAnular.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnGuardar.Enabled = true;
					this.btnExcel.Enabled = true;

					if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;

					this.txtBuscar.Focus();
				}
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void Resumen()
		{
			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			string sSQLResumen = string.Format("Exec ResumenConciliacionBancaria {0}, '{1}', '{2}', {3}", 
				(int)this.txtIdCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), (int)this.txtIdConciliacionBancaria.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLResumen, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtConciliadosEntrada.Value = Convert.ToDecimal(dr.GetValue(0));
				this.txtConciliadosSalidas.Value = Convert.ToDecimal(dr.GetValue(1));
				this.txtNoConciliadosEntrada.Value = Convert.ToDecimal(dr.GetValue(2));
				this.txtNoConciliadosSalida.Value = Convert.ToDecimal(dr.GetValue(3));
				this.txtSaldoEnLibros.Value = Convert.ToDecimal(dr.GetValue(4));
			}
			dr.Close();
			
			this.txtSaldoNoConciliado.Value = Convert.ToDecimal(this.txtNoConciliadosEntrada.Value) - Convert.ToDecimal(this.txtNoConciliadosSalida.Value);
			this.txtSaldoConciliado.Value = Convert.ToDecimal(this.txtSaldoEnLibros.Value) - Convert.ToDecimal(this.txtSaldoNoConciliado.Value);
			this.txtDiferencias.Value = Convert.ToDecimal(this.txtSaldoConciliado.Value) - Convert.ToDecimal(this.txtSaldoBancario.Value);
		}
		
		private void uGridEntrada_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "Confirmado")
//			{
//				this.uGridEntrada.UpdateData();
//
//				string sSQL = string.Format("Exec GuardaDetalleConciliacionBancaria {0}, {1}, {2}", 
//					(int)e.Cell.Row.Cells["idDetAsiento"].Value, (int)this.txtIdConciliacionBancaria.Value, (bool)e.Cell.Row.Cells["Confirmado"].Value);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
//
//				DateTime dtDesde = (DateTime)this.dtDesde.Value;
//
//				//Resumen();
//			}
		}

		private void uGridSalida_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void txtSaldoBancario_Validated(object sender, System.EventArgs e)
		{
			Resumen();
		}

		private void Consulta(int idConciliacionBancaria)
		{			
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnBloquear.Enabled = false;	
				this.btnImprimir.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnBuscar.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAsiento.Enabled = false;			
				this.btnEditar.Enabled = false;
				this.btnBuscar.Enabled = false;
				this.btnExcel.Enabled = false;

				#region Consulta Registro
				string sSQL = string.Format("Exec CBConsultaIndividual {0}", idConciliacionBancaria);
				SqlDataReader drBusca = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drBusca.Read();
				if (drBusca.HasRows)
				{					
					this.txtIdConciliacionBancaria.Value = drBusca.GetInt32(0);
					this.txtIdCuenta.Value = drBusca.GetInt32(1);
					this.txtBanco.Text = drBusca.GetString(2);
					this.txtCuenta.Text = drBusca.GetString(3);
					this.dtDesde.Value = drBusca.GetDateTime(4);
					this.dtHasta.Value = drBusca.GetDateTime(5);
					this.txtNumero.Value = drBusca.GetString(6);					
					this.txtSaldoBancario.Value = drBusca.GetDecimal(7);
					this.txtEstado.Value = drBusca.GetInt32(8);
					this.lblEstado.Text = drBusca.GetString(9);
					this.lblUsuario.Text = drBusca.GetString(10);
				}
				drBusca.Close();

				EntradasSalidas();
				#endregion Consulta Registro
				
				#region Controles
				if ((int)this.txtEstado.Value != 2)
				{
					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

					if ((int)this.txtEstado.Value == 0) 
					{
						if (miAcceso.BBloquearCB) this.btnBloquear.Enabled = true;
						this.btnBloquear.Text = "Cerrar Conciliación";
				
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Anular) this.btnAnular.Enabled = true;					
					}
					else if ((int)this.txtEstado.Value == 1) 
					{
						if (miAcceso.BDesbloquearCB) this.btnBloquear.Enabled = true;
						this.btnBloquear.Text = "Abrir Conciliación";				
					}						
				
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;

					this.btnExcel.Enabled = true;					
				}

				this.btnCancelar.Enabled = true;
				#endregion Controles
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmBuscaConciliaciones Busqueda = new frmBuscaConciliaciones())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridConciliaciones.ActiveRow.Cells["idConciliacionBancaria"].Value);
				}
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 12", idBloqueaTransacciones, (int)this.txtIdConciliacionBancaria.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bNuevo = false;
			bEdicion = false;

			ModificaGrillaEntrada(false);
			
			ModificaGrillaSalida(false);

			this.chkSelIngresos.Checked = false;
			this.txtIdConciliacionBancaria.Value = 0;
			this.txtIdCuenta.Value = 0;
			this.txtEstado.Value = 0;
			this.txtBanco.Text = "";
			this.txtCuenta.Text = "";
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;			
			this.txtNumero.Text = "";
			this.txtBuscar.Text = "";
			this.txtSistemaEntrada.Value = 0.00m;
			this.txtSistemaSalida.Value = 0.00m;
			this.txtConciliadosEntrada.Value = 0.00m;
			this.txtConciliadosSalidas.Value = 0.00m;
			this.txtNoConciliadosEntrada.Value = 0.00m;
			this.txtNoConciliadosSalida.Value = 0.00m;
			this.txtSaldoEnLibros.Value = 0.00m;
			this.txtSaldoNoConciliado.Value = 0.00m;
			this.txtSaldoConciliado.Value = 0.00m;
			this.txtSaldoBancario.Value = 0.00m;
			this.txtDiferencias.Value = 0.00m;
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";

			this.btnBloquear.Text = "Bloquear";

			this.txtBuscar.Enabled = false;
			this.txtSaldoBancario.Enabled = false;

			this.chkSelIngresos.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnBloquear.Enabled = false;	
			this.btnImprimir.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnBuscar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAsiento.Enabled = false;			
			this.btnEditar.Enabled = false;
			this.btnBuscar.Enabled = false;
			this.btnExcel.Enabled = false;
			idBloqueaTransacciones = 0;

			this.uGridEntrada.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConciliacionBancariaMovimientos 0, '', '', '', 0, 0");
			this.uGridSalida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConciliacionBancariaMovimientos 0, '', '', '', 0, 1");      
		}

		private void btnBloquear_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 0)
			{
				if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Cerrar esta Conciliación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					string sSQL = string.Format("Update ConciliacionBancaria Set Estado = 1 Where idConciliacionBancaria = {0}", (int)this.txtIdConciliacionBancaria.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
					this.lblEstado.Text = "CERRADO";

					this.btnBloquear.Text = "Abrir Conciliación";

					this.txtEstado.Value = 1;

					if (!miAcceso.BDesbloquearCB) this.btnBloquear.Enabled = false;	
		
					this.btnEditar.Enabled = false;
					this.btnAnular.Enabled = false;
				}
			}
			else if ((int)this.txtEstado.Value == 1)
			{
				if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Abrir esta Conciliación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					string sSQL = string.Format("Update ConciliacionBancaria Set Estado = 0 Where idConciliacionBancaria = {0}", (int)this.txtIdConciliacionBancaria.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
					this.lblEstado.Text = "ABIERTO";

					this.btnBloquear.Text = "Cerrar Conciliación";

					this.txtEstado.Value = 0;

					if (!miAcceso.BBloquearCB) this.btnBloquear.Enabled = false;

					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Anular) this.btnAnular.Enabled = true;
				}
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaBloqueo((int)this.txtIdConciliacionBancaria.Value, 12, cdsSeteoF)) return;

			this.Consulta((int)this.txtIdConciliacionBancaria.Value);

			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("No puede anular esta conciliacion esta CERRADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("Esta conciliacion ya esta ANULADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("Esta Seguro de Anular esta Conciliación", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
			{
				string sSQLAnula = string.Format("Exec AnulaConciliacionBancaria {0}", (int)this.txtIdConciliacionBancaria.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAnula, true);

				this.lblEstado.Text = "ANULADO";
			
				this.txtEstado.Value = 2;

				this.btnBloquear.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnExcel.Enabled = false;
			}
		}

		private void uGridEntrada_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtBuscar_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSaldoBancario_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void uGridEntrada_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idAsiento"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Confirmado")
			{
				this.uGridEntrada.UpdateData();

				string sSQL = string.Format("Exec GuardaDetalleConciliacionBancaria {0}, {1}, {2}", 
					(int)e.Cell.Row.Cells["idDetAsiento"].Value, (int)this.txtIdConciliacionBancaria.Value, (bool)e.Cell.Row.Cells["Confirmado"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				DateTime dtDesde = (DateTime)this.dtDesde.Value;

				//Resumen();
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{			
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idConciliacion";
			ValIdCompra.Value = int.Parse(this.txtIdConciliacionBancaria.Value.ToString());
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);
				
			Reporte miRep = new Reporte("ResumenConciliacionBancaria.rpt", "");
			miRep.MdiParent = this.MdiParent;			
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
				
			ParameterFields pfsDetalle = new ParameterFields ();

			ParameterField pFIdConciliacion = new ParameterField ();
			ParameterDiscreteValue pdvIdConciliacion = new ParameterDiscreteValue ();
			pFIdConciliacion.ParameterFieldName = "@idConciliacion";
			pdvIdConciliacion.Value = int.Parse(this.txtIdConciliacionBancaria.Value.ToString());
			pFIdConciliacion.CurrentValues.Add (pdvIdConciliacion);
			pfsDetalle.Add (pFIdConciliacion);
			
			ParameterField pfIdCuenta = new ParameterField ();
			ParameterDiscreteValue pdvIdCuenta = new ParameterDiscreteValue ();
			pfIdCuenta.ParameterFieldName = "@idCuenta";
			pdvIdCuenta.Value = int.Parse(this.txtIdCuenta.Value.ToString());
			pfIdCuenta.CurrentValues.Add (pdvIdCuenta);
			pfsDetalle.Add (pfIdCuenta);

			ParameterField pfDesde = new ParameterField ();
			ParameterDiscreteValue pdvDesde = new ParameterDiscreteValue ();
			pfDesde.ParameterFieldName = "@Desde";
			pdvDesde.Value = (DateTime)this.dtDesde.Value;
			pfDesde.CurrentValues.Add (pdvDesde);
			pfsDetalle.Add (pfDesde);
			
			ParameterField pfHasta = new ParameterField ();
			ParameterDiscreteValue pdvHasta = new ParameterDiscreteValue ();
			pfHasta.ParameterFieldName = "@Hasta";
			pdvHasta.Value = (DateTime)this.dtHasta.Value;
			pfHasta.CurrentValues.Add (pdvHasta);
			pfsDetalle.Add (pfHasta);

			Reporte Rep = new Reporte("ListaDetalleConciliacionBancaria.rpt", "");
			Rep.MdiParent = this.MdiParent;			
			Rep.crVista.ParameterFieldInfo = pfsDetalle;
			Rep.Show();
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			EntradasSalidas();
		}

		private void txtBuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btnBuscar_Click(sender, e);
		}

		private void uGridSalida_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void uGridSalida_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idAsiento"].Value == System.DBNull.Value) return;

			if (e.Cell.Column.ToString() == "Confirmado")
			{
				this.uGridSalida.UpdateData();

				string sSQL = string.Format("Exec GuardaDetalleConciliacionBancaria {0}, {1}, {2}", 
					(int)e.Cell.Row.Cells["idDetAsiento"].Value, (int)this.txtIdConciliacionBancaria.Value, (bool)e.Cell.Row.Cells["Confirmado"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);				
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			if (Funcion.ValidaBloqueo((int)this.txtIdConciliacionBancaria.Value, 12, cdsSeteoF)) return;

			this.txtBuscar.Enabled = true;
			this.btnBuscar.Enabled = true;

			if (miAcceso.BModificarSaldoCB) this.txtSaldoBancario.Enabled = true;

			ModificaGrillaEntrada(true);
			ModificaGrillaSalida(true);

			bNuevo = false;
			bEdicion = true;
					
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnBloquear.Enabled = false;
			this.btnAnular.Enabled = false;			
			if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;

			this.chkSelIngresos.Enabled = true;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 12", (int)this.txtIdConciliacionBancaria.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) this.dtDesde.Value;
			DateTime dtHasta = (DateTime) this.dtHasta.Value;

			string sSQL = string.Format("Exec GuardaConciliacionBancaria {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}",
				(int)this.txtIdConciliacionBancaria.Value, (int)this.txtIdCuenta.Value, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtNumero.Text.ToString(), (int)this.txtEstado.Value,
				Convert.ToDecimal(this.txtSaldoEnLibros.Value), Convert.ToDecimal(this.txtSaldoConciliado.Value), Convert.ToDecimal(this.txtSaldoNoConciliado.Value), Convert.ToDecimal(this.txtSaldoBancario.Value), Convert.ToDecimal(this.txtDiferencias.Value));
			this.txtIdConciliacionBancaria.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			
			if (Convert.ToDecimal(this.txtDiferencias.Value) == 0.00m)
			{
				if (DialogResult.Yes == MessageBox.Show("La Conciliacion esta cuadrada,\n\n¿Desea Cerrar la Conciliacion?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
				{
					this.btnBloquear_Click(sender, e);
					this.btnImprimir_Click(sender, e);
				}
				else 
				{
					if (miAcceso.BBloquearCB) this.btnBloquear.Enabled = true;
				}
			}
			else if (miAcceso.BBloquearCB) this.btnBloquear.Enabled = true;

			if (idBloqueaTransacciones > 0)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 12", idBloqueaTransacciones, (int)this.txtIdConciliacionBancaria.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			ModificaGrillaEntrada(false);

			ModificaGrillaSalida(false);
						
			this.txtBuscar.Enabled = false;
			this.btnBuscar.Enabled = false;
			
			this.txtSaldoBancario.Enabled = false;
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;						
			if (miAcceso.Anular) this.btnAnular.Enabled = true;
			if (miAcceso.CrearAsiento) this.btnAsiento.Enabled = true;
			
			this.btnGuardar.Enabled = false;
			this.chkSelIngresos.Enabled = false;
			
			bNuevo = false;
			bEdicion = false;			
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			this.mnuMovimiento.Show(this.btnAsiento, new Point(1, 25));
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			if (this.uGridEntrada.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridEntrada, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			if (this.uGridSalida.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
				if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
				{				 
					ultraGridExcelExporter1.Export(this.uGridSalida, saveFileDialog1.FileName);
					System.Diagnostics.Process.Start(saveFileDialog1.FileName);
				}
			}
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			this.mnuExportar.Show(this.btnExcel, new Point(1, 25));
		}

		private void mnuExportarEntradas_Click(object sender, System.EventArgs e)
		{
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridEntrada, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void mnuExportarSalidas_Click(object sender, System.EventArgs e)
		{
			this.saveFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";
			
			if(saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFileDialog1.FileName.Length > 0) 
			{				 
				ultraGridExcelExporter1.Export(this.uGridSalida, saveFileDialog1.FileName);
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
			}
		}

		private void mnuMovimientoEntrada_Click(object sender, System.EventArgs e)
		{
			if (this.uGridEntrada.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Movimiento para ver el Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (this.uGridEntrada.ActiveRow.Cells["idAsiento"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un Movimiento para ver el Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{				
				frmAsientos Asiento = new frmAsientos((int)this.uGridEntrada.ActiveRow.Cells["idAsiento"].Value);
				Asiento.ShowDialog();
			}
		}

		private void mnuMovimientoSalida_Click(object sender, System.EventArgs e)
		{
			if (this.uGridSalida.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un Movimiento para ver el Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else if (this.uGridSalida.ActiveRow.Cells["idAsiento"].Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un Movimiento para ver el Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
			}
			else
			{
				frmAsientos Asiento = new frmAsientos((int)this.uGridSalida.ActiveRow.Cells["idAsiento"].Value);
				Asiento.ShowDialog();
			}
		}

		private void frmConciliacionBancaria_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Conciliación Bancaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmConciliacionBancaria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdConciliacionBancaria.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
					if (e.KeyCode == Keys.A) if (miAcceso.Anular) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);
																	 
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0416");

						this.Consulta((int)this.txtIdConciliacionBancaria.Value);					
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
				if (e.KeyCode == Keys.F5) EntradasSalidas();
			}
		}

		private void chkSelIngresos_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridEntrada.Rows.All)
					dr.Cells["Confirmado"].Value = this.chkSelIngresos.Checked;
		}
	}
}

