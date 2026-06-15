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
	/// Descripción breve de frmConciliacionBancariaReporte.
	/// </summary>
	public class frmConciliacionBancariaReporte : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSelIngresos;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu mnuExportar;
		private System.Windows.Forms.MenuItem mnuExportarEntradas;
		private System.Windows.Forms.MenuItem mnuExportarSalidas;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBanco;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdConciliacionBancaria;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCuenta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiferencias;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoBancario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoConciliado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoNoConciliado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldoEnLibros;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNoConciliadosSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNoConciliadosEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConciliadosSalidas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConciliadosEntrada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSistemaSalida;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSistemaEntrada;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSalida;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEntrada;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.ContextMenu mnuMovimiento;
		private System.Windows.Forms.MenuItem mnuMovimientoEntrada;
		private System.Windows.Forms.MenuItem mnuMovimientoSalida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnConsultar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmConciliacionBancariaReporte()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmConciliacionBancariaReporte));
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
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConciliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cheque");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaConciliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Confirmado");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.chkSelIngresos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.mnuExportar = new System.Windows.Forms.ContextMenu();
			this.mnuExportarEntradas = new System.Windows.Forms.MenuItem();
			this.mnuExportarSalidas = new System.Windows.Forms.MenuItem();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.label20 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.txtBanco = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdConciliacionBancaria = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCuenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDiferencias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoBancario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoConciliado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoNoConciliado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSaldoEnLibros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNoConciliadosSalida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNoConciliadosEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConciliadosSalidas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConciliadosEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSistemaSalida = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSistemaEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.uGridSalida = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridEntrada = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.mnuMovimiento = new System.Windows.Forms.ContextMenu();
			this.mnuMovimientoEntrada = new System.Windows.Forms.MenuItem();
			this.mnuMovimientoSalida = new System.Windows.Forms.MenuItem();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cmbCuenta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnConsultar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConciliacionBancaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBancario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoConciliado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNoConciliado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoEnLibros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// chkSelIngresos
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSelIngresos.Appearance = appearance1;
			this.chkSelIngresos.BackColor = System.Drawing.Color.Transparent;
			this.chkSelIngresos.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.chkSelIngresos.Enabled = false;
			this.chkSelIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.chkSelIngresos.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.chkSelIngresos.Location = new System.Drawing.Point(730, 104);
			this.chkSelIngresos.Name = "chkSelIngresos";
			this.chkSelIngresos.Size = new System.Drawing.Size(24, 22);
			this.chkSelIngresos.TabIndex = 751;
			this.chkSelIngresos.Visible = false;
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
			this.lblUsuario.TabIndex = 750;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(808, 576);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 749;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// ultraDataSource3
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
																																 ultraDataColumn10});
			// 
			// ultraDataSource1
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
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(512, 88);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(75, 22);
			this.btnBuscar.TabIndex = 745;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Visible = false;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// label20
			// 
			this.label20.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Location = new System.Drawing.Point(456, 10);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(41, 16);
			this.label20.TabIndex = 742;
			this.label20.Text = "Cuenta";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(8, 11);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(36, 16);
			this.label17.TabIndex = 741;
			this.label17.Text = "Banco";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Location = new System.Drawing.Point(784, 488);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(61, 16);
			this.label13.TabIndex = 735;
			this.label13.Text = "Diferencias";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Visible = false;
			// 
			// label14
			// 
			this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Location = new System.Drawing.Point(784, 464);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(81, 16);
			this.label14.TabIndex = 734;
			this.label14.Text = "Saldo Bancario";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label14.Visible = false;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label15.Location = new System.Drawing.Point(784, 440);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(89, 16);
			this.label15.TabIndex = 733;
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
			this.label16.Size = new System.Drawing.Size(107, 16);
			this.label16.TabIndex = 732;
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
			this.label18.TabIndex = 731;
			this.label18.Text = "CONCILIACIÓN";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label19.Location = new System.Drawing.Point(784, 392);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(83, 16);
			this.label19.TabIndex = 726;
			this.label19.Text = "Saldo en Libros";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label8.Location = new System.Drawing.Point(784, 312);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(41, 16);
			this.label8.TabIndex = 724;
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
			this.label9.Size = new System.Drawing.Size(50, 16);
			this.label9.TabIndex = 723;
			this.label9.Text = "Entradas";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label5.Location = new System.Drawing.Point(784, 224);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(41, 16);
			this.label5.TabIndex = 722;
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
			this.label7.Size = new System.Drawing.Size(50, 16);
			this.label7.TabIndex = 721;
			this.label7.Text = "Entradas";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.label12.TabIndex = 720;
			this.label12.Text = "NO CONCILIADOS";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.label11.TabIndex = 719;
			this.label11.Text = "CONCILIADOS";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			this.label10.TabIndex = 718;
			this.label10.Text = "SISTEMA";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Location = new System.Drawing.Point(784, 136);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(41, 16);
			this.label4.TabIndex = 713;
			this.label4.Text = "Salidas";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Location = new System.Drawing.Point(784, 112);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(50, 16);
			this.label2.TabIndex = 711;
			this.label2.Text = "Entradas";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(208, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(14, 16);
			this.label1.TabIndex = 704;
			this.label1.Text = "Al";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 91);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 702;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 43);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(21, 16);
			this.label6.TabIndex = 700;
			this.label6.Text = "Del";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBanco
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBanco.Appearance = appearance2;
			this.txtBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBanco.Enabled = false;
			this.txtBanco.Location = new System.Drawing.Point(64, 8);
			this.txtBanco.Name = "txtBanco";
			this.txtBanco.Size = new System.Drawing.Size(360, 21);
			this.txtBanco.TabIndex = 739;
			// 
			// txtIdConciliacionBancaria
			// 
			this.txtIdConciliacionBancaria.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdConciliacionBancaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdConciliacionBancaria.Enabled = false;
			this.txtIdConciliacionBancaria.Location = new System.Drawing.Point(968, 576);
			this.txtIdConciliacionBancaria.Name = "txtIdConciliacionBancaria";
			this.txtIdConciliacionBancaria.PromptChar = ' ';
			this.txtIdConciliacionBancaria.Size = new System.Drawing.Size(16, 21);
			this.txtIdConciliacionBancaria.TabIndex = 738;
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
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 737;
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
			this.txtIdCuenta.Size = new System.Drawing.Size(16, 21);
			this.txtIdCuenta.TabIndex = 736;
			this.txtIdCuenta.Visible = false;
			// 
			// txtDiferencias
			// 
			this.txtDiferencias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiferencias.Appearance = appearance3;
			this.txtDiferencias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiferencias.Enabled = false;
			this.txtDiferencias.FormatString = "#,##0.00";
			this.txtDiferencias.Location = new System.Drawing.Point(904, 488);
			this.txtDiferencias.Name = "txtDiferencias";
			this.txtDiferencias.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiferencias.PromptChar = ' ';
			this.txtDiferencias.Size = new System.Drawing.Size(112, 21);
			this.txtDiferencias.TabIndex = 730;
			this.txtDiferencias.Visible = false;
			// 
			// txtSaldoBancario
			// 
			this.txtSaldoBancario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoBancario.Appearance = appearance4;
			this.txtSaldoBancario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoBancario.Enabled = false;
			this.txtSaldoBancario.FormatString = "#,##0.00";
			this.txtSaldoBancario.Location = new System.Drawing.Point(904, 464);
			this.txtSaldoBancario.Name = "txtSaldoBancario";
			this.txtSaldoBancario.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoBancario.PromptChar = ' ';
			this.txtSaldoBancario.Size = new System.Drawing.Size(112, 21);
			this.txtSaldoBancario.TabIndex = 729;
			this.txtSaldoBancario.Visible = false;
			// 
			// txtSaldoConciliado
			// 
			this.txtSaldoConciliado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoConciliado.Appearance = appearance5;
			this.txtSaldoConciliado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoConciliado.Enabled = false;
			this.txtSaldoConciliado.FormatString = "#,##0.00";
			this.txtSaldoConciliado.Location = new System.Drawing.Point(904, 440);
			this.txtSaldoConciliado.Name = "txtSaldoConciliado";
			this.txtSaldoConciliado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoConciliado.PromptChar = ' ';
			this.txtSaldoConciliado.Size = new System.Drawing.Size(112, 21);
			this.txtSaldoConciliado.TabIndex = 728;
			// 
			// txtSaldoNoConciliado
			// 
			this.txtSaldoNoConciliado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoNoConciliado.Appearance = appearance6;
			this.txtSaldoNoConciliado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoNoConciliado.Enabled = false;
			this.txtSaldoNoConciliado.FormatString = "#,##0.00";
			this.txtSaldoNoConciliado.Location = new System.Drawing.Point(904, 416);
			this.txtSaldoNoConciliado.Name = "txtSaldoNoConciliado";
			this.txtSaldoNoConciliado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoNoConciliado.PromptChar = ' ';
			this.txtSaldoNoConciliado.Size = new System.Drawing.Size(112, 21);
			this.txtSaldoNoConciliado.TabIndex = 727;
			// 
			// txtSaldoEnLibros
			// 
			this.txtSaldoEnLibros.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldoEnLibros.Appearance = appearance7;
			this.txtSaldoEnLibros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldoEnLibros.Enabled = false;
			this.txtSaldoEnLibros.FormatString = "#,##0.00";
			this.txtSaldoEnLibros.Location = new System.Drawing.Point(904, 392);
			this.txtSaldoEnLibros.Name = "txtSaldoEnLibros";
			this.txtSaldoEnLibros.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldoEnLibros.PromptChar = ' ';
			this.txtSaldoEnLibros.Size = new System.Drawing.Size(112, 21);
			this.txtSaldoEnLibros.TabIndex = 725;
			// 
			// txtNoConciliadosSalida
			// 
			this.txtNoConciliadosSalida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNoConciliadosSalida.Appearance = appearance8;
			this.txtNoConciliadosSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNoConciliadosSalida.Enabled = false;
			this.txtNoConciliadosSalida.FormatString = "#,##0.00";
			this.txtNoConciliadosSalida.Location = new System.Drawing.Point(904, 312);
			this.txtNoConciliadosSalida.MaxValue = ((long)(22147483647));
			this.txtNoConciliadosSalida.Name = "txtNoConciliadosSalida";
			this.txtNoConciliadosSalida.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNoConciliadosSalida.PromptChar = ' ';
			this.txtNoConciliadosSalida.Size = new System.Drawing.Size(112, 21);
			this.txtNoConciliadosSalida.TabIndex = 717;
			// 
			// txtNoConciliadosEntrada
			// 
			this.txtNoConciliadosEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNoConciliadosEntrada.Appearance = appearance9;
			this.txtNoConciliadosEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNoConciliadosEntrada.Enabled = false;
			this.txtNoConciliadosEntrada.FormatString = "#,##0.00";
			this.txtNoConciliadosEntrada.Location = new System.Drawing.Point(904, 288);
			this.txtNoConciliadosEntrada.MaxValue = ((long)(22147483647));
			this.txtNoConciliadosEntrada.Name = "txtNoConciliadosEntrada";
			this.txtNoConciliadosEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNoConciliadosEntrada.PromptChar = ' ';
			this.txtNoConciliadosEntrada.Size = new System.Drawing.Size(112, 21);
			this.txtNoConciliadosEntrada.TabIndex = 716;
			// 
			// txtConciliadosSalidas
			// 
			this.txtConciliadosSalidas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConciliadosSalidas.Appearance = appearance10;
			this.txtConciliadosSalidas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConciliadosSalidas.Enabled = false;
			this.txtConciliadosSalidas.FormatString = "#,##0.00";
			this.txtConciliadosSalidas.Location = new System.Drawing.Point(904, 224);
			this.txtConciliadosSalidas.Name = "txtConciliadosSalidas";
			this.txtConciliadosSalidas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtConciliadosSalidas.PromptChar = ' ';
			this.txtConciliadosSalidas.Size = new System.Drawing.Size(112, 21);
			this.txtConciliadosSalidas.TabIndex = 715;
			// 
			// txtConciliadosEntrada
			// 
			this.txtConciliadosEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtConciliadosEntrada.Appearance = appearance11;
			this.txtConciliadosEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConciliadosEntrada.Enabled = false;
			this.txtConciliadosEntrada.FormatString = "#,##0.00";
			this.txtConciliadosEntrada.Location = new System.Drawing.Point(904, 200);
			this.txtConciliadosEntrada.Name = "txtConciliadosEntrada";
			this.txtConciliadosEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtConciliadosEntrada.PromptChar = ' ';
			this.txtConciliadosEntrada.Size = new System.Drawing.Size(112, 21);
			this.txtConciliadosEntrada.TabIndex = 714;
			// 
			// txtSistemaSalida
			// 
			this.txtSistemaSalida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSistemaSalida.Appearance = appearance12;
			this.txtSistemaSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSistemaSalida.Enabled = false;
			this.txtSistemaSalida.FormatString = "#,##0.00";
			this.txtSistemaSalida.Location = new System.Drawing.Point(904, 136);
			this.txtSistemaSalida.Name = "txtSistemaSalida";
			this.txtSistemaSalida.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSistemaSalida.PromptChar = ' ';
			this.txtSistemaSalida.Size = new System.Drawing.Size(112, 21);
			this.txtSistemaSalida.TabIndex = 712;
			// 
			// txtSistemaEntrada
			// 
			this.txtSistemaEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSistemaEntrada.Appearance = appearance13;
			this.txtSistemaEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSistemaEntrada.Enabled = false;
			this.txtSistemaEntrada.FormatString = "#,##0.00";
			this.txtSistemaEntrada.Location = new System.Drawing.Point(904, 112);
			this.txtSistemaEntrada.Name = "txtSistemaEntrada";
			this.txtSistemaEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSistemaEntrada.PromptChar = ' ';
			this.txtSistemaEntrada.Size = new System.Drawing.Size(112, 21);
			this.txtSistemaEntrada.TabIndex = 710;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(840, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 709;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
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
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance15;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 90;
			ultraGridColumn9.Format = "dd/MM/yyyy HH:MM";
			ultraGridColumn9.Header.Caption = "Conciliado";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 92;
			ultraGridColumn10.Header.Caption = "...";
			ultraGridColumn10.Header.VisiblePosition = 9;
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
			ultraGridBand1.Header.Caption = "Movimientos de Salida";
			ultraGridBand1.HeaderVisible = true;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance16;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridSalida.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
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
			this.uGridSalida.TabIndex = 708;
			this.uGridSalida.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			// 
			// uGridEntrada
			// 
			this.uGridEntrada.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEntrada.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntrada.DataSource = this.ultraDataSource1;
			appearance23.BackColor = System.Drawing.Color.White;
			this.uGridEntrada.DisplayLayout.Appearance = appearance23;
			this.uGridEntrada.DisplayLayout.AutoFitColumns = true;
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
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance24;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 7;
			ultraGridColumn18.Width = 90;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Format = "dd/MM/yyyy HH:MM";
			ultraGridColumn19.Header.Caption = "Conciliado";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Width = 92;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 9;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
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
			ultraGridBand2.Header.Caption = "Movimientos de Entrada";
			ultraGridBand2.HeaderVisible = true;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance25;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings2});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridEntrada.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntrada.DisplayLayout.Override.ActiveRowAppearance = appearance26;
			this.uGridEntrada.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridEntrada.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridEntrada.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance27.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntrada.DisplayLayout.Override.CardAreaAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8.25F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntrada.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.uGridEntrada.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance29.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance29.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowAlternateAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntrada.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.uGridEntrada.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntrada.Location = new System.Drawing.Point(8, 128);
			this.uGridEntrada.Name = "uGridEntrada";
			this.uGridEntrada.Size = new System.Drawing.Size(768, 208);
			this.uGridEntrada.TabIndex = 707;
			this.uGridEntrada.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 560);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1024, 8);
			this.groupBox2.TabIndex = 706;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1024, 8);
			this.groupBox1.TabIndex = 705;
			this.groupBox1.TabStop = false;
			// 
			// dtHasta
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance32;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(256, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 703;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance33;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 40);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 699;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
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
			// 
			// mnuMovimientoSalida
			// 
			this.mnuMovimientoSalida.Index = 1;
			this.mnuMovimientoSalida.Text = "Salida";
			// 
			// txtBuscar
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance34;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(64, 88);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(248, 21);
			this.txtBuscar.TabIndex = 698;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(889, 576);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 696;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// cmbCuenta
			// 
			this.cmbCuenta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCuenta.DataSource = this.ultraDataSource4;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.Caption = "Cuenta";
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 360;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCuenta.DisplayMember = "Descripcion";
			this.cmbCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCuenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCuenta.Location = new System.Drawing.Point(520, 8);
			this.cmbCuenta.MaxDropDownItems = 30;
			this.cmbCuenta.Name = "cmbCuenta";
			this.cmbCuenta.Size = new System.Drawing.Size(288, 21);
			this.cmbCuenta.TabIndex = 752;
			this.cmbCuenta.ValueMember = "idCuenta";
			this.cmbCuenta.ValueChanged += new System.EventHandler(this.cmbCuenta_ValueChanged);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(320, 88);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 753;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click_1);
			// 
			// frmConciliacionBancariaReporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 606);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.cmbCuenta);
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
			this.Controls.Add(this.label6);
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
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.chkSelIngresos);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnBuscar);
			this.KeyPreview = true;
			this.Name = "frmConciliacionBancariaReporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Reporte Conciliacion Bancaria ";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmConciliacionBancariaReporte_KeyDown);
			this.Load += new System.EventHandler(this.frmConciliacionBancariaReporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdConciliacionBancaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoBancario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoConciliado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoNoConciliado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldoEnLibros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNoConciliadosEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConciliadosEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSistemaEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		
		}

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmConciliacionBancariaReporte_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "47004");
			
			if (!Funcion.Permiso("2087", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar al reporte de Conciliación Bancaria.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Visible = true;

			DateTime dtFecha = DateTime.Parse("01/07/2018");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

			this.dtHasta.Value = DateTime.Today;
	

			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);

			this.cmbCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select c.idCuenta, c.Descripcion From Cuenta c Inner Join CuentaCorriente cc On cc.idCuenta = c.idCuenta Order By c.Descripcion");
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			
		}
		private void EntradasSalidas()
		{
			int IdCuenta = 0;
			if (this.cmbCuenta.ActiveRow != null) IdCuenta = (int)this.cmbCuenta.Value;


			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;
      			
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec [ConciliacionBancariaMovimientosReporte] {0}, '{1}', '{2}', '{3}', 0", IdCuenta, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtBuscar.Text.ToString()), this.uGridEntrada);
			FuncionesProcedimientos.dsGeneralModulo(string.Format("Exec [ConciliacionBancariaMovimientosReporte] {0}, '{1}', '{2}', '{3}', 1", IdCuenta, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), this.txtBuscar.Text.ToString()), this.uGridSalida);      

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Exec SumaValoresConciliacionBancaria {0}, '{1}', '{2}'", IdCuenta, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd")), true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.txtSistemaEntrada.Value = Convert.ToDecimal(dr.GetValue(0));
				this.txtSistemaSalida.Value = Convert.ToDecimal(dr.GetValue(1));
			}
			dr.Close();

			//if ((int)this.txtEstado.Value != 2) 
			Resumen();
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

		private void btnConsultar_Click_1(object sender, System.EventArgs e)
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

			EntradasSalidas();
		}
		private void Resumen()
		{
			int IdCuenta = 0;
			if (this.cmbCuenta.ActiveRow != null) IdCuenta = (int)this.cmbCuenta.Value;

			DateTime dtDesde = (DateTime)this.dtDesde.Value;
			DateTime dtHasta = (DateTime)this.dtHasta.Value;

			string sSQLResumen = string.Format("Exec ResumenConciliacionBancariaReporte {0}, '{1}', '{2}'", 
				IdCuenta, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
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

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCuenta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCuenta.ActiveRow != null)
			{
				this.txtBanco.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select b.Nombre from Cuenta c  Inner Join dbo.CuentaCorriente cc On cc.idCuenta = c.idCuenta Inner Join dbo.Banco b On b.idBanco = cc.idBanco where c.idCuenta  = {0}", (int) this.cmbCuenta.Value));
			}
			else
			{
				this.txtBanco.Text ="";
			}
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
						this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
						this.mnuExportar.Show(this.btnExcel, new Point(1, 25));
		}

		private void mnuExportarEntradas_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.uGridEntrada);
		}

		private void mnuExportarSalidas_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridSalida);
		}

		private void frmConciliacionBancariaReporte_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if (e.KeyCode == Keys.Escape) this.Close();
		}

	}
}
