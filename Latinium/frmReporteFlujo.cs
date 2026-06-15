using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;



namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteFlujo.
	/// </summary>
	public class frmReporteFlujo : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtestado;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		bool bNuevo = false;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidReporteFlujo;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Button button1;
		bool bEdicion = false;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnActualizarSaldos;
		DateTime dtFechaMin;

		public frmReporteFlujo()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMaestroFlujo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoLibro");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoBancos");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCheques");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depositos");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasCredito");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NotasDebito");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoTotales");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comprobacion");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesSaldoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesDeposito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesNotasCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesNotasDebito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSaldoBancos", 0);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteFlujo));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnDeposito", 1);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNotasCredito", 2);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnNotasDebito", 3);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoCheques", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoCheques", 8, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Depositos", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Depositos", 9, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NotasCredito", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NotasCredito", 10, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NotasDebito", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NotasDebito", 11, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoTotales", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoTotales", 12, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comprobacion", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comprobacion", 13, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoBancos", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoBancos", 7, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoLibro", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoLibro", 6, true);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMaestroFlujo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoLibro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoBancos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCheques");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Depositos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NotasDebito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoTotales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comprobacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesSaldoCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesDeposito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesNotasCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesNotasDebito");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btnSalir = new System.Windows.Forms.Button();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtestado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidReporteFlujo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.lblContador = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnActualizarSaldos = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidReporteFlujo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(456, 456);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 1019;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance1;
			this.uGridInformacion.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 150;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 200;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellButtonAppearance = appearance3;
			ultraGridColumn7.Format = "#,##0.00";
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.FormulaErrorAppearance = appearance4;
			ultraGridColumn7.Header.Caption = "Saldo Libros";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 100;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellButtonAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.FormulaErrorAppearance = appearance7;
			ultraGridColumn8.Header.Caption = "Saldo Bancos";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 100;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellButtonAppearance = appearance9;
			ultraGridColumn9.Format = "#,##0.00";
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.FormulaErrorAppearance = appearance10;
			ultraGridColumn9.Header.Caption = "Cheques";
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 100;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance11;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellButtonAppearance = appearance12;
			ultraGridColumn10.Format = "#,##0.00";
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.FormulaErrorAppearance = appearance13;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 100;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance14;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellButtonAppearance = appearance15;
			ultraGridColumn11.Format = "#,##0.00";
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.FormulaErrorAppearance = appearance16;
			ultraGridColumn11.Header.Caption = "Notas De Crédito";
			ultraGridColumn11.Header.VisiblePosition = 12;
			ultraGridColumn11.Width = 100;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance17;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellButtonAppearance = appearance18;
			ultraGridColumn12.Format = "#,##0.00";
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.FormulaErrorAppearance = appearance19;
			ultraGridColumn12.Header.Caption = "Notas De Débito";
			ultraGridColumn12.Header.VisiblePosition = 14;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(192)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance20;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellButtonAppearance = appearance21;
			ultraGridColumn13.Format = "#,##0.00";
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.FormulaErrorAppearance = appearance22;
			ultraGridColumn13.Header.Caption = "Saldo Total";
			ultraGridColumn13.Header.VisiblePosition = 16;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance23;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellButtonAppearance = appearance24;
			ultraGridColumn14.Format = "#,##0.00";
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.FormulaErrorAppearance = appearance25;
			ultraGridColumn14.Header.Caption = "Comprobación";
			ultraGridColumn14.Header.VisiblePosition = 17;
			ultraGridColumn14.Width = 120;
			ultraGridColumn15.Header.VisiblePosition = 18;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 19;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 20;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 21;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance26.Image = ((object)(resources.GetObject("appearance26.Image")));
			appearance26.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance26.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn19.CellButtonAppearance = appearance26;
			ultraGridColumn19.Header.Caption = "...";
			ultraGridColumn19.Header.VisiblePosition = 8;
			ultraGridColumn19.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn19.Width = 25;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance27.Image = ((object)(resources.GetObject("appearance27.Image")));
			appearance27.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance27.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn20.CellButtonAppearance = appearance27;
			ultraGridColumn20.Header.Caption = "...";
			ultraGridColumn20.Header.VisiblePosition = 11;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn20.Width = 25;
			ultraGridColumn21.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance28.Image = ((object)(resources.GetObject("appearance28.Image")));
			appearance28.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance28.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn21.CellButtonAppearance = appearance28;
			ultraGridColumn21.Header.Caption = "...";
			ultraGridColumn21.Header.VisiblePosition = 13;
			ultraGridColumn21.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn21.Width = 25;
			ultraGridColumn22.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance29.Image = ((object)(resources.GetObject("appearance29.Image")));
			appearance29.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance29.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn22.CellButtonAppearance = appearance29;
			ultraGridColumn22.Header.Caption = "...";
			ultraGridColumn22.Header.VisiblePosition = 15;
			ultraGridColumn22.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn22.Width = 25;
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
																										 ultraGridColumn22});
			ultraGridBand1.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance30;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance31;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance32;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance33;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance34;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance35;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance36;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance37.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance37;
			summarySettings8.DisplayFormat = "{0: #,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8});
			ultraGridBand1.SummaryFooterCaption = "Totales:";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance38.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance38;
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance40.FontData.BoldAsString = "True";
			appearance40.FontData.Name = "Arial";
			appearance40.FontData.SizeInPoints = 8F;
			appearance40.ForeColor = System.Drawing.Color.White;
			appearance40.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance40;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance42;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 80);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1291, 360);
			this.uGridInformacion.TabIndex = 1017;
			this.uGridInformacion.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridInformacion_BeforeCellUpdate);
			this.uGridInformacion.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_ClickCellButton);
			this.uGridInformacion.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridInformacion_AfterCellUpdate);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn18});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 440);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1307, 8);
			this.groupBox1.TabIndex = 1016;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1307, 8);
			this.groupBox2.TabIndex = 1015;
			this.groupBox2.TabStop = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(520, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 1014;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtestado
			// 
			this.txtestado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtestado.Enabled = false;
			this.txtestado.Location = new System.Drawing.Point(664, 464);
			this.txtestado.MaxValue = 3600;
			this.txtestado.MinValue = 0;
			this.txtestado.Name = "txtestado";
			this.txtestado.PromptChar = ' ';
			this.txtestado.Size = new System.Drawing.Size(48, 22);
			this.txtestado.TabIndex = 1013;
			this.txtestado.Visible = false;
			// 
			// txtidReporteFlujo
			// 
			this.txtidReporteFlujo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidReporteFlujo.Location = new System.Drawing.Point(600, 464);
			this.txtidReporteFlujo.MaxValue = 3600;
			this.txtidReporteFlujo.MinValue = 0;
			this.txtidReporteFlujo.Name = "txtidReporteFlujo";
			this.txtidReporteFlujo.PromptChar = ' ';
			this.txtidReporteFlujo.Size = new System.Drawing.Size(56, 22);
			this.txtidReporteFlujo.TabIndex = 1012;
			this.txtidReporteFlujo.Visible = false;
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(192, 456);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(80, 23);
			this.btnEditar.TabIndex = 1010;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(280, 456);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(80, 23);
			this.btnGuardar.TabIndex = 1009;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(368, 456);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 23);
			this.btnCancelar.TabIndex = 1011;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(16, 456);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 23);
			this.btnNuevo.TabIndex = 1007;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(104, 456);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 1008;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(520, 7);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 1006;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(16, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 17);
			this.label4.TabIndex = 1000;
			this.label4.Text = "Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(216, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 17);
			this.label2.TabIndex = 1002;
			this.label2.Text = "Hasta";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 17);
			this.label1.TabIndex = 1001;
			this.label1.Text = "Desde";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescripcion
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance43;
			this.txtDescripcion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(80, 7);
			this.txtDescripcion.Multiline = true;
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(400, 22);
			this.txtDescripcion.TabIndex = 999;
			this.txtDescripcion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescripcion_KeyDown);
			// 
			// dtHasta
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance44;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(264, 41);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(124, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1004;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtHasta_BeforeDropDown);
			// 
			// dtDesde
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance45;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Enabled = false;
			this.dtDesde.Format = "dd/MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(80, 41);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(124, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1003;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
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
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(920, 40);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(88, 23);
			this.btnGenerar.TabIndex = 1020;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Enabled = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1016, 40);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(88, 23);
			this.button1.TabIndex = 1021;
			this.button1.Text = "Cerrar ";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1112, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 23);
			this.btnExcel.TabIndex = 1022;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnActualizarSaldos
			// 
			this.btnActualizarSaldos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnActualizarSaldos.Enabled = false;
			this.btnActualizarSaldos.Image = ((System.Drawing.Image)(resources.GetObject("btnActualizarSaldos.Image")));
			this.btnActualizarSaldos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnActualizarSaldos.Location = new System.Drawing.Point(1208, 40);
			this.btnActualizarSaldos.Name = "btnActualizarSaldos";
			this.btnActualizarSaldos.Size = new System.Drawing.Size(88, 23);
			this.btnActualizarSaldos.TabIndex = 1023;
			this.btnActualizarSaldos.Text = "Actualizar";
			this.btnActualizarSaldos.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
			this.btnActualizarSaldos.Click += new System.EventHandler(this.btnActualizarSaldos_Click);
			// 
			// frmReporteFlujo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1307, 486);
			this.Controls.Add(this.btnActualizarSaldos);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtestado);
			this.Controls.Add(this.txtidReporteFlujo);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmReporteFlujo";
			this.Text = "Conciliación Bancaria Diaria";
			this.Load += new System.EventHandler(this.frmReporteFlujo_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtestado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidReporteFlujo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
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

		private void frmReporteFlujo_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "040812");

			if (!Funcion.Permiso("1764", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Conciliación bancaria Diaria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;		
			}

			//this.dtDesde.CalendarInfo.MinDate = DateTime.Today;
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Today;
			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			dtFechaMin= Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT  TOP 1 dateadd(day,1,CONVERT(DATETIME,HASTA)) FROM  dbo.MaestroFlujo ORDER  BY  HASTA DESC"));
			//this.dtDesde.CalendarInfo.MinDate = dtFechaMin;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void txtDescripcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtDesde.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Vacia Campos 

			this.txtidReporteFlujo.Value = 0;
			this.txtDescripcion.Text = "";
			this.dtDesde.Value = System.DBNull.Value;
			this.dtHasta.Value = System.DBNull.Value;
			this.txtestado.Value = 0;
			this.lblContador.Text = "";
			this.lblEstado.Text="";

		
			#endregion Vacia Campos 

			#region Bloquea Campo

			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

			#endregion Bloquea Campo

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFlujo 0"));

			this.btnNuevo.Enabled = true;
			this.btnGenerar.Enabled = false;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.button1.Enabled = false;
			this.btnExcel.Enabled = false;
			this.btnActualizarSaldos.Enabled = false;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();	
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
// edicion en estado cerrado
//			if ( (int)this.txtestado.Value == 2)
//			{
//				this.btnExcel.Enabled = true;
//				this.btnGuardar.Enabled = false;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnSaldoBancos"].CellActivation = Activation.AllowEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasCredito"].CellActivation = Activation.AllowEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnDeposito"].CellActivation = Activation.AllowEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasDebito"].CellActivation = Activation.AllowEdit;				
//			}
	
			this.txtDescripcion.Enabled = false;
			this.dtHasta.Enabled = false;
			this.dtDesde.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

//			if ((int) this.txtestado.Value==1)
//			{
//				this.btnActualizarSaldos.Enabled = true;
//				this.button1.Enabled = true;
//				this.btnGuardar.Enabled = true;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnSaldoBancos"].CellActivation = Activation.AllowEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasCredito"].CellActivation = Activation.AllowEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnDeposito"].CellActivation = Activation.AllowEdit;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasDebito"].CellActivation = Activation.AllowEdit;
//			}

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnExcel.Enabled = false;

			this.btnActualizarSaldos.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnGuardar.Enabled = true;					

			bNuevo = false;
			bEdicion = true;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			this.txtDescripcion.Enabled = true;
			this.dtHasta.Enabled = true;
			this.dtDesde.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.uGridInformacion);

			this.txtestado.Value = 1;

			//Estado ((int) this.txtestado.Value);
   
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnCancelar.Enabled = true;	
			this.btnGenerar.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			this.txtDescripcion.Focus();		
		}

		private void Estado ( int IEstado)
		{
			this.lblContador.Text = "";

			if (IEstado == 1)
			{
				this.lblEstado.Text = "ABIERTO";
				this.button1.Enabled = true;
			}
			
			if (IEstado == 2)
			{
				this.lblEstado.Text = "CERRADO";
			}

		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{				
			if (!bEdicion) return;

			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripción")) return;

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
			
			if (Convert.ToDateTime(this.dtDesde.Value) < Convert.ToDateTime(dtFechaMin))
			{
				MessageBox.Show(string.Format("La Fecha Inicial no puede ser menor a la ultima fecha de consignacion ,FECHA = {0}",Convert.ToDateTime(dtFechaMin).ToString("yyyyMMdd")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From MaestroFlujo Where Estado=0")) > 0)
			{
				MessageBox.Show(string.Format("No se puede crear una nueva Conciliación Bancaria mientras se mantengan Conciliaciones Abiertas" ), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de Generar desde  {0}  al {1}", this.dtDesde.Text.ToString(),this.dtHasta.Text.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				string sSQL = string.Format("Exec GeneraFlujo '{0}', '{1}'", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
				FuncionesProcedimientos.dsGeneral(sSQL,uGridInformacion);

				this.lblEstado.Text = "ABIERTO";
			
				MessageBox.Show("Generación Terminada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnGenerar.Enabled = false;
				this.btnActualizarSaldos.Enabled = true;
				this.btnGuardar.Enabled = true;

				this.dtDesde.Enabled = false;
				this.dtHasta.Enabled = false;
			}
		}

		private void uGridInformacion_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			decimal Valor =0.00m;
//			#region Calculos
//			if (e.Cell.Column.ToString() == "SaldoBancos" ||e.Cell.Column.ToString() == "Depositos" ||e.Cell.Column.ToString() == "NotasCredito" ||e.Cell.Column.ToString() == "NotasDebito")
//			{
//				e.Cell.Row.Cells["Comprobacion"].Value = Convert.ToDecimal(e.Cell.Row.Cells["SaldoLibro"].Value)+
//					Convert.ToDecimal(e.Cell.Row.Cells["SaldoCheques"].Value)+
//					Convert.ToDecimal(e.Cell.Row.Cells["NotasCredito"].Value)-
//					Convert.ToDecimal(e.Cell.Row.Cells["NotasDebito"].Value)-
//					Convert.ToDecimal(e.Cell.Row.Cells["Depositos"].Value)-
//					Convert.ToDecimal(e.Cell.Row.Cells["SaldoBancos"].Value); 
//
//				 e.Cell.Row.Cells["SaldoTotales"].Value = Convert.ToDecimal(e.Cell.Row.Cells["SaldoLibro"].Value)+
//							Convert.ToDecimal(e.Cell.Row.Cells["NotasCredito"].Value)-
//							Convert.ToDecimal(e.Cell.Row.Cells["NotasDebito"].Value)-
//							Convert.ToDecimal(e.Cell.Row.Cells["Depositos"].Value); 
//// Comprobacion
//			}
//		#endregion Calculos

			Sumatorias(e);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbTexto(this.txtDescripcion, 5, 100, "Descripción")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Ingrese la fecha Inicial ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Ingrese la fecha Final ", true, "VENTA", cdsSeteoF)) return;
			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtDesde.Value), Convert.ToDateTime(this.dtHasta.Value), this.dtDesde, "La Fecha Inicial no puede ser mayor a la Fecha Final", false)) return;

			#endregion Validacion
      
			
			#region Guarda Maestro
			string sSQLl = string.Format("Exec GuardaMaestroFlujo {0}, '{1}' , '{2}' ,'{3}', {4}",
				(int)this.txtidReporteFlujo.Value,
				this.txtDescripcion.Text,
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				(int)this.txtestado.Value);
			#endregion Guarda Maestro

			this.txtidReporteFlujo.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQLl);

			#region Guarda detalle Maestro
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridInformacion.Rows.All)
			{	

				string sSQLDetalle = string.Format("Exec GuardaDetalleFlujo  {0}, {1}, {2}, {3}, '{4}', '{5}',{6}, {7},{8}, {9},{10},{11},{12},{13},'{14}','{15}','{16}','{17}'", 
					(int)drr.Cells["idDetalleFlujo"].Value,
					(int)this.txtidReporteFlujo.Value,
					(int)drr.Cells["idBanco"].Value,
					(int)drr.Cells["idCuenta"].Value,
					drr.Cells["Banco"].Value.ToString(),
					drr.Cells["Cuenta"].Value.ToString(),
					(decimal)drr.Cells["SaldoLibro"].Value,
					(decimal)drr.Cells["SaldoBancos"].Value,
					(decimal)drr.Cells["SaldoCheques"].Value,
					(decimal)drr.Cells["Depositos"].Value,
					(decimal)drr.Cells["NotasCredito"].Value,
					(decimal)drr.Cells["NotasDebito"].Value,
					(decimal)drr.Cells["SaldoTotales"].Value,
					(decimal)drr.Cells["Comprobacion"].Value,
					drr.Cells["ObservacionesSaldoCuenta"].Value.ToString(),
					drr.Cells["ObservacionesDeposito"].Value.ToString(),
					drr.Cells["ObservacionesNotasCredito"].Value.ToString(),
					drr.Cells["ObservacionesNotasDebito"].Value.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
			}

			#endregion Guarda detalle Maestro

			#region Controles
			Consultar ((int)this.txtidReporteFlujo.Value);
			this.txtDescripcion.Enabled = false;
			this.dtDesde.Enabled = false;
			this.dtHasta.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.uGridInformacion);

//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnSaldoBancos"].CellActivation = Activation.Disabled;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasCredito"].CellActivation = Activation.Disabled;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnDeposito"].CellActivation = Activation.Disabled;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasDebito"].CellActivation = Activation.Disabled;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnExcel.Enabled = true;
			this.btnActualizarSaldos.Enabled = false;
			this.btnGenerar.Enabled = false;
			this.button1.Enabled = true;
		
			bNuevo = false;
			bEdicion = false;
			#endregion Controles	

		}

		private void dtHasta_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaFlujo consultab = new frmConsultaFlujo())
				if (DialogResult.OK == consultab.ShowDialog())
					this.Consultar((int)consultab.uGridInformacion.ActiveRow.Cells["idMaestroFlujo"].Value);
		}
		private void Consultar (int IdMaestroFlujo)
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualMaestroFlujo {0}",IdMaestroFlujo); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtidReporteFlujo.Value= dr.GetInt32 (0);
			if (dr.GetValue(1) != System.DBNull.Value)this.txtDescripcion.Text = dr.GetString (1);
			if (dr.GetValue(2) != System.DBNull.Value)this.dtDesde.Value = dr.GetDateTime (2);
			if (dr.GetValue(3) != System.DBNull.Value)this.dtHasta.Value = dr.GetDateTime (3);
			if (dr.GetValue(4) != System.DBNull.Value)this.txtestado.Value = dr.GetInt32 (4);			
			dr.Close();

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFlujo {0}", IdMaestroFlujo));
			#endregion Consulta Individual

			Estado ((int)this.txtestado.Value);
			cmdTotal();



	//if (e.Cell.Column.ToString() == "SaldoBancos" ||e.Cell.Column.ToString() == "Depositos" ||e.Cell.Column.ToString() == "NotasCredito" ||e.Cell.Column.ToString() == "NotasDebito")
		
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnSaldoBancos"].CellActivation = Activation.Disabled;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasCredito"].CellActivation = Activation.Disabled;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnDeposito"].CellActivation = Activation.Disabled;
//				this.uGridInformacion.DisplayLayout.Bands[0].Columns["btnNotasDebito"].CellActivation = Activation.Disabled;

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnExcel.Enabled = true;

			if ((int)this.txtestado.Value == 1) 
			{
				this.button1.Enabled = true;
				this.btnEditar.Enabled = true;
			}
			
			this.btnCancelar.Enabled = true;

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			string sMensaje = "Cerrar";

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{	
				if (Convert.ToDecimal(dr.Cells["Comprobacion"].Value) >0.00m )
				{
					MessageBox.Show("No se puede cerrar la Conciliación bancaria con diferencias", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}	
			}

			if (DialogResult.Yes ==	MessageBox.Show(string.Format("¿Esta seguro de '{0}' esta Conciliación bancaria , recuerde que  este proceso es irreversible ", sMensaje), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{

				string sSQL = string.Format("Exec CierreFlujo {0}", this.txtidReporteFlujo.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);
				
				this.lblEstado.Text = "CERRADO";
				this.txtestado.Value = 2;
				#region Controles
				this.btnNuevo.Enabled = true;
				this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnCancelar.Enabled = true;	
				this.button1.Enabled = false;	
				this.btnActualizarSaldos.Enabled = false;
				#endregion Controles	
				
				MessageBox.Show("Proceso terminado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			}
		}

		private void uGridInformacion_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//dt.ObservacionesSaldoBancos,dt.ObservacionesDeposito,dt.ObservacionesNotasCredito,dt.ObservacionesNotasDebito
			#region SaldoBancoss
			if (e.Cell.Column.ToString() == "btnSaldoBancos")
			{	
				using (frmObservacionesFlujo Observaciones = new frmObservacionesFlujo(1, e.Cell.Row.Cells["ObservacionesSaldoCuenta"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["ObservacionesSaldoCuenta"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion SaldoBancoss

			#region Deposito
			if (e.Cell.Column.ToString() == "btnDeposito")
			{	
				using (frmObservacionesFlujo Observaciones = new frmObservacionesFlujo(2, e.Cell.Row.Cells["ObservacionesDeposito"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["ObservacionesDeposito"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion Deposito

			#region NotasCredito
			if (e.Cell.Column.ToString() == "btnNotasCredito")
			{	
				using (frmObservacionesFlujo Observaciones = new frmObservacionesFlujo(3, e.Cell.Row.Cells["ObservacionesNotasCredito"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["ObservacionesNotasCredito"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion NotasCredito

			#region NotasDebito
			if (e.Cell.Column.ToString() == "btnNotasDebito")
			{	
				using (frmObservacionesFlujo Observaciones = new frmObservacionesFlujo(4, e.Cell.Row.Cells["ObservacionesNotasDebito"].Value.ToString().Trim(), bEdicion))
				{				
					if (DialogResult.OK == Observaciones.ShowDialog())
					{
						e.Cell.Row.Cells["ObservacionesNotasDebito"].Value = Observaciones.txtObservaciones.Text.ToString().Trim();
					}
				}
			}
			#endregion NotasDebito
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void btnActualizarSaldos_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Actualizar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3))
			{	
				string sSQLConsulta = string.Format("Exec ActualizaConsignacion  '{0}', '{1}', {2}", 
					Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"),
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
					(int)this.txtidReporteFlujo.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLConsulta);

				MessageBox.Show("Valores Actualizados Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaIndividualDetalleFlujo {0}", (int)this.txtidReporteFlujo.Value));

				
			}
		}
		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Calculos
				e.Cell.Row.Cells["Comprobacion"].Value = Convert.ToDecimal(e.Cell.Row.Cells["SaldoLibro"].Value)+
					Convert.ToDecimal(e.Cell.Row.Cells["SaldoCheques"].Value)+
					Convert.ToDecimal(e.Cell.Row.Cells["NotasCredito"].Value)-
					Convert.ToDecimal(e.Cell.Row.Cells["NotasDebito"].Value)-
					Convert.ToDecimal(e.Cell.Row.Cells["Depositos"].Value)-
					Convert.ToDecimal(e.Cell.Row.Cells["SaldoBancos"].Value); 
				e.Cell.Row.Cells["SaldoTotales"].Value= Convert.ToDecimal(e.Cell.Row.Cells["SaldoLibro"].Value)+
					Convert.ToDecimal(e.Cell.Row.Cells["NotasCredito"].Value)-
					Convert.ToDecimal(e.Cell.Row.Cells["NotasDebito"].Value)-
					Convert.ToDecimal(e.Cell.Row.Cells["Depositos"].Value); 
			#endregion Calculos		
		}

		private void uGridInformacion_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
		}
		private void cmdTotal()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in uGridInformacion.Rows.All)
			{
					drr.Cells["Comprobacion"].Value = Convert.ToDecimal(drr.Cells["SaldoLibro"].Value)+
						Convert.ToDecimal(drr.Cells["SaldoCheques"].Value)+
						Convert.ToDecimal(drr.Cells["NotasCredito"].Value)-
						Convert.ToDecimal(drr.Cells["NotasDebito"].Value)-
						Convert.ToDecimal(drr.Cells["Depositos"].Value)-
						Convert.ToDecimal(drr.Cells["SaldoBancos"].Value); 
				drr.Cells["SaldoTotales"].Value= Convert.ToDecimal(drr.Cells["SaldoLibro"].Value)+
					Convert.ToDecimal(drr.Cells["NotasCredito"].Value)-
					Convert.ToDecimal(drr.Cells["NotasDebito"].Value)-
					Convert.ToDecimal(drr.Cells["Depositos"].Value); 
			}
		}
	}
}
