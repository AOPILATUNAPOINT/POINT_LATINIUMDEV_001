using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_PresupuestoSemanal.
	/// </summary>
	public class frmCom_PresupuestoSemanal : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSemanal;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSemanalContado;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSemanalGEx;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridSemanalCredito;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.MenuItem mnuAccionesGenerarPeriodo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_PresupuestoSemanal()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_PresupuestoSemanal));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PresupuestoSemanal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaContado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaCredito");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MetaGEx");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Meta", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Meta", 3, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MetaContado", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MetaContado", 4, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MetaCredito", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MetaCredito", 5, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MetaGEx", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MetaGEx", 6, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PresupuestoSemanal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PresupuestoSemanalContado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PresupuestoSemanalContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Meta", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Meta", 3, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PresupuestoSemanalGEx");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Meta", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Meta", 3, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PresupuestoSemanalGEx");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_PresupuestoSemanalCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Semana");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_PresupuestoSemanalCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Semana");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Meta", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Meta", 3, true);
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			this.lblBodega = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemanal = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemanalContado = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.uGridSemanalGEx = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridSemanalCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesGenerarPeriodo = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.btnExportar = new System.Windows.Forms.Button();
			this.btnComandos = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanalContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanalGEx)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanalCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 649;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(192, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 648;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 42);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 647;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Enabled = false;
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(248, 40);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 646;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton2.Caption = "Today";
			dateButton2.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(64, 40);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 645;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 60;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 288;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(64, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(288, 21);
			this.cmbBodega.TabIndex = 650;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3});
			// 
			// uGridSemanal
			// 
			this.uGridSemanal.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridSemanal.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemanal.DataSource = this.ultraDataSource1;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridSemanal.DisplayLayout.Appearance = appearance3;
			this.uGridSemanal.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 200;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Width = 60;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Meta Total";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Width = 100;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "Meta Contado";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Width = 100;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance7;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Meta Credito";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Width = 100;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance8;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Meta G Ex";
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Width = 100;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			ultraGridBand2.Header.Caption = "Presupuesto Semanal";
			ultraGridBand2.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance9;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance10;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance11;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance12;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridSemanal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridSemanal.DisplayLayout.GroupByBox.Hidden = true;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanal.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.uGridSemanal.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSemanal.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSemanal.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemanal.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanal.DisplayLayout.Override.CellAppearance = appearance15;
			this.uGridSemanal.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemanal.DisplayLayout.Override.HeaderAppearance = appearance16;
			this.uGridSemanal.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanal.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanal.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanal.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.uGridSemanal.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridSemanal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemanal.Location = new System.Drawing.Point(8, 80);
			this.uGridSemanal.Name = "uGridSemanal";
			this.uGridSemanal.Size = new System.Drawing.Size(681, 480);
			this.uGridSemanal.TabIndex = 682;
			this.uGridSemanal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridSemanal_KeyDown);
			this.uGridSemanal.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridSemanal_AfterCellUpdate);
			this.uGridSemanal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemanal_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// uGridSemanalContado
			// 
			this.uGridSemanalContado.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemanalContado.DataSource = this.ultraDataSource3;
			appearance20.BackColor = System.Drawing.Color.White;
			this.uGridSemanalContado.DisplayLayout.Appearance = appearance20;
			this.uGridSemanalContado.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 200;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance21;
			ultraGridColumn13.Header.VisiblePosition = 2;
			ultraGridColumn13.Width = 60;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance22;
			ultraGridColumn14.Format = ".";
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn14.Width = 100;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			ultraGridBand3.Header.Caption = "Presupuesto Semanal Contado";
			ultraGridBand3.HeaderVisible = true;
			ultraGridBand3.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance23;
			summarySettings5.DisplayFormat = "{0: #,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings5});
			ultraGridBand3.SummaryFooterCaption = "";
			this.uGridSemanalContado.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.uGridSemanalContado.DisplayLayout.GroupByBox.Hidden = true;
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanalContado.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.uGridSemanalContado.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSemanalContado.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSemanalContado.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemanalContado.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanalContado.DisplayLayout.Override.CellAppearance = appearance26;
			this.uGridSemanalContado.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 8F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemanalContado.DisplayLayout.Override.HeaderAppearance = appearance27;
			this.uGridSemanalContado.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance28.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance28.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalContado.DisplayLayout.Override.RowAlternateAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalContado.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalContado.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.uGridSemanalContado.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridSemanalContado.Enabled = false;
			this.uGridSemanalContado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemanalContado.Location = new System.Drawing.Point(720, 80);
			this.uGridSemanalContado.Name = "uGridSemanalContado";
			this.uGridSemanalContado.Size = new System.Drawing.Size(381, 480);
			this.uGridSemanalContado.TabIndex = 683;
			this.uGridSemanalContado.Visible = false;
			this.uGridSemanalContado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemanalContado_InitializeLayout);
			// 
			// uGridSemanalGEx
			// 
			this.uGridSemanalGEx.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemanalGEx.DataSource = this.ultraDataSource4;
			appearance31.BackColor = System.Drawing.Color.White;
			this.uGridSemanalGEx.DisplayLayout.Appearance = appearance31;
			this.uGridSemanalGEx.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 200;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn17.CellAppearance = appearance32;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Width = 60;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance33;
			ultraGridColumn18.Format = ".";
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Width = 100;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			ultraGridBand4.Header.Caption = "Presupuesto Semanal Garantia Extendida";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance34;
			summarySettings6.DisplayFormat = "{0: #,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings6});
			ultraGridBand4.SummaryFooterCaption = "";
			this.uGridSemanalGEx.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.uGridSemanalGEx.DisplayLayout.GroupByBox.Hidden = true;
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanalGEx.DisplayLayout.Override.ActiveRowAppearance = appearance35;
			this.uGridSemanalGEx.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSemanalGEx.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSemanalGEx.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance36.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemanalGEx.DisplayLayout.Override.CardAreaAppearance = appearance36;
			appearance37.ForeColor = System.Drawing.Color.Black;
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanalGEx.DisplayLayout.Override.CellAppearance = appearance37;
			this.uGridSemanalGEx.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance38.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance38.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance38.FontData.BoldAsString = "True";
			appearance38.FontData.Name = "Arial";
			appearance38.FontData.SizeInPoints = 8F;
			appearance38.ForeColor = System.Drawing.Color.White;
			appearance38.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemanalGEx.DisplayLayout.Override.HeaderAppearance = appearance38;
			this.uGridSemanalGEx.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance39.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance39.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalGEx.DisplayLayout.Override.RowAlternateAppearance = appearance39;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalGEx.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalGEx.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.uGridSemanalGEx.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridSemanalGEx.Enabled = false;
			this.uGridSemanalGEx.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemanalGEx.Location = new System.Drawing.Point(1184, 80);
			this.uGridSemanalGEx.Name = "uGridSemanalGEx";
			this.uGridSemanalGEx.Size = new System.Drawing.Size(381, 480);
			this.uGridSemanalGEx.TabIndex = 685;
			this.uGridSemanalGEx.Visible = false;
			this.uGridSemanalGEx.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemanalGEx_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// uGridSemanalCredito
			// 
			this.uGridSemanalCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridSemanalCredito.DataSource = this.ultraDataSource5;
			appearance42.BackColor = System.Drawing.Color.White;
			this.uGridSemanalCredito.DisplayLayout.Appearance = appearance42;
			this.uGridSemanalCredito.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 200;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn21.CellAppearance = appearance43;
			ultraGridColumn21.Header.VisiblePosition = 2;
			ultraGridColumn21.Width = 60;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance44;
			ultraGridColumn22.Format = ".";
			ultraGridColumn22.Header.VisiblePosition = 3;
			ultraGridColumn22.Width = 100;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			ultraGridBand5.Header.Caption = "Presupuesto Semanal Credito";
			ultraGridBand5.HeaderVisible = true;
			ultraGridBand5.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance45.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance45;
			summarySettings7.DisplayFormat = "{0: #,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings7});
			ultraGridBand5.SummaryFooterCaption = "";
			this.uGridSemanalCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.uGridSemanalCredito.DisplayLayout.GroupByBox.Hidden = true;
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanalCredito.DisplayLayout.Override.ActiveRowAppearance = appearance46;
			this.uGridSemanalCredito.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridSemanalCredito.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridSemanalCredito.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance47.BackColor = System.Drawing.Color.Transparent;
			this.uGridSemanalCredito.DisplayLayout.Override.CardAreaAppearance = appearance47;
			appearance48.ForeColor = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridSemanalCredito.DisplayLayout.Override.CellAppearance = appearance48;
			this.uGridSemanalCredito.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance49.FontData.BoldAsString = "True";
			appearance49.FontData.Name = "Arial";
			appearance49.FontData.SizeInPoints = 8F;
			appearance49.ForeColor = System.Drawing.Color.White;
			appearance49.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridSemanalCredito.DisplayLayout.Override.HeaderAppearance = appearance49;
			this.uGridSemanalCredito.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance50.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance50.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance50.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalCredito.DisplayLayout.Override.RowAlternateAppearance = appearance50;
			appearance51.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance51.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance51.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalCredito.DisplayLayout.Override.RowSelectorAppearance = appearance51;
			appearance52.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance52.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance52.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridSemanalCredito.DisplayLayout.Override.SelectedRowAppearance = appearance52;
			this.uGridSemanalCredito.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridSemanalCredito.Enabled = false;
			this.uGridSemanalCredito.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridSemanalCredito.Location = new System.Drawing.Point(760, 80);
			this.uGridSemanalCredito.Name = "uGridSemanalCredito";
			this.uGridSemanalCredito.Size = new System.Drawing.Size(381, 480);
			this.uGridSemanalCredito.TabIndex = 684;
			this.uGridSemanalCredito.Visible = false;
			this.uGridSemanalCredito.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridSemanalCredito_InitializeLayout);
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
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(368, 39);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 726;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 65);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(696, 8);
			this.groupBox1.TabIndex = 727;
			this.groupBox1.TabStop = false;
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesGenerarPeriodo,
																																								this.mnuAccionesAgregarNuevoLocal});
			// 
			// mnuAccionesGenerarPeriodo
			// 
			this.mnuAccionesGenerarPeriodo.Index = 0;
			this.mnuAccionesGenerarPeriodo.Text = "Generar Nuevo Periodo Semanal";
			this.mnuAccionesGenerarPeriodo.Click += new System.EventHandler(this.mnuAccionesGenerarPeriodo_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Agregar Nuevo Local";
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(616, 39);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.TabIndex = 734;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(616, 7);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 733;
			this.btnComandos.Text = "Acciones";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// frmCom_PresupuestoSemanal
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(695, 566);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.uGridSemanal);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.uGridSemanalContado);
			this.Controls.Add(this.uGridSemanalGEx);
			this.Controls.Add(this.uGridSemanalCredito);
			this.Controls.Add(this.btnExportar);
			this.Name = "frmCom_PresupuestoSemanal";
			this.Text = "Presupuesto Semanal";
			this.Load += new System.EventHandler(this.frmCom_PresupuestoSemanal_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanalContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanalGEx)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridSemanalCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_PresupuestoSemanal_Load(object sender, System.EventArgs e)
		{
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 And Credito = 1 Order By Nombre");
						
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("04/06/2018");

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Exec Com_RetornaPeriodosSemanales", true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.dtDesde.Value = dr.GetDateTime(0);
				this.dtHasta.Value = dr.GetDateTime(1);
			}
			dr.Close();

			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void Consulta()
		{
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec Com_ConsultaPresupuestoSemanal {0}, '{1}', '{2}', 1", 
				iBodega, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridSemanal);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			this.Consulta();	

//			string sSQL = string.Format("Exec Com_ConsultaPresupuestoSemanal {0}, '{1}', '{2}', 2", 
//				iBodega, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
//			this.uGridSemanalContado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
//
//			string sSQL = string.Format("Exec Com_ConsultaPresupuestoSemanal {0}, '{1}', '{2}', 3", 
//				iBodega, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
//			this.uGridSemanalCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
//
//			string sSQL = string.Format("Exec Com_ConsultaPresupuestoSemanal {0}, '{1}', '{2}', 4", 
//				iBodega, Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
//			this.uGridSemanalGEx.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void uGridSemanal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Meta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MetaContado", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MetaCredito", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MetaGEx", 2);
		}

		private void uGridSemanalContado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Meta", 2);
		}

		private void uGridSemanalCredito_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Meta", 2);
		}

		private void uGridSemanalGEx_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Meta", 2);
		}

		private void uGridSemanal_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Meta")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Com_PresupuestoSemanal Set Meta = {0} Where idCom_PresupuestoSemanal = {1}", 
					Convert.ToDecimal(e.Cell.Row.Cells["Meta"].Value), (int)e.Cell.Row.Cells["idCom_PresupuestoSemanal"].Value));
			}

			if (e.Cell.Column.ToString() == "MetaContado")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Com_PresupuestoSemanal Set MetaContado = {0} Where idCom_PresupuestoSemanal = {1}", 
					Convert.ToDecimal(e.Cell.Row.Cells["MetaContado"].Value), (int)e.Cell.Row.Cells["idCom_PresupuestoSemanal"].Value));

//				e.Cell.Row.Cells["MetaCredito"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Meta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["MetaContado"].Value);
			}

			if (e.Cell.Column.ToString() == "MetaCredito")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Com_PresupuestoSemanal Set MetaCredito = {0} Where idCom_PresupuestoSemanal = {1}", 
					Convert.ToDecimal(e.Cell.Row.Cells["MetaCredito"].Value), (int)e.Cell.Row.Cells["idCom_PresupuestoSemanal"].Value));

//				e.Cell.Row.Cells["MetaContado"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Meta"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["MetaCredito"].Value);
			}

			if (e.Cell.Column.ToString() == "MetaGEx")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Com_PresupuestoSemanal Set MetaGEx = {0} Where idCom_PresupuestoSemanal = {1}",
					Convert.ToDecimal(e.Cell.Row.Cells["MetaGEx"].Value), (int)e.Cell.Row.Cells["idCom_PresupuestoSemanal"].Value));
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridSemanal);
		}

		private void mnuAccionesGenerarPeriodo_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de generar un nuevo periodo semanal?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{				
				Funcion.EjecutaSQL(cdsSeteoF, "Exec Com_GeneraPresupuestoSemanal", true);

				MessageBox.Show("Periodo semanal generado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			
				this.btnVer_Click(sender, e);
			}
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value) 
				this.dtHasta.Value = Convert.ToDateTime(this.dtDesde.Value).AddDays(6);
		}

		private void uGridSemanal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridSemanal);
		}
	}
}
