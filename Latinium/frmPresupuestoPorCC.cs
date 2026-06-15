using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoPorCC.
	/// </summary>
	public class frmPresupuestoPorCC : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAnio;
		private Infragistics.Win.Misc.UltraLabel lblPeriodo;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidPresupuestoPorCC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbResponsable;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCuentas;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugrValores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroDeCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnSubir;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button btnAprobar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnGenerar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoPorCC()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoPorCC));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoPorCC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TOTAL");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ENERO");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FEBRERO");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MARZO");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ABRIL");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MAYO");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JUNIO");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("JULIO");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AGOSTO");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SEPTIEMBRE");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OCTUBRE");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NOVIEMBRE");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DICIEMBRE");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dupli");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ENERO", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ENERO", 5, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "FEBRERO", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "FEBRERO", 6, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MARZO", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MARZO", 7, true);
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ABRIL", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ABRIL", 8, true);
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "MAYO", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "MAYO", 9, true);
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JUNIO", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JUNIO", 10, true);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "JULIO", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "JULIO", 11, true);
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "AGOSTO", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "AGOSTO", 12, true);
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SEPTIEMBRE", 13, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SEPTIEMBRE", 13, true);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "OCTUBRE", 14, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "OCTUBRE", 14, true);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "NOVIEMBRE", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "NOVIEMBRE", 15, true);
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "DICIEMBRE", 16, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "DICIEMBRE", 16, true);
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TOTAL", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TOTAL", 4, true);
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPresupuestoPorCC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TOTAL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ENERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FEBRERO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MARZO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ABRIL");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MAYO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JUNIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("JULIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AGOSTO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SEPTIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OCTUBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NOVIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DICIEMBRE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dupli");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.txtAnio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblPeriodo = new Infragistics.Win.Misc.UltraLabel();
			this.label6 = new System.Windows.Forms.Label();
			this.cmbResponsable = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.txtidPresupuestoPorCC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAprobar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridCuentas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ugrValores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCentroDeCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnSubir = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnGenerar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPresupuestoPorCC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrValores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// txtAnio
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAnio.Appearance = appearance1;
			this.txtAnio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAnio.Enabled = false;
			this.txtAnio.FormatString = "";
			this.txtAnio.Location = new System.Drawing.Point(96, 6);
			this.txtAnio.MaskInput = "nnnn";
			this.txtAnio.Name = "txtAnio";
			this.txtAnio.PromptChar = ' ';
			this.txtAnio.Size = new System.Drawing.Size(96, 21);
			this.txtAnio.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtAnio.TabIndex = 759;
			this.txtAnio.ValueChanged += new System.EventHandler(this.txtAnio_ValueChanged);
			// 
			// lblPeriodo
			// 
			this.lblPeriodo.AutoSize = true;
			this.lblPeriodo.BackColor = System.Drawing.Color.Transparent;
			this.lblPeriodo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblPeriodo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.lblPeriodo.Location = new System.Drawing.Point(8, 9);
			this.lblPeriodo.Name = "lblPeriodo";
			this.lblPeriodo.Size = new System.Drawing.Size(42, 15);
			this.lblPeriodo.TabIndex = 758;
			this.lblPeriodo.Text = "Periodo";
			this.lblPeriodo.Click += new System.EventHandler(this.lblPeriodo_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 42);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(86, 16);
			this.label6.TabIndex = 996;
			this.label6.Text = "Centro de Costo";
			// 
			// cmbResponsable
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbResponsable.Appearance = appearance2;
			this.cmbResponsable.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbResponsable.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbResponsable.DataSource = this.ultraDataSource1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 250;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbResponsable.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbResponsable.DisplayMember = "Nombre";
			this.cmbResponsable.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbResponsable.Enabled = false;
			this.cmbResponsable.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbResponsable.Location = new System.Drawing.Point(960, 376);
			this.cmbResponsable.Name = "cmbResponsable";
			this.cmbResponsable.Size = new System.Drawing.Size(24, 21);
			this.cmbResponsable.TabIndex = 999;
			this.cmbResponsable.ValueMember = "idNomina";
			this.cmbResponsable.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(744, 384);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(70, 16);
			this.label1.TabIndex = 998;
			this.label1.Text = "Responsable";
			this.label1.Visible = false;
			// 
			// txtidPresupuestoPorCC
			// 
			this.txtidPresupuestoPorCC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidPresupuestoPorCC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidPresupuestoPorCC.Enabled = false;
			this.txtidPresupuestoPorCC.Location = new System.Drawing.Point(880, 384);
			this.txtidPresupuestoPorCC.Name = "txtidPresupuestoPorCC";
			this.txtidPresupuestoPorCC.PromptChar = ' ';
			this.txtidPresupuestoPorCC.Size = new System.Drawing.Size(72, 21);
			this.txtidPresupuestoPorCC.TabIndex = 1000;
			this.txtidPresupuestoPorCC.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(848, 376);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(72, 21);
			this.txtEstado.TabIndex = 1001;
			this.txtEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(208, 8);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 16);
			this.lblEstado.TabIndex = 1002;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 376);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 1003;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 376);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 1004;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(168, 376);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 1006;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(248, 376);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 1005;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAprobar
			// 
			this.btnAprobar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAprobar.CausesValidation = false;
			this.btnAprobar.Enabled = false;
			this.btnAprobar.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobar.Image")));
			this.btnAprobar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobar.Location = new System.Drawing.Point(488, 376);
			this.btnAprobar.Name = "btnAprobar";
			this.btnAprobar.Size = new System.Drawing.Size(76, 23);
			this.btnAprobar.TabIndex = 1007;
			this.btnAprobar.Text = "&Aprobar";
			this.btnAprobar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobar.Visible = false;
			this.btnAprobar.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(568, 376);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 1008;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(648, 376);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1009;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox2.Location = new System.Drawing.Point(-152, 360);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1336, 8);
			this.groupBox2.TabIndex = 1010;
			this.groupBox2.TabStop = false;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(8, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(368, 8);
			this.groupBox1.TabIndex = 1011;
			this.groupBox1.TabStop = false;
			// 
			// uGridCuentas
			// 
			this.uGridCuentas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.uGridCuentas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCuentas.DataSource = this.ultraDataSource4;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridCuentas.DisplayLayout.Appearance = appearance3;
			ultraGridColumn3.Header.Caption = "...";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 52;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 220;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			ultraGridBand2.Header.Caption = "Gastos";
			ultraGridBand2.HeaderVisible = true;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridCuentas.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCuentas.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			this.uGridCuentas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridCuentas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridCuentas.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCuentas.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridCuentas.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCuentas.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridCuentas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCuentas.Location = new System.Drawing.Point(8, 88);
			this.uGridCuentas.Name = "uGridCuentas";
			this.uGridCuentas.Size = new System.Drawing.Size(368, 264);
			this.uGridCuentas.TabIndex = 1012;
			this.uGridCuentas.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridCuentas_CellChange);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(bool);
			ultraDataColumn4.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			this.ultraDataSource4.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource4_CellDataRequested);
			// 
			// ugrValores
			// 
			this.ugrValores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugrValores.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugrValores.DataSource = this.ultraDataSource6;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ugrValores.DisplayLayout.Appearance = appearance10;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Width = 230;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance11;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 16;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance12;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 4;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance13;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 5;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance14;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 6;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance15;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 7;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance16;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.VisiblePosition = 8;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance17;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.VisiblePosition = 9;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance18;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 10;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance19;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 11;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance20;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 12;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance21;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.VisiblePosition = 13;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance22;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.VisiblePosition = 14;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance23;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 15;
			ultraGridColumn24.Header.VisiblePosition = 17;
			ultraGridColumn24.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn24});
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance24;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance25;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance26.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance26;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance27.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance27;
			summarySettings4.DisplayFormat = "{0:#,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance28;
			summarySettings5.DisplayFormat = "{0:#,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance29;
			summarySettings6.DisplayFormat = "{0:#,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance30;
			summarySettings7.DisplayFormat = "{0:#,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance31;
			summarySettings8.DisplayFormat = "{0:#,##0.00}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings9.Appearance = appearance32;
			summarySettings9.DisplayFormat = "{0:#,##0.00}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance33.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings10.Appearance = appearance33;
			summarySettings10.DisplayFormat = "{0:#,##0.00}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance34.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings11.Appearance = appearance34;
			summarySettings11.DisplayFormat = "{0:#,##0.00}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance35.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings12.Appearance = appearance35;
			summarySettings12.DisplayFormat = "{0:#,##0.00}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance36.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings13.Appearance = appearance36;
			summarySettings13.DisplayFormat = "{0:#,##0.00}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand3.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13});
			ultraGridBand3.SummaryFooterCaption = "";
			this.ugrValores.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ugrValores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugrValores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.ugrValores.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugrValores.DisplayLayout.Override.CellAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 8F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugrValores.DisplayLayout.Override.HeaderAppearance = appearance39;
			this.ugrValores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance40.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance40.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrValores.DisplayLayout.Override.RowSelectorAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugrValores.DisplayLayout.Override.SelectedRowAppearance = appearance41;
			this.ugrValores.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugrValores.Location = new System.Drawing.Point(384, 8);
			this.ugrValores.Name = "ugrValores";
			this.ugrValores.Size = new System.Drawing.Size(800, 344);
			this.ugrValores.TabIndex = 1013;
			this.ugrValores.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugrValores_BeforeRowsDeleted);
			this.ugrValores.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugrValores_KeyDown);
			this.ugrValores.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugrValores_AfterCellUpdate);
			this.ugrValores.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugrValores_InitializeLayout);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn8.DataType = typeof(int);
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
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// cmbCentroDeCosto
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentroDeCosto.Appearance = appearance42;
			this.cmbCentroDeCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroDeCosto.DataSource = this.ultraDataSource3;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 250;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 250;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.cmbCentroDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCentroDeCosto.DisplayMember = "Nombre";
			this.cmbCentroDeCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroDeCosto.Enabled = false;
			this.cmbCentroDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroDeCosto.Location = new System.Drawing.Point(96, 40);
			this.cmbCentroDeCosto.Name = "cmbCentroDeCosto";
			this.cmbCentroDeCosto.Size = new System.Drawing.Size(256, 21);
			this.cmbCentroDeCosto.TabIndex = 1014;
			this.cmbCentroDeCosto.ValueMember = "idProyecto";
			this.cmbCentroDeCosto.ValueChanged += new System.EventHandler(this.cmbCentroDeCosto_ValueChanged);
			this.cmbCentroDeCosto.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCentroDeCosto_InitializeLayout);
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
			// btnSubir
			// 
			this.btnSubir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSubir.CausesValidation = false;
			this.btnSubir.Enabled = false;
			this.btnSubir.Image = ((System.Drawing.Image)(resources.GetObject("btnSubir.Image")));
			this.btnSubir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSubir.Location = new System.Drawing.Point(328, 376);
			this.btnSubir.Name = "btnSubir";
			this.btnSubir.Size = new System.Drawing.Size(76, 23);
			this.btnSubir.TabIndex = 1015;
			this.btnSubir.Text = "Importar";
			this.btnSubir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSubir.Click += new System.EventHandler(this.btnSubir_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button1.CausesValidation = false;
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(1104, 376);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(76, 23);
			this.button1.TabIndex = 1016;
			this.button1.Text = "&Ayuda ";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(408, 376);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 1017;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click_1);
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(1016, 376);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(76, 23);
			this.btnGenerar.TabIndex = 1018;
			this.btnGenerar.Text = "Periodo";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Visible = false;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// frmPresupuestoPorCC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1192, 414);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.btnSubir);
			this.Controls.Add(this.cmbCentroDeCosto);
			this.Controls.Add(this.ugrValores);
			this.Controls.Add(this.uGridCuentas);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAprobar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtidPresupuestoPorCC);
			this.Controls.Add(this.cmbResponsable);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtAnio);
			this.Controls.Add(this.lblPeriodo);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoPorCC";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Presupuesto por Departamento";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoPorCC_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoPorCC_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtAnio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbResponsable)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidPresupuestoPorCC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridCuentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugrValores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void lblPeriodo_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtAnio_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void UnloadMe()
		{
			this.Close();
		}
		private Acceso miAcceso;
		bool bNuevo = false;
		bool bEdicion = false;




		public  bool ValidaPeriodo (  string Mes, int iYear )
		{
			string sPeriodoFecha ="";
			bool bVeirifca = false;
			#region ValidaMeses
			if (Mes == "ENERO")  sPeriodoFecha =iYear +""+"-01-01";
			if (Mes == "FEBRERO")  sPeriodoFecha =iYear +""+"-02-01";
			if (Mes == "MARZO")  sPeriodoFecha =iYear +""+"-03-01";
			if (Mes == "ABRIL")  sPeriodoFecha =iYear +""+"-04-01";
			if (Mes == "MAYO")  sPeriodoFecha =iYear+""+"-05-01";
			if (Mes == "JUNIO")  sPeriodoFecha =iYear +""+"-06-01";
			if (Mes == "JULIO")  sPeriodoFecha =iYear +""+"-07-01";
			if (Mes == "AGOSTO")  sPeriodoFecha =iYear +""+"-08-01";
			if (Mes == "SEPTIEMBRE")  sPeriodoFecha =iYear +""+"-09-01";
			if (Mes == "OCTUBRE")  sPeriodoFecha =iYear +""+"-10-01";
			if (Mes == "NOVIEMBRE")  sPeriodoFecha =iYear +""+"-11-01";
			if (Mes == "DICIEMBRE")  sPeriodoFecha =iYear +""+"-12-01";

			#endregion ValidaMeses
			if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
			{
				bVeirifca = true;
			}
			return bVeirifca;
		}

		private void frmPresupuestoPorCC_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2605");
			
			if (!Funcion.Permiso("1706", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Presupuesto Por Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			//	this.cmbResponsable.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec VerCentroDecosto 2");
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec VerCentroDecosto 1");
			this.txtAnio.MinValue = 2024;
			this.txtAnio.MaxValue = DateTime.Today.Year + 1;
			this.txtAnio.Value = DateTime.Today.Year;
		}

		private void frmPresupuestoPorCC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidPresupuestoPorCC.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);

					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "2605");

						this.Consultar((int)this.txtidPresupuestoPorCC.Value);					
					}
				}

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

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region Bloqueo Controles 
			this.txtAnio.Enabled = false;
			this.cmbCentroDeCosto.Enabled= false;
			this.cmbResponsable.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.uGridCuentas);
			FuncionesProcedimientos.EstadoGrids(false, this.ugrValores);
			this.ugrValores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePresupuestoCC 0"));
			this.uGridCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaCuentasPresupuesto 3"));

			#endregion Bloqueo Controles 

			#region Vaciar Campos
			this.txtidPresupuestoPorCC.Value = 0;
			this.txtEstado.Value = 0;
			//this.txtAnio.Value = 0;
			this.cmbResponsable.Value = System.DBNull.Value;
			this.cmbCentroDeCosto.Value = System.DBNull.Value;	
			this.lblEstado.Text ="";
			#endregion Vaciar Campos
    
			#region Bloquea Botones 
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnSubir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnAprobar.Enabled = false;
			#endregion Bloquea Botones

		}

		private void ultraDataSource4_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			/* 1: Pendiente */
			/* 2: Aprobado */
			/* 3: Anulado */
			this.btnCancelar_Click(sender, e);
			this.txtAnio.Value = DateTime.Today.Year;


			#region Habilita Controles
			this.txtAnio.Enabled = true;
			this.cmbCentroDeCosto.Enabled = true;
			this.cmbResponsable.Enabled = true;
			this.txtEstado.Value = 1;

			FuncionesProcedimientos.EstadoGrids(true, this.uGridCuentas);
			FuncionesProcedimientos.EstadoGrids(true, this.ugrValores);
						 
			
			#endregion Habilita Controles 

			#region Bloquea Botones 
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnSubir.Enabled = true;
			#endregion Bloquea Botones 

			bNuevo = true;
			bEdicion = true;

			this.txtAnio.Focus();


			// this.uGridCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaCuentasPresupuesto 1"));
			//	this.ugrValores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePresupuestoCC 0"));
	
		}

		private void ugrValores_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ENERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "FEBRERO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MARZO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "ABRIL", 2);
			FuncionesProcedimientos.FormatoGrid(e, "MAYO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JUNIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "JULIO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "AGOSTO", 2);
			FuncionesProcedimientos.FormatoGrid(e, "SEPTIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "OCTUBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "NOVIEMBRE", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DICIEMBRE", 2);	
		}

		private void ugrValores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugrValores);
		
		}
		private void Sumatorias(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			Decimal dEnero = 0.00m;
			Decimal dFebrero= 0.00m;
			Decimal dMarzo = 0.00m;
			Decimal dAbril = 0.00m;
			Decimal dMayo = 0.00m;
			Decimal dJunio = 0.00m;
			Decimal dJulio = 0.00m;
			Decimal dAgosto = 0.00m;
			Decimal dSeptiembre = 0.00m;
			Decimal dOctubre = 0.00m;
			Decimal dNoviembre = 0.00m;
			Decimal dDiciembre = 0.00m;

			if ((e.Cell.Row.Cells["ENERO"].Value) != System.DBNull.Value ) dEnero = Convert.ToDecimal(e.Cell.Row.Cells["ENERO"].Value);
			if ((e.Cell.Row.Cells["FEBRERO"].Value) != System.DBNull.Value ) dFebrero = Convert.ToDecimal(e.Cell.Row.Cells["FEBRERO"].Value);
			if ((e.Cell.Row.Cells["MARZO"].Value) != System.DBNull.Value ) dMarzo = Convert.ToDecimal(e.Cell.Row.Cells["MARZO"].Value);
			if ((e.Cell.Row.Cells["ABRIL"].Value) != System.DBNull.Value ) dAbril = Convert.ToDecimal(e.Cell.Row.Cells["ABRIL"].Value);
			if ((e.Cell.Row.Cells["MAYO"].Value) != System.DBNull.Value ) dMayo = Convert.ToDecimal(e.Cell.Row.Cells["MAYO"].Value);
			if ((e.Cell.Row.Cells["JUNIO"].Value) != System.DBNull.Value ) dJunio = Convert.ToDecimal(e.Cell.Row.Cells["JUNIO"].Value);
			if ((e.Cell.Row.Cells["JULIO"].Value) != System.DBNull.Value ) dJulio = Convert.ToDecimal(e.Cell.Row.Cells["JULIO"].Value);
			if ((e.Cell.Row.Cells["AGOSTO"].Value) != System.DBNull.Value ) dAgosto = Convert.ToDecimal(e.Cell.Row.Cells["AGOSTO"].Value);
			if ((e.Cell.Row.Cells["SEPTIEMBRE"].Value) != System.DBNull.Value ) dSeptiembre = Convert.ToDecimal(e.Cell.Row.Cells["SEPTIEMBRE"].Value);
			if ((e.Cell.Row.Cells["OCTUBRE"].Value) != System.DBNull.Value ) dOctubre = Convert.ToDecimal(e.Cell.Row.Cells["OCTUBRE"].Value);
			if ((e.Cell.Row.Cells["NOVIEMBRE"].Value) != System.DBNull.Value ) dNoviembre = Convert.ToDecimal(e.Cell.Row.Cells["NOVIEMBRE"].Value);
			if ((e.Cell.Row.Cells["DICIEMBRE"].Value) != System.DBNull.Value ) dDiciembre = Convert.ToDecimal(e.Cell.Row.Cells["DICIEMBRE"].Value);


			e.Cell.Row.Cells["TOTAL"].Value = 
				dEnero +
				dFebrero +
				dMarzo +
				dAbril +
				dMayo +
				dJunio +
				dJulio +
				dAgosto +
				dSeptiembre +
				dOctubre +
				dNoviembre +
				dDiciembre ;			
		}
		bool bActualiza = false;
		bool bEditarMeses = false;

		private void ValidarPeriodo (string Periodo, string Mes)
		{
			if (Funcion.ValidaPeriodo(Convert.ToDateTime (Periodo), cdsSeteoF, "Presupuesto"))
			{
				DateTime dtFechaVP = Convert.ToDateTime (Periodo);
				MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbCentroDeCosto.Focus();			
				//	e.Cell.Row.Cells[Mes].Value = (decimal)e.Cell.OriginalValue;
				bActualiza = false;
				return;
			}
		}
		private void ugrValores_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEditarMeses )return;
			if (bActualiza) return;
			#region Actualiza Valores
			string sPeriodoFecha ="";
			#region validanulos
			if (e.Cell.Column.ToString() == "ENERO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["ENERO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["ENERO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["FEBRERO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["FEBRERO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "MARZO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["MARZO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["MARZO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "ABRIL")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["ABRIL"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["ABRIL"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "MAYO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["MAYO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["MAYO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "JUNIO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["JUNIO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["JUNIO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "JULIO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["JULIO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["JULIO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["AGOSTO"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["AGOSTO"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["SEPTIEMBRE"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["SEPTIEMBRE"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["OCTUBRE"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["OCTUBRE"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["NOVIEMBRE"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["NOVIEMBRE"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				bActualiza = true;
				if (e.Cell.Row.Cells["DICIEMBRE"].Value == System.DBNull.Value)
				{
					e.Cell.Row.Cells["DICIEMBRE"].Value = (decimal)e.Cell.OriginalValue;
					bActualiza = false;
					return;
				}
				bActualiza = false;	
			}
			


			#endregion validanulos
			if (e.Cell.Column.ToString() == "ENERO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-01-01";
				// ValidarPeriodo (sPeriodoFecha, "ENERO");
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["ENERO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["ENERO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;	
			}

			if (e.Cell.Column.ToString() == "FEBRERO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-02-01";
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["FEBRERO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["FEBRERO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;	
			}

			if (e.Cell.Column.ToString() == "MARZO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-03-01";
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["MARZO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["MARZO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;	
			}

			if (e.Cell.Column.ToString() == "ABRIL")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-04-01";
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["ABRIL"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["ABRIL"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;		
			}

			if (e.Cell.Column.ToString() == "MAYO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-05-01";
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["MAYO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["MAYO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "JUNIO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-06-01";
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["JUNIO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["JUNIO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "JULIO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-07-01";
				#region Valida Fecha 
				
				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["JULIO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["JULIO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "AGOSTO")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-08-01";
				#region Valida Fecha 

				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["AGOSTO"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["AGOSTO"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "SEPTIEMBRE")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-09-01";
				#region Valida Fecha

				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["SEPTIEMBRE"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["SEPTIEMBRE"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "OCTUBRE")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-10-01";
				#region Valida Fecha 

				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["OCTUBRE"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["OCTUBRE"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "NOVIEMBRE")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-11-01";
				#region Valida Fecha 

				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["NOVIEMBRE"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["NOVIEMBRE"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "DICIEMBRE")
			{
				bActualiza = true;
				sPeriodoFecha =this.txtAnio.Value +""+"-12-01";
				#region Valida Fecha 

				if (Funcion.ValidaPeriodo(Convert.ToDateTime (sPeriodoFecha), cdsSeteoF, "Presupuesto"))
				{
					if(!miAcceso.CambiarPrecio)
					{
						DateTime dtFechaVP = Convert.ToDateTime (sPeriodoFecha);
						MessageBox.Show(string.Format(" El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", dtFechaVP.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.cmbCentroDeCosto.Focus();			
						e.Cell.Row.Cells["DICIEMBRE"].Value = (decimal)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}
				}
				#endregion Valida Fecha 
				if (e.Cell.Row.Cells["DICIEMBRE"].Value != System.DBNull.Value) Sumatorias(e);
				bActualiza = false;
			}
			#endregion Actualiza Valores
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
					
			this.btnCancelar_Click(sender, e);

			using (frmBuscaPresupuestoCC consultab = new frmBuscaPresupuestoCC())
			{
				if (DialogResult.OK == consultab.ShowDialog())
				{
					this.Consultar((int)consultab.uGrid.ActiveRow.Cells["idPresupuestoPorCC"].Value);
				}				
			}		
		}
		private void Consultar (int idPresupuestoPorCC )
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec [ConsultaIndividualPresupuestoCC] {0}",idPresupuestoPorCC); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidPresupuestoPorCC.Value= dr.GetInt32 (0);
				if (dr.GetValue(1) != System.DBNull.Value)this.txtAnio.Value = dr.GetInt32 (1);
				if (dr.GetValue(2) != System.DBNull.Value)this.cmbCentroDeCosto.Value = dr.GetInt32 (2);
				//if (dr.GetValue(3) != System.DBNull.Value)this.cmbResponsable.Value = dr.GetInt32 (3);
				if (dr.GetValue(4) != System.DBNull.Value)this.txtEstado.Value = dr.GetInt32 (4);
				if (dr.GetValue(5) != System.DBNull.Value)this.lblEstado.Text = dr.GetString (5);
			}
			dr.Close();
			#endregion Consulta Individual

			string sSQLPa = string.Format(" EXEC ConsultaDetallePresupuestoCC   {0} ", idPresupuestoPorCC);
			FuncionesProcedimientos.dsGeneral(sSQLPa, this.ugrValores);


			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnSubir.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ugrValores);
			FuncionesProcedimientos.EstadoGrids(false, this.uGridCuentas);

			#region sumatoria
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drA in this.ugrValores.Rows.All)
			{
				drA.Cells["TOTAL"].Value = 0;
        
				Decimal dEnero = 0.00m;
				Decimal dFebrero= 0.00m;
				Decimal dMarzo = 0.00m;
				Decimal dAbril = 0.00m;
				Decimal dMayo = 0.00m;
				Decimal dJunio = 0.00m;
				Decimal dJulio = 0.00m;
				Decimal dAgosto = 0.00m;
				Decimal dSeptiembre = 0.00m;
				Decimal dOctubre = 0.00m;
				Decimal dNoviembre = 0.00m;
				Decimal dDiciembre = 0.00m;

				if ((drA.Cells["ENERO"].Value) != System.DBNull.Value ) dEnero = Convert.ToDecimal(drA.Cells["ENERO"].Value);
				if ((drA.Cells["FEBRERO"].Value) != System.DBNull.Value ) dFebrero = Convert.ToDecimal(drA.Cells["FEBRERO"].Value);
				if ((drA.Cells["MARZO"].Value) != System.DBNull.Value ) dMarzo = Convert.ToDecimal(drA.Cells["MARZO"].Value);
				if ((drA.Cells["ABRIL"].Value) != System.DBNull.Value ) dAbril = Convert.ToDecimal(drA.Cells["ABRIL"].Value);
				if ((drA.Cells["MAYO"].Value) != System.DBNull.Value ) dMayo = Convert.ToDecimal(drA.Cells["MAYO"].Value);
				if ((drA.Cells["JUNIO"].Value) != System.DBNull.Value ) dJunio = Convert.ToDecimal(drA.Cells["JUNIO"].Value);
				if ((drA.Cells["JULIO"].Value) != System.DBNull.Value ) dJulio = Convert.ToDecimal(drA.Cells["JULIO"].Value);
				if ((drA.Cells["AGOSTO"].Value) != System.DBNull.Value ) dAgosto = Convert.ToDecimal(drA.Cells["AGOSTO"].Value);
				if ((drA.Cells["SEPTIEMBRE"].Value) != System.DBNull.Value ) dSeptiembre = Convert.ToDecimal(drA.Cells["SEPTIEMBRE"].Value);
				if ((drA.Cells["OCTUBRE"].Value) != System.DBNull.Value ) dOctubre = Convert.ToDecimal(drA.Cells["OCTUBRE"].Value);
				if ((drA.Cells["NOVIEMBRE"].Value) != System.DBNull.Value ) dNoviembre = Convert.ToDecimal(drA.Cells["NOVIEMBRE"].Value);
				if ((drA.Cells["DICIEMBRE"].Value) != System.DBNull.Value ) dDiciembre = Convert.ToDecimal(drA.Cells["DICIEMBRE"].Value);


				drA.Cells["TOTAL"].Value = 	( dEnero +
					dFebrero+
					dMarzo+
					dAbril +
					dMayo +
					dJunio +
					dJulio +
					dAgosto +
					dSeptiembre+ 
					dOctubre +
					dNoviembre+ 
					dDiciembre );

			}
			#endregion sumatoria

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar)
				if ((int)this.txtEstado.Value == 0 || (int)this.txtEstado.Value == 1) 
					this.btnEditar.Enabled = true;

			if (miAcceso.Anular)
				if ((int)this.txtEstado.Value <= 2) 
					this.btnAnular.Enabled = true;

			if (miAcceso.BAutorizar)
				if ((int)this.txtEstado.Value <= 1) 
					this.btnAprobar.Enabled = true;
			

		}


		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consultar ((int)this.txtidPresupuestoPorCC.Value);
			
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

			if ((int)this.txtEstado.Value <= 1)
			{
				//this.txtAnio.Enabled = true;
				//this.cmbCentroDeCosto.Enabled = true;
				this.cmbResponsable.Enabled = true;
				FuncionesProcedimientos.EstadoGrids(true, this.ugrValores);
				FuncionesProcedimientos.EstadoGrids(true, this.uGridCuentas);
				this.uGridCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaCuentasPresupuesto 1, {0}, {1}", (int) this.txtAnio.Value, (int) this.cmbCentroDeCosto.Value));
			}
			

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugrValores.Rows.All)
			{
				int icuentas = 0;
				icuentas = Convert.ToInt32 ( dr.Cells["idCuenta"].Value);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridCuentas.Rows.All)
				{
					if ( icuentas == Convert.ToInt32 ( drr.Cells["idCuenta"].Value))
					{
						drr.Cells["Sel"].Value = true;
					}
				}
			}

		

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
			this.btnSubir.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnAprobar.Enabled = false;


			bNuevo = false;
			bEdicion = true;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2017, DateTime.Today.Year)) return;
			if (!Validacion.vbComboVacio(this.cmbCentroDeCosto, "Seleccione el Centro de Costo")) return;
//			if (!Validacion.vbComboVacio(this.cmbResponsable, "Seleccione la persona Responsable")) return;
			int icontProeyecto = 0;
			icontProeyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaCentroDeCostoYResponsable {0}, {1}", (int) this.cmbCentroDeCosto.Value, (int) this.txtAnio.Value));	
			if ( bNuevo)
			{
				if (icontProeyecto > 0)
				{
					MessageBox.Show("Ya existe un presupuesto creado para el \n Centro de Costo seleccionado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ugrValores.Focus ();
					return;
				}
			}
			else
			{
				if (icontProeyecto > 1)
				{
					MessageBox.Show("Ya existe un presupuesto creado para el \n Centro de Costo seleccionado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ugrValores.Focus ();
					return;
				}
			}

			if ((int) this.ugrValores.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese al menos una Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ugrValores.Focus ();
				return;
			}
			Decimal dTotales = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow draC in this.ugrValores.Rows.All)
			{

				int iContPres = 0;
				iContPres = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCentroDeCostoExistente] {0}, {1}, {2}", Convert.ToInt32 ( draC.Cells["idCuenta"].Value), (int) this.txtAnio.Value, (int) this.cmbCentroDeCosto.Value));	
				if (iContPres == 0)
				{
					MessageBox.Show("La Cuenta no esta asignado al centro de costo seleccionado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ugrValores.ActiveRow = draC;
					this.ugrValores.ActiveRow.Selected = true;
					return;
				}
				
			}


			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.ugrValores.Rows.All)
			{

				dTotales = ((decimal)dra.Cells["TOTAL"].Value + dTotales );
				
			}

			if (dTotales == 0)
			{
				MessageBox.Show("El monto Total debe ser mayor a 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ugrValores.Focus ();
				return;
			}
			if (dTotales < 0)
			{
				MessageBox.Show("El monto Total no purse ser menor a 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.ugrValores.Focus ();
				return;
			}

			#region Guardar
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

					

					#region Graba Maestro
					int idProyecto = 0;
					int idNomina = 0;
					if (this.cmbCentroDeCosto.ActiveRow != null) idProyecto = (int) this.cmbCentroDeCosto.Value;
					if(this.cmbResponsable.ActiveRow != null) idNomina = (int) this.cmbResponsable.Value;

					string  sSQL1 = string.Format("Exec [GuardaPresupuestoPorCC] {0}, {1}, {2}, {3}, {4}" ,
						(int)this.txtidPresupuestoPorCC.Value, 
						(int)this.txtAnio.Value, 
						idProyecto,
						idNomina,
						(int)this.txtEstado.Value
						);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidPresupuestoPorCC.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					#region Graba Ingreso Gastos
					if((int) this.ugrValores.Rows.Count > 0)
					{						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugrValores.Rows.All)
						{
							Decimal dEnero = 0.00m;
							Decimal dFebrero= 0.00m;
							Decimal dMarzo = 0.00m;
							Decimal dAbril = 0.00m;
							Decimal dMayo = 0.00m;
							Decimal dJunio = 0.00m;
							Decimal dJulio = 0.00m;
							Decimal dAgosto = 0.00m;
							Decimal dSeptiembre = 0.00m;
							Decimal dOctubre = 0.00m;
							Decimal dNoviembre = 0.00m;
							Decimal dDiciembre = 0.00m;

							if ( dr.Cells["ENERO"].Value != System.DBNull.Value )  dEnero = Convert.ToDecimal(dr.Cells["ENERO"].Value);
							if ( dr.Cells["FEBRERO"].Value != System.DBNull.Value )  dFebrero = Convert.ToDecimal(dr.Cells["FEBRERO"].Value);
							if ( dr.Cells["MARZO"].Value != System.DBNull.Value )  dMarzo = Convert.ToDecimal(dr.Cells["MARZO"].Value);
							if ( dr.Cells["ABRIL"].Value != System.DBNull.Value )  dAbril = Convert.ToDecimal(dr.Cells["ABRIL"].Value);
							if ( dr.Cells["MAYO"].Value != System.DBNull.Value )  dMayo = Convert.ToDecimal(dr.Cells["MAYO"].Value);
							if ( dr.Cells["JUNIO"].Value != System.DBNull.Value )  dJunio = Convert.ToDecimal(dr.Cells["JUNIO"].Value);
							if ( dr.Cells["JULIO"].Value != System.DBNull.Value )  dJulio = Convert.ToDecimal(dr.Cells["JULIO"].Value);
							if ( dr.Cells["AGOSTO"].Value != System.DBNull.Value )  dAgosto = Convert.ToDecimal(dr.Cells["AGOSTO"].Value);
							if ( dr.Cells["SEPTIEMBRE"].Value != System.DBNull.Value )  dSeptiembre = Convert.ToDecimal(dr.Cells["SEPTIEMBRE"].Value);
							if ( dr.Cells["OCTUBRE"].Value != System.DBNull.Value )  dOctubre = Convert.ToDecimal(dr.Cells["OCTUBRE"].Value);
							if ( dr.Cells["NOVIEMBRE"].Value != System.DBNull.Value )  dNoviembre = Convert.ToDecimal(dr.Cells["NOVIEMBRE"].Value);
							if ( dr.Cells["DICIEMBRE"].Value != System.DBNull.Value )  dDiciembre  = Convert.ToDecimal(dr.Cells["DICIEMBRE"].Value);


							string  sSQLDeta = string.Format("Exec GuardaDetallePresupuestoPorCC {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8},{9},{10},{11},{12},{13}" ,
								(int) this.txtidPresupuestoPorCC.Value, 
								Convert.ToInt32 ( dr.Cells["idCuenta"].Value), 
								dEnero ,
								dFebrero ,
								dMarzo ,
								dAbril ,
								dMayo ,
								dJunio ,
								dJulio ,
								dAgosto ,
								dSeptiembre ,
								dOctubre ,
								dNoviembre ,
								dDiciembre);

							oCmdActualiza.CommandText = sSQLDeta;
							oCmdActualiza.ExecuteNonQuery();	
						}						
					}
					#endregion Graba Ingreso Gastos

					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
					
					this.txtAnio.Enabled = false;
					this.cmbCentroDeCosto.Enabled=false;
					this.cmbResponsable.Enabled=false;

					FuncionesProcedimientos.EstadoGrids(false, this.uGridCuentas);
					FuncionesProcedimientos.EstadoGrids(false, this.ugrValores);

					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;
					this.btnSubir.Enabled = false;
			
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

		private void uGridCuentas_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idCuenta"].Value != System.DBNull.Value)
			{
				this.uGridCuentas.UpdateData();

				bool bDisponible = false;		
				if ((bool)e.Cell.Row.Cells["Sel"].Value) bDisponible = true;
				else bDisponible = false;
				if (bDisponible)
				{
					bEditarMeses = true;
					#region Ingresa valores
					this.ugrValores.Rows.Band.AddNew();
					//	this.ugrValores.ActiveRow.Cells["idDetallePresupuestoPorCC"].Value = 0;
					this.ugrValores.ActiveRow.Cells["idPresupuestoPorCC"].Value = (int) this.txtidPresupuestoPorCC.Value;
					this.ugrValores.ActiveRow.Cells["idCuenta"].Value = (int) e.Cell.Row.Cells["idCuenta"].Value;
					this.ugrValores.ActiveRow.Cells["Codigo"].Value = (string) e.Cell.Row.Cells["Codigo"].Value;
					this.ugrValores.ActiveRow.Cells["Cuenta"].Value = (string) e.Cell.Row.Cells["Cuenta"].Value;
					this.ugrValores.ActiveRow.Cells["ENERO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["FEBRERO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["MARZO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["ABRIL"].Value = 0;
					this.ugrValores.ActiveRow.Cells["MAYO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["JUNIO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["JULIO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["AGOSTO"].Value = 0;
					this.ugrValores.ActiveRow.Cells["SEPTIEMBRE"].Value = 0;
					this.ugrValores.ActiveRow.Cells["OCTUBRE"].Value = 0;
					this.ugrValores.ActiveRow.Cells["NOVIEMBRE"].Value = 0;
					this.ugrValores.ActiveRow.Cells["DICIEMBRE"].Value = 0;
					this.ugrValores.ActiveRow.Cells["TOTAL"].Value = 0;
					this.ugrValores.ActiveRow.Cells["Dupli"].Value = 0;
					bEditarMeses = false;
					#endregion Ingresa valores
				}
				else
				{
					#region Valida cuenta Existe
					int idCuenta = 0;
					int icont = 0;
					idCuenta = (int)e.Cell.Row.Cells["idCuenta"].Value;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ugrValores.Rows.All)
					{
						if ((int)drr.Cells["idCuenta"].Value == idCuenta)
						{
							if ((int)drr.Cells["idPresupuestoPorCC"].Value > 0 )
							{
								icont = icont +1;
							}
						}
					}
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ugrValores.Rows.All)
					{
						if ((int)drr.Cells["idCuenta"].Value == idCuenta)
						{
							icont = icont +1;
						}
					}
					#endregion Valida cuenta Existe
					if (icont> 0)
					{
						#region Elimina
						if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						{		
							bEditarMeses = true;
							EliminaCuenta (this.ugrValores, idCuenta)	;		
							bEditarMeses = true;
						}
						else
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ugrValores.Rows.All)
							{
								if ((int)drr.Cells["idCuenta"].Value == idCuenta)
								{
									e.Cell.Row.Cells["Sel"].Value = true;
								}
							}

						}
						#endregion Elimina
					}
				}
			}
		}

		private void EliminaCuenta(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idCuenta )
		{
			#region Promocion

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if ((int)dr.Cells["idCuenta"].Value == idCuenta && (int)dr.Cells["idPresupuestoPorCC"].Value == 0)
				{
					dr.Delete();
					return;
				}
				if ((int)dr.Cells["idCuenta"].Value == idCuenta && (int)dr.Cells["idPresupuestoPorCC"].Value > 0)
				{
					string sSQL = string.Format("Delete from  DetallePresupuestoPorCC where idPresupuestoPorCC = {0} and idcuenta =  {1} ", 
						(int)dr.Cells["idPresupuestoPorCC"].Value, (int)dr.Cells["idCuenta"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
					dr.Delete();
					return;
				}
			}
			#endregion Promocion


		}


		private void cmbCentroDeCosto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close ();
		}

		private void ugrValores_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEditarMeses) 
			{
				e.Cancel = true;
				//return;
			}
			else
			{
				e.DisplayPromptMsg = false;
			}
				
			
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show(string.Format("Esta Seguro de Aprobar el presupuesto. \n\n ¿ Desea Continuar ?"), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE PresupuestoPorCC SET Estado = 2, UsuarioAprueba = '{0}', FechaAprueba = Getdate ()  WHERE idPresupuestoPorCC = {1} ", 
					MenuLatinium.stUsuario,  (int) this.txtidPresupuestoPorCC.Value));

				Consultar ((int) this.txtidPresupuestoPorCC.Value);
			}
		}

		private void btnSubir_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbCentroDeCosto, "Seleccione el Centro de Costo")) return;
			#region Excel
			OpenFileDialog sArchivo = new OpenFileDialog();

			sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";
			string sCuentasNot ="";
			string sCuentaseXIS ="";
			if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
			{
				this.Cursor = Cursors.WaitCursor;                                  

				int iCont = 0;
				
				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CODIGO,	ENERO, FEBRERO, MARZO, ABRIL,	MAYO,	JUNIO, JULIO,	AGOSTO,	SEPTIEMBRE,	OCTUBRE,	NOVIEMBRE,	DICIEMBRE FROM [PRESUPUESTO$] ", "PRESUPUESTO").Tables[0].Rows)
				{
					#region Variables
					string CODIGO = "";
					decimal ENERO = 0.00m;
					decimal	FEBRERO = 0.00m; 
					decimal	MARZO = 0.00m;
					decimal	ABRIL = 0.00m;	
					decimal	MAYO = 0.00m;	
					decimal	JUNIO = 0.00m;
					decimal	JULIO = 0.00m;	
					decimal	AGOSTO = 0.00m;
					decimal	SEPTIEMBRE = 0.00m;	
					decimal	OCTUBRE = 0.00m;	
					decimal	NOVIEMBRE = 0.00m;	
					decimal	DICIEMBRE = 0.00m;

					if (row["CODIGO"] != System.DBNull.Value) CODIGO = row["CODIGO"].ToString();
					if (row["ENERO"] != System.DBNull.Value) ENERO = Convert.ToDecimal( row["ENERO"].ToString());
					if (row["FEBRERO"] != System.DBNull.Value) FEBRERO = Convert.ToDecimal( row["FEBRERO"].ToString());
					if (row["MARZO"] != System.DBNull.Value) MARZO = Convert.ToDecimal( row["MARZO"].ToString());
					if (row["ABRIL"] != System.DBNull.Value) ABRIL = Convert.ToDecimal( row["ABRIL"].ToString());
					if (row["MAYO"] != System.DBNull.Value) MAYO = Convert.ToDecimal( row["MAYO"].ToString());
					if (row["JUNIO"] != System.DBNull.Value) JUNIO = Convert.ToDecimal( row["JUNIO"].ToString());
					if (row["JULIO"] != System.DBNull.Value) JULIO = Convert.ToDecimal( row["JULIO"].ToString());
					if (row["AGOSTO"] != System.DBNull.Value) AGOSTO = Convert.ToDecimal( row["AGOSTO"].ToString());
					if (row["SEPTIEMBRE"] != System.DBNull.Value) SEPTIEMBRE = Convert.ToDecimal( row["SEPTIEMBRE"].ToString());
					if (row["OCTUBRE"] != System.DBNull.Value) OCTUBRE = Convert.ToDecimal( row["OCTUBRE"].ToString());
					if (row["NOVIEMBRE"] != System.DBNull.Value) NOVIEMBRE = Convert.ToDecimal( row["NOVIEMBRE"].ToString());
					if (row["DICIEMBRE"] != System.DBNull.Value) DICIEMBRE = Convert.ToDecimal( row["DICIEMBRE"].ToString());
					#endregion variables;
					int iContcuenta = 0;
					int iContcuentaeX = 0;
					#region Valida cuenat existente
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)
					{
						
						if ( CODIGO == Convert.ToString ( dr.Cells["Codigo"].Value))
						{
							iContcuenta ++;
						}

					}
					if (iContcuenta == 0) sCuentasNot = sCuentasNot +" "+ CODIGO+"," ;
					#endregion Valida Cuentas ya creadas

					#region Valida cuenat existente
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.ugrValores.Rows.All)
					{
						
						if ( CODIGO == Convert.ToString ( drr.Cells["Codigo"].Value) && Convert.ToDecimal ( drr.Cells["TOTAL"].Value) > 0 )
						{
							iContcuentaeX ++;
						}

					}
					if (iContcuentaeX > 0) sCuentaseXIS = sCuentaseXIS + " "+ CODIGO+"," ;
					#endregion Valida Cuentas ya creadas
				}	
				this.Cursor = Cursors.Default;

				if (sCuentasNot.Length > 0 )
				{
					if (DialogResult.No ==	MessageBox.Show(string.Format("Cuentas:, '{0}' no Existen o no estan asignadas \n\n ¿ Desea Continuar ?", sCuentasNot), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						return;
					}
				}
				if (sCuentaseXIS.Length > 0 )
				{
					if (DialogResult.No ==	MessageBox.Show(string.Format("Cuentas:, '{0}' , Tienen valores registrados.\n\n ¿ Desea Sobreescribir ?", sCuentaseXIS), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						return;
					}
				}


				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CODIGO,	ENERO, FEBRERO, MARZO, ABRIL,	MAYO,	JUNIO, JULIO,	AGOSTO,	SEPTIEMBRE,	OCTUBRE,	NOVIEMBRE,	DICIEMBRE FROM [PRESUPUESTO$] ", "PRESUPUESTO").Tables[0].Rows)
				{
					#region Variables
					string CODIGO = "";
					decimal ENERO = 0.00m;
					decimal	FEBRERO = 0.00m; 
					decimal	MARZO = 0.00m;
					decimal	ABRIL = 0.00m;	
					decimal	MAYO = 0.00m;	
					decimal	JUNIO = 0.00m;
					decimal	JULIO = 0.00m;	
					decimal	AGOSTO = 0.00m;
					decimal	SEPTIEMBRE = 0.00m;	
					decimal	OCTUBRE = 0.00m;	
					decimal	NOVIEMBRE = 0.00m;	
					decimal	DICIEMBRE = 0.00m;

	
					bool  bENERO = false;
					bool	bFEBRERO = false;
					bool	bMARZO = false;
					bool	bABRIL = false;	
					bool	bMAYO = false;	
					bool	bJUNIO = false;
					bool	bJULIO = false;	
					bool	bAGOSTO = false;
					bool	bSEPTIEMBRE = false;
					bool	bOCTUBRE = false;
					bool	bNOVIEMBRE = false;
					bool	bDICIEMBRE = false;

					bENERO = ValidaPeriodo (  "ENERO", (int) this.txtAnio.Value );
					bFEBRERO = ValidaPeriodo (  "FEBRERO", (int) this.txtAnio.Value );
					bMARZO = ValidaPeriodo (  "MARZO", (int) this.txtAnio.Value );
					bABRIL = ValidaPeriodo (  "ABRIL", (int) this.txtAnio.Value );
					bMAYO = ValidaPeriodo (  "MAYO", (int) this.txtAnio.Value );
					bJUNIO = ValidaPeriodo (  "JUNIO", (int) this.txtAnio.Value );
					bJULIO = ValidaPeriodo (  "JULIO", (int) this.txtAnio.Value );
					bAGOSTO = ValidaPeriodo (  "AGOSTO", (int) this.txtAnio.Value );
					bSEPTIEMBRE = ValidaPeriodo (  "SEPTIEMBRE", (int) this.txtAnio.Value );
					bOCTUBRE = ValidaPeriodo (  "OCTUBRE", (int) this.txtAnio.Value );
					bNOVIEMBRE = ValidaPeriodo (  "NOVIEMBRE", (int) this.txtAnio.Value );
					bDICIEMBRE = ValidaPeriodo (  "DICIEMBRE", (int) this.txtAnio.Value );
	

					if (row["CODIGO"] != System.DBNull.Value) CODIGO = row["CODIGO"].ToString();
					if (row["ENERO"] != System.DBNull.Value) ENERO = Convert.ToDecimal( row["ENERO"].ToString());
					if (row["FEBRERO"] != System.DBNull.Value) FEBRERO = Convert.ToDecimal( row["FEBRERO"].ToString());
					if (row["MARZO"] != System.DBNull.Value) MARZO = Convert.ToDecimal( row["MARZO"].ToString());
					if (row["ABRIL"] != System.DBNull.Value) ABRIL = Convert.ToDecimal( row["ABRIL"].ToString());
					if (row["MAYO"] != System.DBNull.Value) MAYO = Convert.ToDecimal( row["MAYO"].ToString());
					if (row["JUNIO"] != System.DBNull.Value) JUNIO = Convert.ToDecimal( row["JUNIO"].ToString());
					if (row["JULIO"] != System.DBNull.Value) JULIO = Convert.ToDecimal( row["JULIO"].ToString());
					if (row["AGOSTO"] != System.DBNull.Value) AGOSTO = Convert.ToDecimal( row["AGOSTO"].ToString());
					if (row["SEPTIEMBRE"] != System.DBNull.Value) SEPTIEMBRE = Convert.ToDecimal( row["SEPTIEMBRE"].ToString());
					if (row["OCTUBRE"] != System.DBNull.Value) OCTUBRE = Convert.ToDecimal( row["OCTUBRE"].ToString());
					if (row["NOVIEMBRE"] != System.DBNull.Value) NOVIEMBRE = Convert.ToDecimal( row["NOVIEMBRE"].ToString());
					if (row["DICIEMBRE"] != System.DBNull.Value) DICIEMBRE = Convert.ToDecimal( row["DICIEMBRE"].ToString());
					#endregion variables;

					#region Valores dond existen cuena y tienen valores ya ingresados
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)
					{
						int icontador = 0;
						int iContcuenta = 0;
						if ( CODIGO == Convert.ToString ( dr.Cells["Codigo"].Value))
						{
							iContcuenta ++;
						}
						if (iContcuenta > 0 )/*El codigo si existe en el grid de cuentas*/
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drV in this.ugrValores.Rows.All)
							{
								int iContcuentaEx = 0;
								if ( CODIGO == Convert.ToString ( drV.Cells["Codigo"].Value)  )
								{
									iContcuentaEx ++ ;
									
								}
								if (iContcuentaEx > 0 )
								{
									//	if (Convert.ToDecimal ( drV.Cells["TOTAL"].Value) > 0 )
									//	{
									bEditarMeses = true;
									#region Ingresa valores Grid
									//	this.ugrValores.Rows.Band.AddNew();
									//this.ugrValores.ActiveRow.Cells["idPresupuestoPorCC"].Value = 0;
									//this.ugrValores.ActiveRow.Cells["idCuenta"].Value = 0;
									drV.Cells["Codigo"].Value = CODIGO;
									//drV.Cells["Cuenta"].Value = "Actualiza";
									if (!bENERO)drV.Cells["ENERO"].Value =  ENERO;
									if (!bFEBRERO)drV.Cells["FEBRERO"].Value = FEBRERO;
									if (!bMARZO)drV.Cells["MARZO"].Value = MARZO;
									if (!bABRIL)drV.Cells["ABRIL"].Value = ABRIL;
									if (!bMAYO)drV.Cells["MAYO"].Value = MAYO;
									if (!bJUNIO)drV.Cells["JUNIO"].Value = JUNIO;
									if (!bJULIO)drV.Cells["JULIO"].Value = JULIO;
									if (!bAGOSTO)drV.Cells["AGOSTO"].Value = AGOSTO;
									if (!bSEPTIEMBRE)drV.Cells["SEPTIEMBRE"].Value = SEPTIEMBRE;
									if (!bOCTUBRE)drV.Cells["OCTUBRE"].Value  = OCTUBRE;
									if (!bNOVIEMBRE)drV.Cells["NOVIEMBRE"].Value = NOVIEMBRE;
									if (!bDICIEMBRE)drV.Cells["DICIEMBRE"].Value  = DICIEMBRE;
									drV.Cells["TOTAL"].Value = 0;
									drV.Cells["Dupli"].Value = 0;
									/*
										this.ugrValores.ActiveRow.Cells["Codigo"].Value = CODIGO;
										this.ugrValores.ActiveRow.Cells["Cuenta"].Value = "Actualiza";
										this.ugrValores.ActiveRow.Cells["ENERO"].Value = ENERO;
										this.ugrValores.ActiveRow.Cells["FEBRERO"].Value = FEBRERO;
										this.ugrValores.ActiveRow.Cells["MARZO"].Value = MARZO;
										this.ugrValores.ActiveRow.Cells["ABRIL"].Value = ABRIL;
										this.ugrValores.ActiveRow.Cells["MAYO"].Value = MAYO;
										this.ugrValores.ActiveRow.Cells["JUNIO"].Value = JUNIO;
										this.ugrValores.ActiveRow.Cells["JULIO"].Value = JULIO;
										this.ugrValores.ActiveRow.Cells["AGOSTO"].Value = AGOSTO;
										this.ugrValores.ActiveRow.Cells["SEPTIEMBRE"].Value = SEPTIEMBRE;
										this.ugrValores.ActiveRow.Cells["OCTUBRE"].Value = OCTUBRE;
										this.ugrValores.ActiveRow.Cells["NOVIEMBRE"].Value = NOVIEMBRE;
										this.ugrValores.ActiveRow.Cells["DICIEMBRE"].Value = DICIEMBRE;
										this.ugrValores.ActiveRow.Cells["TOTAL"].Value = 0;
										*/
									//	Sumatorias(e);
									#endregion Ingresa valores Grid
									bEditarMeses = false;
									//	}
								}
							}					
						}
					}
					#endregion Valores dond existen cuena y tienen valores ya ingresados

				}

				foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT CODIGO,	ENERO, FEBRERO, MARZO, ABRIL,	MAYO,	JUNIO, JULIO,	AGOSTO,	SEPTIEMBRE,	OCTUBRE,	NOVIEMBRE,	DICIEMBRE FROM [PRESUPUESTO$] ", "PRESUPUESTO").Tables[0].Rows)
				{
					#region Variables
					string CODIGO = "";
					decimal ENERO = 0.00m;
					decimal	FEBRERO = 0.00m; 
					decimal	MARZO = 0.00m;
					decimal	ABRIL = 0.00m;	
					decimal	MAYO = 0.00m;	
					decimal	JUNIO = 0.00m;
					decimal	JULIO = 0.00m;	
					decimal	AGOSTO = 0.00m;
					decimal	SEPTIEMBRE = 0.00m;	
					decimal	OCTUBRE = 0.00m;	
					decimal	NOVIEMBRE = 0.00m;	
					decimal	DICIEMBRE = 0.00m;

					bool  bENERO = false;
					bool	bFEBRERO = false;
					bool	bMARZO = false;
					bool	bABRIL = false;	
					bool	bMAYO = false;	
					bool	bJUNIO = false;
					bool	bJULIO = false;	
					bool	bAGOSTO = false;
					bool	bSEPTIEMBRE = false;
					bool	bOCTUBRE = false;
					bool	bNOVIEMBRE = false;
					bool	bDICIEMBRE = false;

					bENERO = ValidaPeriodo (  "ENERO", (int) this.txtAnio.Value );
					bFEBRERO = ValidaPeriodo (  "FEBRERO", (int) this.txtAnio.Value );
					bMARZO = ValidaPeriodo (  "MARZO", (int) this.txtAnio.Value );
					bABRIL = ValidaPeriodo (  "ABRIL", (int) this.txtAnio.Value );
					bMAYO = ValidaPeriodo (  "MAYO", (int) this.txtAnio.Value );
					bJUNIO = ValidaPeriodo (  "JUNIO", (int) this.txtAnio.Value );
					bJULIO = ValidaPeriodo (  "JULIO", (int) this.txtAnio.Value );
					bAGOSTO = ValidaPeriodo (  "AGOSTO", (int) this.txtAnio.Value );
					bSEPTIEMBRE = ValidaPeriodo (  "SEPTIEMBRE", (int) this.txtAnio.Value );
					bOCTUBRE = ValidaPeriodo (  "OCTUBRE", (int) this.txtAnio.Value );
					bNOVIEMBRE = ValidaPeriodo (  "NOVIEMBRE", (int) this.txtAnio.Value );
					bDICIEMBRE = ValidaPeriodo (  "DICIEMBRE", (int) this.txtAnio.Value );
          


					if (row["CODIGO"] != System.DBNull.Value) CODIGO = row["CODIGO"].ToString();
					if (!bENERO)if (row["ENERO"] != System.DBNull.Value) ENERO = Convert.ToDecimal( row["ENERO"].ToString());
					if (!bFEBRERO)if (row["FEBRERO"] != System.DBNull.Value) FEBRERO = Convert.ToDecimal( row["FEBRERO"].ToString());
					if (!bMARZO)if (row["MARZO"] != System.DBNull.Value) MARZO = Convert.ToDecimal( row["MARZO"].ToString());
					if (!bABRIL)if (row["ABRIL"] != System.DBNull.Value) ABRIL = Convert.ToDecimal( row["ABRIL"].ToString());
					if (!bMAYO)if (row["MAYO"] != System.DBNull.Value) MAYO = Convert.ToDecimal( row["MAYO"].ToString());
					if (!bJUNIO)if (row["JUNIO"] != System.DBNull.Value) JUNIO = Convert.ToDecimal( row["JUNIO"].ToString());
					if (!bJULIO)if (row["JULIO"] != System.DBNull.Value) JULIO = Convert.ToDecimal( row["JULIO"].ToString());
					if (!bAGOSTO)if (row["AGOSTO"] != System.DBNull.Value) AGOSTO = Convert.ToDecimal( row["AGOSTO"].ToString());
					if (!bSEPTIEMBRE)if (row["SEPTIEMBRE"] != System.DBNull.Value) SEPTIEMBRE = Convert.ToDecimal( row["SEPTIEMBRE"].ToString());
					if (!bOCTUBRE)if (row["OCTUBRE"] != System.DBNull.Value) OCTUBRE = Convert.ToDecimal( row["OCTUBRE"].ToString());
					if (!bNOVIEMBRE)if (row["NOVIEMBRE"] != System.DBNull.Value) NOVIEMBRE = Convert.ToDecimal( row["NOVIEMBRE"].ToString());
					if (!bDICIEMBRE)if (row["DICIEMBRE"] != System.DBNull.Value) DICIEMBRE = Convert.ToDecimal( row["DICIEMBRE"].ToString());
					#endregion variables;

					#region Valores dond existen cuena y tienen valores ya ingresados
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)
					{
						int icontador = 0;
						int iContcuenta = 0;
						if ( CODIGO == Convert.ToString ( dr.Cells["Codigo"].Value))
						{
							iContcuenta ++;
						}
						if (iContcuenta > 0 )/*El codigo si existe en el grid de cuentas*/
						{
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drV in this.ugrValores.Rows.All)
							{
								int iContcuentaEx = 0;
								//								int icuentas = 0;
								//								string sCuentas ="";
								//
								//								//icuentas = (int)drV.Cells["idCuenta"].Value;
								//								//sCuentas = (string)	drV.Cells["Cuenta"].Value;

								if ( CODIGO == Convert.ToString ( drV.Cells["Codigo"].Value)  )
								{
									iContcuentaEx ++ ;
									//break;
									//continue;
								}
								if (iContcuentaEx == 0 )
								{
									bEditarMeses = true;
									#region Ingresa valores Grid
									this.ugrValores.Rows.Band.AddNew();
									this.ugrValores.ActiveRow.Cells["idPresupuestoPorCC"].Value = 0;
									this.ugrValores.ActiveRow.Cells["idCuenta"].Value = 0;
									this.ugrValores.ActiveRow.Cells["Codigo"].Value = CODIGO;
									this.ugrValores.ActiveRow.Cells["Cuenta"].Value = "";
									this.ugrValores.ActiveRow.Cells["ENERO"].Value = ENERO;
									this.ugrValores.ActiveRow.Cells["FEBRERO"].Value = FEBRERO;
									this.ugrValores.ActiveRow.Cells["MARZO"].Value = MARZO;
									this.ugrValores.ActiveRow.Cells["ABRIL"].Value = ABRIL;
									this.ugrValores.ActiveRow.Cells["MAYO"].Value = MAYO;
									this.ugrValores.ActiveRow.Cells["JUNIO"].Value = JUNIO;
									this.ugrValores.ActiveRow.Cells["JULIO"].Value = JULIO;
									this.ugrValores.ActiveRow.Cells["AGOSTO"].Value = AGOSTO;
									this.ugrValores.ActiveRow.Cells["SEPTIEMBRE"].Value = SEPTIEMBRE;
									this.ugrValores.ActiveRow.Cells["OCTUBRE"].Value = OCTUBRE;
									this.ugrValores.ActiveRow.Cells["NOVIEMBRE"].Value = NOVIEMBRE;
									this.ugrValores.ActiveRow.Cells["DICIEMBRE"].Value = DICIEMBRE;
									this.ugrValores.ActiveRow.Cells["TOTAL"].Value = 0;
									this.ugrValores.ActiveRow.Cells["Dupli"].Value = 1;
									//Sumatorias(e);
									#endregion Ingresa valores Grid
									bEditarMeses = false;
									break;
								}
							}					
						}
					}
					#endregion Valores dond existen cuena y tienen valores ya ingresados
				}	
			}
			#endregion Excel

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugrValores.Rows.All)
			{
				string scodigo = "";
				int i = 0;
				scodigo = Convert.ToString ( dr.Cells["Codigo"].Value) ;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.ugrValores.Rows.All)
				{
					if ( scodigo == Convert.ToString(dra.Cells["Codigo"].Value))
					{
						i++;
					}
				}
				if (i > 1)
				{
					if ( scodigo == Convert.ToString(dr.Cells["Codigo"].Value)&& Convert.ToInt32 (dr.Cells["Dupli"].Value) > 0)
						// &&  Convert.ToInt32 ( dra.Cells["Dupli"].Value) > 0 )
					{
						bEditarMeses = true;
						dr.Delete();
						bEditarMeses = false;
				
					}
				}
			}

			#region rellna idcuenta, codigo , cuen
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridCuentas.Rows.All)
			{
				string scodigo = "";
				int i = 0;
				int icuentas = 0;
				string sCuentas ="";

				icuentas = (int)dr.Cells["idCuenta"].Value;
				sCuentas = (string)	dr.Cells["Cuenta"].Value;

				scodigo = Convert.ToString ( dr.Cells["Codigo"].Value) ;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.ugrValores.Rows.All)
				{
					if ( scodigo == Convert.ToString(dra.Cells["Codigo"].Value) && Convert.ToInt32(dra.Cells["idCuenta"].Value)== 0)
					{
						dra.Cells["idCuenta"].Value =  icuentas;
						dra.Cells["Cuenta"].Value = sCuentas;
					}
				}
			}
			#endregion rellna idcuenta, codigo , cuen
			#region Selecciona el visto 
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugrValores.Rows.All)
			{
				int icuentas = 0;
				icuentas = Convert.ToInt32 ( dr.Cells["idCuenta"].Value);
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drr in this.uGridCuentas.Rows.All)
				{
					if ( icuentas == Convert.ToInt32 ( drr.Cells["idCuenta"].Value))
					{
						drr.Cells["Sel"].Value = true;
					}
				}
			}
			#endregion Selecciona el visto 
			#region sumatorio
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drA in this.ugrValores.Rows.All)
			{
				drA.Cells["TOTAL"].Value = 0;
        
				Decimal dEnero = 0.00m;
				Decimal dFebrero= 0.00m;
				Decimal dMarzo = 0.00m;
				Decimal dAbril = 0.00m;
				Decimal dMayo = 0.00m;
				Decimal dJunio = 0.00m;
				Decimal dJulio = 0.00m;
				Decimal dAgosto = 0.00m;
				Decimal dSeptiembre = 0.00m;
				Decimal dOctubre = 0.00m;
				Decimal dNoviembre = 0.00m;
				Decimal dDiciembre = 0.00m;

				if ((drA.Cells["ENERO"].Value) != System.DBNull.Value ) dEnero = Convert.ToDecimal(drA.Cells["ENERO"].Value);
				if ((drA.Cells["FEBRERO"].Value) != System.DBNull.Value ) dFebrero = Convert.ToDecimal(drA.Cells["FEBRERO"].Value);
				if ((drA.Cells["MARZO"].Value) != System.DBNull.Value ) dMarzo = Convert.ToDecimal(drA.Cells["MARZO"].Value);
				if ((drA.Cells["ABRIL"].Value) != System.DBNull.Value ) dAbril = Convert.ToDecimal(drA.Cells["ABRIL"].Value);
				if ((drA.Cells["MAYO"].Value) != System.DBNull.Value ) dMayo = Convert.ToDecimal(drA.Cells["MAYO"].Value);
				if ((drA.Cells["JUNIO"].Value) != System.DBNull.Value ) dJunio = Convert.ToDecimal(drA.Cells["JUNIO"].Value);
				if ((drA.Cells["JULIO"].Value) != System.DBNull.Value ) dJulio = Convert.ToDecimal(drA.Cells["JULIO"].Value);
				if ((drA.Cells["AGOSTO"].Value) != System.DBNull.Value ) dAgosto = Convert.ToDecimal(drA.Cells["AGOSTO"].Value);
				if ((drA.Cells["SEPTIEMBRE"].Value) != System.DBNull.Value ) dSeptiembre = Convert.ToDecimal(drA.Cells["SEPTIEMBRE"].Value);
				if ((drA.Cells["OCTUBRE"].Value) != System.DBNull.Value ) dOctubre = Convert.ToDecimal(drA.Cells["OCTUBRE"].Value);
				if ((drA.Cells["NOVIEMBRE"].Value) != System.DBNull.Value ) dNoviembre = Convert.ToDecimal(drA.Cells["NOVIEMBRE"].Value);
				if ((drA.Cells["DICIEMBRE"].Value) != System.DBNull.Value ) dDiciembre = Convert.ToDecimal(drA.Cells["DICIEMBRE"].Value);


				drA.Cells["TOTAL"].Value = 	( dEnero +
					dFebrero+
					dMarzo+
					dAbril +
					dMayo +
					dJunio +
					dJulio +
					dAgosto +
					dSeptiembre+ 
					dOctubre +
					dNoviembre+ 
					dDiciembre );

			}
			#endregion sumatorio
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("La estructura del archivo a subir debe ser: CODIGO, ENERO, FEBRERO, MARZO, ABRIL, MAYO, JUNIO\n JULIO, AGOSTO, SEPTIEMBRE, OCTUBRE, NOVIEMBRE, DICIEMBRE \n con nombre PRESUPUESTO, con la versión 2003 de Excel.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				
		}

		private void cmbCentroDeCosto_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCentroDeCosto.ActiveRow != null)
			{

				//this.cmbResponsable.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec VerCentroDecosto 2, {0}, {1}",(int) this.txtAnio.Value ,(int) this.cmbCentroDeCosto.Value));
				this.uGridCuentas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaCuentasPresupuesto 1, {0}, {1}", (int) this.txtAnio.Value, (int) this.cmbCentroDeCosto.Value));
				this.ugrValores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetallePresupuestoCC 0"));
			
			}
		}

		private void btnAnular_Click_1(object sender, System.EventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show(string.Format("esta Seguro de Anular el presupuesto. \n\n ¿ Desea Continuar ?"), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				
				using (frmNotas Nota = new frmNotas())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("UPDATE PresupuestoPorCC SET Estado = 3, UusarioAnula = '{0}', FechaAnula = Getdate (), NotaAnula = '{1}'  WHERE idPresupuestoPorCC = {2} ", 
							MenuLatinium.stUsuario, Nota.txtNotas.Text, (int) this.txtidPresupuestoPorCC.Value));

						Consultar ((int) this.txtidPresupuestoPorCC.Value);
					}
				}
			}
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbCampoEnteroVacio(this.txtAnio, "Ingrese el periodo", 2017, DateTime.Today.Year)) return;
			int iContPres = 0;
			iContPres = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select COUNT (1) from PresupuestoCCxCuenta where Periodo =  {0}", (int) this.txtAnio.Value));	
				
			if( iContPres > 0 ) 
			{
				MessageBox.Show(string.Format("Existen Cuentas Asignadas para el Periodo : {0}, comuniquese con el Administrador", (int) this.txtAnio.Value), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			#region Guardar
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

					

					#region Graba Maestro
					int idProyecto = 0;
					int idNomina = 0;
					if (this.cmbCentroDeCosto.ActiveRow != null) idProyecto = (int) this.cmbCentroDeCosto.Value;
					if(this.cmbResponsable.ActiveRow != null) idNomina = (int) this.cmbResponsable.Value;

					string  sSQL1 = string.Format("Exec [CreaPeriodoPresupuestoCCxCuenta] {0}, {1}, {2}, {3}, {4}" ,
						(int)this.txtidPresupuestoPorCC.Value, 
						(int)this.txtAnio.Value, 
						idProyecto,
						idNomina,
						(int)this.txtEstado.Value
						);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidPresupuestoPorCC.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					

					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("La asignación de cuentas se realizo correctamente.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
	
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
	}
}






//if (sCuentasNot.Length > 0 )
//{
//if (DialogResult.No ==	MessageBox.Show(string.Format("Cuentas:, '{0}' no Existen o no estan asignadas \n\n ¿ Desea Continuar ?", sCuentasNot), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
//					
//
//
//if (sCuentasNot.Length > 0 )
//{
//MessageBox.Show(string.Format("Cuentas:, {0} no Existen o no estan asignadas ", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//}
