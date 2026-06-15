using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmReposicionCarteraSolidario.
	/// </summary>
	public class frmReposicionCarteraSolidario : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label5;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiferencia;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdRolPrestamo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscador;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSeleccionado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReposicionCarteraSolidario()
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReposicionCarteraSolidario));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCruceAnticiposSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AnticipoNC");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 9, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 10, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCruceAnticiposSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRolPrestamo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AnticipoNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label5 = new System.Windows.Forms.Label();
			this.txtDiferencia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdRolPrestamo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscador = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.txtSeleccionado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSeleccionado)).BeginInit();
			this.SuspendLayout();
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(192, 75);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(54, 16);
			this.label5.TabIndex = 298;
			this.label5.Text = "Diferencia";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDiferencia
			// 
			this.txtDiferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiferencia.Enabled = false;
			this.txtDiferencia.FormatString = "#,##0.00";
			this.txtDiferencia.Location = new System.Drawing.Point(264, 72);
			this.txtDiferencia.Name = "txtDiferencia";
			this.txtDiferencia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiferencia.PromptChar = ' ';
			this.txtDiferencia.Size = new System.Drawing.Size(80, 22);
			this.txtDiferencia.TabIndex = 297;
			this.txtDiferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDiferencia_KeyDown);
			this.txtDiferencia.Click += new System.EventHandler(this.txtDiferencia_Click);
			this.txtDiferencia.Validated += new System.EventHandler(this.txtDiferencia_Validated);
			this.txtDiferencia.Enter += new System.EventHandler(this.txtDiferencia_Enter);
			// 
			// txtIdRolPrestamo
			// 
			this.txtIdRolPrestamo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdRolPrestamo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdRolPrestamo.Enabled = false;
			this.txtIdRolPrestamo.Location = new System.Drawing.Point(984, 144);
			this.txtIdRolPrestamo.Name = "txtIdRolPrestamo";
			this.txtIdRolPrestamo.PromptChar = ' ';
			this.txtIdRolPrestamo.Size = new System.Drawing.Size(16, 22);
			this.txtIdRolPrestamo.TabIndex = 296;
			this.txtIdRolPrestamo.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(1000, 144);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 295;
			this.txtEstado.Visible = false;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(416, 42);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(44, 16);
			this.label8.TabIndex = 294;
			this.label8.Text = "Número";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(472, 39);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 293;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(416, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 292;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(368, 75);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(30, 16);
			this.label3.TabIndex = 291;
			this.label3.Text = "Valor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 289;
			this.label2.Text = "Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 128);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 288;
			this.label1.Text = "Concepto";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.AllowMonthSelection = true;
			this.dtFecha.AutoSelectionUpdate = true;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.dtFecha.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(472, 8);
			this.dtFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 286;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// txtTotal
			// 
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(416, 72);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(80, 22);
			this.txtTotal.TabIndex = 290;
			// 
			// txtConcepto
			// 
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Enabled = false;
			this.txtConcepto.Location = new System.Drawing.Point(96, 104);
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(616, 66);
			this.txtConcepto.TabIndex = 287;
			this.txtConcepto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConcepto_KeyDown);
			// 
			// txtBuscador
			// 
			this.txtBuscador.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscador.Enabled = false;
			this.txtBuscador.Location = new System.Drawing.Point(72, 192);
			this.txtBuscador.Name = "txtBuscador";
			this.txtBuscador.Size = new System.Drawing.Size(168, 22);
			this.txtBuscador.TabIndex = 365;
			this.txtBuscador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscador_KeyDown);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "Serie";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 111;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Local";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 53;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 70;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 90;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 89;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 250;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 219;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.PromptChar = ' ';
			ultraGridColumn10.Width = 70;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 70;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 24;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 88;
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
																										 ultraGridColumn14});
			ultraGridBand1.Header.Caption = "";
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance6;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance7;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 224);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1026, 176);
			this.ultraGrid1.TabIndex = 366;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn5.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn13.DataType = typeof(bool);
			ultraDataColumn13.DefaultValue = false;
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
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
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Enabled = false;
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnBuscar.Location = new System.Drawing.Point(256, 192);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 23);
			this.btnBuscar.TabIndex = 368;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(8, 196);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel3.TabIndex = 367;
			this.ultraLabel3.Text = "Buscar";
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
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 176);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1042, 8);
			this.groupBox1.TabIndex = 369;
			this.groupBox1.TabStop = false;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.Image = ((System.Drawing.Image)(resources.GetObject("btnAsiento.Image")));
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.Location = new System.Drawing.Point(407, 424);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(80, 23);
			this.btnAsiento.TabIndex = 376;
			this.btnAsiento.Text = "Asiento";
			this.btnAsiento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 424);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 370;
			this.btnNuevo.Text = "Nuevo";
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
			this.btnConsultar.Location = new System.Drawing.Point(84, 424);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 371;
			this.btnConsultar.Text = "Consultar";
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
			this.btnEditar.Location = new System.Drawing.Point(166, 424);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 374;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(246, 424);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 373;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(327, 424);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 372;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(491, 424);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 377;
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
			this.btnCancelar.Location = new System.Drawing.Point(571, 424);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 378;
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
			this.btnSalir.Location = new System.Drawing.Point(647, 424);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 379;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 408);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1042, 8);
			this.groupBox2.TabIndex = 380;
			this.groupBox2.TabStop = false;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(1016, 144);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 381;
			this.txtIdAsiento.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn15,
																																 ultraDataColumn16});
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(8, 10);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(80, 16);
			this.label15.TabIndex = 383;
			this.label15.Text = "Forma de Pago";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource5;
			this.cmbFormaPago.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Personal";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 101;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 277;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 8);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(296, 21);
			this.cmbFormaPago.TabIndex = 382;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			this.cmbFormaPago.ValueChanged += new System.EventHandler(this.cmbFormaPago_ValueChanged);
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Width = 277;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Enabled = false;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(96, 40);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(296, 21);
			this.cmbNombre.TabIndex = 285;
			this.cmbNombre.ValueMember = "idPersonal";
			this.cmbNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNombre_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 75);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(69, 16);
			this.label6.TabIndex = 385;
			this.label6.Text = "Seleccionado";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSeleccionado
			// 
			this.txtSeleccionado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSeleccionado.Enabled = false;
			this.txtSeleccionado.FormatString = "#,##0.00";
			this.txtSeleccionado.Location = new System.Drawing.Point(96, 72);
			this.txtSeleccionado.Name = "txtSeleccionado";
			this.txtSeleccionado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSeleccionado.PromptChar = ' ';
			this.txtSeleccionado.Size = new System.Drawing.Size(80, 22);
			this.txtSeleccionado.TabIndex = 384;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(736, 88);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 386;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuario.Location = new System.Drawing.Point(736, 136);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 19);
			this.lblUsuario.TabIndex = 387;
			// 
			// lblContador
			// 
			this.lblContador.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(408, 194);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 388;
			// 
			// frmReposicionCarteraSolidario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1042, 454);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.txtSeleccionado);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.txtDiferencia);
			this.Controls.Add(this.txtIdRolPrestamo);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.txtConcepto);
			this.Controls.Add(this.txtBuscador);
			this.Controls.Add(this.ultraGrid1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmReposicionCarteraSolidario";
			this.Text = "Reposicion Cartera Solidario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReposicionCarteraSolidario_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmReposicionCarteraSolidario_Closing);
			this.Load += new System.EventHandler(this.frmReposicionCarteraSolidario_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSeleccionado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
	
		int iCont = 0;
		bool bValida = false;
		int idBloqueaTransacciones = 0;
		decimal dTotal = 0.00m;

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{	
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);			
		}
		
	

		private void Totales(bool bGuarda)
		{
			string sSQL = string.Format("Exec SumaAnticiposCapturados {0}", (int)this.txtIdRolPrestamo.Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				iCont = dr.GetInt32(0);
				this.txtSeleccionado.Value = dr.GetDecimal(1);
			}
			dr.Close();

			this.lblContador.Text = iCont.ToString() + " Registros Seleccionados";
			this.txtTotal.Value = System.Convert.ToDecimal(this.txtSeleccionado.Value) + System.Convert.ToDecimal(this.txtDiferencia.Value);

			if (!bGuarda)
				if (!bNuevo && bEdicion) 
					if (dTotal != Convert.ToDecimal(this.txtTotal.Value)) 
						this.btnCancelar.Enabled = false;
		}

		private void frmReposicionCarteraSolidario_Load(object sender, System.EventArgs e)
		{
			string stTipo = "0820";
			string stPermiso = "427";
							
			
			miAcceso = new Acceso(cdsSeteoF, "08320103");

			if (!Funcion.Permiso("957", cdsSeteoF))
			{		
		
				MessageBox.Show("No tiene Acceso Reposicion Cartera Solidario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;

			}
			else
			{
				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (2, 16)");	

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

				this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && bNuevo)
			{
				if (iCont > 0)
					if (DialogResult.No == MessageBox.Show(string.Format("¿Esta seguro de cancelar esta reposición con {0} clientes seleccionados?", iCont), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						return;
					
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec CancelaRolPrestamosSolidario {0}", (int)this.txtIdRolPrestamo.Value));
			}

			if (!bNuevo && bEdicion) 
				FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			bNuevo = false;
			bEdicion = false;

			this.txtIdRolPrestamo.Value = 0;
			this.txtIdAsiento.Value = 0;
			this.txtEstado.Value = 0;

			this.txtNumero.Text = "";
			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbNombre.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtSeleccionado.Value = 0;
			this.txtDiferencia.Value = 0;
			this.txtTotal.Value = 0;
			this.txtConcepto.Text = "";
			this.txtBuscador.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.lblContador.Text = "";

			this.BuscaAnticipos(0, "", 1);

			this.cmbFormaPago.Enabled = false;
			this.cmbNombre.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtSeleccionado.Enabled = false;
			this.txtDiferencia.Enabled = false;
			this.txtTotal.Enabled = false;
			this.txtConcepto.Enabled = false;
			this.txtBuscador.Enabled = false;
			this.btnBuscar.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;

			iCont = 0;
			idBloqueaTransacciones = 0;
			dTotal = 0.00m;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.cmbFormaPago.Enabled = true;
			
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			this.txtDiferencia.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.txtBuscador.Enabled = true;
			this.btnBuscar.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;			

			bNuevo = true;
			bEdicion = true;
			
			this.cmbFormaPago.Value = 2;
			this.dtFecha.Value = DateTime.Today;			
			this.txtEstado.Value = 0;

			DateTime dtFecha = DateTime.Today;
			string sSQL = string.Format("Exec GuardaRolPrestamo {0}, 8, '', 0, 0, '{1}', 0, 0, 0, '', 0", 
				(int)this.txtIdRolPrestamo.Value, dtFecha.ToString("yyyyMMdd"));
			this.txtIdRolPrestamo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			
			this.cmbFormaPago.Focus();
		}

		private void Consulta(int idRolPrestamo)
		{
			/*idRolPrestamo, idRolTipoPrestamo, Numero, idFormaPago, idPersonal, FechaEntrega, Seleccionado, Interes, Capital, Concepto, idAsiento, Estado, Case Estado When 0 Then 'PENDIENTE' When 1 Then 'AUTORIZADO EL PAGO' When 2 Then 'PAGADO' When 3 Then 'ANULADO' End As EstadoM, Usuario*/
			try
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAsiento.Enabled = false;
				this.btnAnular.Enabled = false;				

				string sSQL = string.Format("Exec ConsultaRolPrestamo {0}", idRolPrestamo);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdRolPrestamo.Value = dr.GetInt32(0);
					this.txtNumero.Value = dr.GetString(2);
					this.cmbFormaPago.Value = dr.GetInt32(3);
					this.cmbNombre.Value = dr.GetInt32(4);
					this.dtFecha.Value = dr.GetDateTime(5);
					this.txtSeleccionado.Value = dr.GetDecimal(6);
					this.txtDiferencia.Value = dr.GetDecimal(7);
					this.txtTotal.Value = dr.GetDecimal(8);
					this.txtConcepto.Value = dr.GetString(9);
					this.txtIdAsiento.Value = dr.GetInt32(10);
					this.txtEstado.Value = dr.GetInt32(11);
					this.lblEstado.Text = dr.GetString(12);
					this.lblUsuario.Text = dr.GetString(13);
				}
				dr.Close();
				
				BuscaAnticipos((int)this.txtIdRolPrestamo.Value, "", 1);

				if ((int)this.txtEstado.Value != 3)
				{
					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)
					{
						if ((int)this.cmbFormaPago.Value == 2) if ((int)this.txtEstado.Value == 0) this.btnEditar.Enabled = true;
						if ((int)this.cmbFormaPago.Value == 16) this.btnEditar.Enabled = true;
					}
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;					
					if (miAcceso.CrearAsiento) if ((int)this.txtEstado.Value == 2) this.btnAsiento.Enabled = true;
					if (miAcceso.Anular) 
					{
						if ((int)this.cmbFormaPago.Value == 2) if ((int)this.txtEstado.Value == 0) this.btnAnular.Enabled = true;
						if ((int)this.cmbFormaPago.Value == 16) this.btnAnular.Enabled = true;
					}
				}
				this.btnCancelar.Enabled = true;	
			}	
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			using (frmBuscaRolPrestamo Busqueda = new frmBuscaRolPrestamo(8))
			{
				if (Busqueda.ShowDialog() == DialogResult.OK)
					this.Consulta((int)Busqueda.uGridDatos.ActiveRow.Cells["idRolPrestamo"].Value);				
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			Consulta((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.cmbFormaPago.Value == 2)
			{
				if ((int)this.txtEstado.Value == 1)
				{
					MessageBox.Show("El registro ya se encuentra con el pago autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)this.txtEstado.Value == 2)
				{
					MessageBox.Show("El registro ya esta pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}			
			}

			if ((int)this.cmbFormaPago.Value == 16)
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bancos"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No Puede Editar este registro, El Período '{0}' esta Cerrado, \n\nConsulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
					return;
				}	

				if (!FuncionesContabilidad.bValidaConciliado((int)this.txtIdAsiento.Value, Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106"), cdsSeteoF)) return;

				if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion(Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106"), (DateTime)this.dtFecha.Value, cdsSeteoF)) return;
			}

			if (Funcion.ValidaBloqueo((int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF)) return;

			FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, true, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

			this.cmbFormaPago.Enabled = true;
			if ((int)this.cmbFormaPago.Value == 2) this.cmbNombre.Enabled = true; else this.cmbNombre.Enabled = false;
			this.dtFecha.Enabled = true;
			this.txtDiferencia.Enabled = true;
			this.txtConcepto.Enabled = true;

			this.txtBuscador.Enabled = true;
			this.btnBuscar.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAsiento.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;

			bNuevo = false;
			bEdicion = true;			
		}

		private void Guardar ()
		{
			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			if (bNuevo) this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, "Select dbo.SecuencialPrestamoEmpleados(8)");
			
			int idPersonal = 0;
			if (this.cmbNombre.ActiveRow != null) idPersonal = (int)this.cmbNombre.Value;

			string sSQL = string.Format("Exec GuardaRolPrestamo {0}, 8, '{1}', {2}, {3}, '{4}', {5}, {6}, {7}, '{8}', {9}", 
				(int)this.txtIdRolPrestamo.Value, this.txtNumero.Text.ToString(), (int)this.cmbFormaPago.Value, idPersonal, dtFecha.ToString("yyyyMMdd"), 
				Convert.ToDecimal(this.txtSeleccionado.Value), Convert.ToDecimal(this.txtDiferencia.Value), Convert.ToDecimal(this.txtTotal.Value), 
				this.txtConcepto.Text.ToString(), (int)this.txtEstado.Value);
			this.txtIdRolPrestamo.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			Totales(true);

			#region Validacion
			if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione la forma de pago")) return;
			
			#region Emision de cheque
			if ((int)this.cmbFormaPago.Value == 2)
			{
				if (!Validacion.vbComboVacio(this.cmbNombre, "Seleccione el beneficiario")) return;
						
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", false, "", cdsSeteoF)) return;
			}
			#endregion Emision de cheque

			#region Debito Bancario
			if ((int)this.cmbFormaPago.Value == 16)
			{
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione una fecha", true, "Bancos", cdsSeteoF)) return;

				if (!Validacion.vbTexto(this.txtNumero, 5, 15, "Numero")) return;
				
				int idCuenta = Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106");

				if (!FuncionesContabilidad.vbValidaDocumentoBancario(idCuenta, this.txtNumero, 4, (int)this.txtIdAsiento.Value, cdsSeteoF)) return;

				if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion(idCuenta, (DateTime)this.dtFecha.Value, cdsSeteoF)) return;
			}				
			#endregion Debito Bancario

			#region Valida seleccion de pagos
			if (!Validacion.vbTexto(this.txtConcepto, 10, 100, "Concepto")) return;
			
			if (iCont == 0)
			{
				MessageBox.Show("Seleccione los Pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Convert.ToDecimal(this.txtSeleccionado.Value) == 0.00m)
			{
				MessageBox.Show("No puede crear una reposicion en cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Convert.ToDecimal(this.txtSeleccionado.Value) <= Convert.ToDecimal(this.txtDiferencia.Value))
			{
				MessageBox.Show("El valor diferencia no puede ser mayor al valor seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDiferencia.Focus();
				return;
			}
			
			if (Convert.ToDecimal(this.txtTotal.Value) == 0.00m)
			{
				MessageBox.Show("No puede crear una reposicion en cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Valida seleccion de pagos
			#endregion Validacion	

			if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de Crear esta Reposición de {1} clientes,\n\nCon la suma de {0}\n\nCon Forma de Pago '{2}'?", Convert.ToDecimal(this.txtTotal.Value), iCont, this.cmbFormaPago.Text.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				this.Guardar();

				if ((int)this.cmbFormaPago.Value == 16)
				{					
					if (!bNuevo) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From AutorizacionDePagos Where idTipo = 8 And idOrigen = {0}", (int)this.txtIdRolPrestamo.Value));
					
					string sSQL = string.Format("Exec AsientoNDReposicionSolidario {0}, '{1}', '{2}', {3}", 
						(int)this.txtIdRolPrestamo.Value, this.txtNumero.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.txtIdAsiento.Value);
					this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);				
				}
				else if ((int)this.cmbFormaPago.Value == 2)
				{
					#region Autorización de Pagos
					string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 8, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'CRUCE ANTICIPOS', 0", 
						(int)this.txtIdRolPrestamo.Value, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"), Convert.ToDecimal(this.txtTotal.Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLAP, true);					
					#endregion Autorización de Pagos
				}
				
				FuncionesProcedimientos.RegistraBloqueoTransacciones (idBloqueaTransacciones, false, (int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF);

				this.Consulta((int)this.txtIdRolPrestamo.Value);

				#region Controles
				this.cmbFormaPago.Enabled = false;
				this.cmbNombre.Enabled = false;
				this.dtFecha.Enabled = false;
				this.txtDiferencia.Enabled = false;
				this.txtConcepto.Enabled = false;

				this.txtBuscador.Enabled = false;
				this.btnBuscar.Enabled = false;

				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

				bNuevo = false;
				bEdicion = false;
				#endregion Controles
			}
		}

		private void cmbFormaPago_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if ((int)this.cmbFormaPago.Value == 2) 
				{
					this.txtEstado.Value = 0;
					this.cmbNombre.Enabled = true;
					this.txtNumero.Text = "";
					this.txtNumero.Enabled = false;						
				}
				else if ((int)this.cmbFormaPago.Value == 16) 
				{
					this.txtEstado.Value = 2;
					this.cmbNombre.Enabled = false;
					this.cmbNombre.Value = System.DBNull.Value;
					this.txtNumero.Enabled = true;
				}
			}				
		}

		private void txtBuscador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void BuscaAnticipos(int idRolPrestamo, string sFiltro, int iConsulta)
		{
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaReposicionCarteraSolidario {0}, '{1}', {2}", idRolPrestamo, sFiltro, iConsulta));
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.cmbFormaPago.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la forma de pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cmbNombre.Focus();
				return;
			}

			if ((int)this.cmbFormaPago.Value == 2)
			{
				if (this.cmbNombre.ActiveRow == null)
				{
					MessageBox.Show("Seleccione el beneficiario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbNombre.Focus();
					return;
				}
			}

			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione una fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}		

			if (this.txtConcepto.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el concepto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtConcepto.Focus();
				return;
			}
			#endregion Validacion	

			BuscaAnticipos((int)this.txtIdRolPrestamo.Value, this.txtBuscador.Text.ToString(), 0);
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
				if (this.cmbFormaPago.ActiveRow != null)
					if (miAcceso.CambiarFecha) this.dtFecha.Focus(); else if ((int)this.cmbFormaPago.Value == 2) this.cmbNombre.Focus(); else this.txtDiferencia.Focus();
		}

		private void cmbNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiferencia.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
				if ((int)this.cmbFormaPago.Value == 2) this.cmbNombre.Focus(); else this.txtNumero.Focus();
		}

		private void txtDiferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConcepto.Focus();
		}

		private void txtConcepto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscador.Focus();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			Consulta((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (miAcceso.CrearAsiento)
			{
				if ((int) this.txtIdAsiento.Value == 0) MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				else
				{					
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 0) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtIdAsiento.Value), true))
					{
						Cursor = Cursors.WaitCursor;
						frmAsientos Asientos = new frmAsientos((int)this.txtIdAsiento.Value);
						Asientos.MdiParent = MdiParent;
						Asientos.Show();
						Cursor = Cursors.Default;
					}
					else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 1) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtIdAsiento.Value), true))
					{
						MessageBox.Show("El Asiento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
					else if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", (int)this.txtIdAsiento.Value), true))
					{
						MessageBox.Show("El Asiento fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
			}
		}		

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			Consulta((int)this.txtIdRolPrestamo.Value);

			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!miAcceso.Anular)
			{
				MessageBox.Show("No tiene acceso a Anular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.cmbFormaPago.Value == 2)
			{
				if ((int)this.txtEstado.Value == 1)
				{
					MessageBox.Show("El registro ya se encuentra con el pago autorizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}

				if ((int)this.txtEstado.Value == 2)
				{
					MessageBox.Show("El registro ya esta pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}			
			}

			if ((int)this.cmbFormaPago.Value == 16)
			{
				if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Bancos"))
				{
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					MessageBox.Show(string.Format("No Puede Editar este registro, El Período '{0}' esta Cerrado, \n\nConsulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
					return;
				}	

				if (FuncionesContabilidad.bValidaConciliado((int)this.txtIdAsiento.Value, Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106"), cdsSeteoF)) return;

				if (!FuncionesContabilidad.vBancosValidaEstadoConciliacion(Funcion.iEscalarSQL(cdsSeteoF, "Select idCuenta From AsignaCuenta Where idAsignacion = 106"), (DateTime)this.dtFecha.Value, cdsSeteoF)) return;
			}
			
			if (Funcion.ValidaBloqueo((int)this.txtIdRolPrestamo.Value, 10, cdsSeteoF)) return;

			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de Anular esta Reposición?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						try
						{
							string sSQL = string.Format("Exec AnulaRolPrestamoSolidario {0}, '{1}'", (int)this.txtIdRolPrestamo.Value, Nota.txtNotas.Text.ToString());
							Funcion.EjecutaSQL(cdsSeteoF, sSQL);

							this.Consulta((int)this.txtIdRolPrestamo.Value);
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
		}

		private void txtDiferencia_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.txtDiferencia.Value == System.DBNull.Value) this.txtDiferencia.Value = 0.00m;
				Totales(false);
			}
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Sel")
			{
				this.ultraGrid1.UpdateData();
							
				string sSQLCruce = string.Format("Exec GCruceAnticipos {0}, {1}, {2}, {3}, {4}, {5}", 
					(int)e.Cell.Row.Cells["idCruceAnticiposSolidario"].Value, (int)e.Cell.Row.Cells["idAnticipo"].Value, 
					(int)this.txtIdRolPrestamo.Value, Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value), (bool)e.Cell.Row.Cells["Sel"].Value, 
					(bool)e.Cell.Row.Cells["AnticipoNC"].Value);
				e.Cell.Row.Cells["idCruceAnticiposSolidario"].Value =	Funcion.iEscalarSQL(cdsSeteoF, sSQLCruce, true);

				if (!(bool)e.Cell.Row.Cells["Sel"].Value) e.Cell.Row.Cells["Valor"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value);

				this.Totales(false);
			}			
		}

		private void frmReposicionCarteraSolidario_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una Reposición", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmReposicionCarteraSolidario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdRolPrestamo.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);																	 
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0820");

						this.Consulta((int)this.txtIdRolPrestamo.Value);					
					}
				}
				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) this.btnCancelar_Click(sender, e);
				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			if ((int)this.txtEstado.Value == 3)
			{
				MessageBox.Show("El registro esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField Compra = new ParameterField ();
			ParameterDiscreteValue IdCompra = new ParameterDiscreteValue ();
			Compra.ParameterFieldName = "@idRolPrestamo";
			IdCompra.Value = int.Parse(this.txtIdRolPrestamo.Value.ToString());
			Compra.CurrentValues.Add (IdCompra);
			paramFields.Add (Compra);
			
			Reporte miRep = new Reporte("ReposicionCarteraSolidario.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Reposición Cartera Solidario");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
						
			Cursor = Cursors.Default;
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}		

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			if (e.Cell.Column.ToString() == "Valor")
			{	
				if (e.Cell.Row.Cells["idCruceAnticiposSolidario"].Value == DBNull.Value) return;
			
				if (!bValida)
				{
					string sSQLSaldo = string.Format("Declare @Saldo Decimal(18, 2) SELECT @Saldo = ISNULL(SUM(Valor), 0) FROM CruceAnticiposSolidario WHERE idAnticipo = {0} And idRolPrestamo <> {1} Select ISNULL(@Saldo, 0)", 
						(int)e.Cell.Row.Cells["idAnticipo"].Value, (int)this.txtIdRolPrestamo.Value);
					decimal dUsado = Funcion.decEscalarSQL(cdsSeteoF, sSQLSaldo);
					
					decimal dSaldo = Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value) - dUsado;

					if (Convert.ToDecimal(e.Cell.Value) > dSaldo)
					{					
						bValida = true;
						e.Cell.Value = Convert.ToDecimal(e.Cell.OriginalValue);
						MessageBox.Show(string.Format("No puede ser mayor a {0}.", dSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
						bValida = false;
						return;
					}
								
					e.Cell.Row.Cells["Saldo"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Total"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value) - dUsado;

					string sSQL = string.Format("Update CruceAnticiposSolidario Set Valor = {1} Where idCruceAnticiposSolidario = {0}", 
						(int)e.Cell.Row.Cells["idCruceAnticiposSolidario"].Value, Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value));
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}

				if ((bool)e.Cell.Row.Cells["Sel"].Value) 
				{
					if (!(bool)e.Cell.Row.Cells["AnticipoNC"].Value)					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 24, 9", (int)e.Cell.Row.Cells["idAnticipo"].Value));
					if ((bool)e.Cell.Row.Cells["AnticipoNC"].Value)					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec ActualizaSaldoFacturas {0}, 5, 9", (int)e.Cell.Row.Cells["idAnticipo"].Value));

					this.Totales(false);
				}
			}
		}

		private void txtDiferencia_Click(object sender, System.EventArgs e)
		{
			this.txtDiferencia.SelectAll();
		}

		private void txtDiferencia_Enter(object sender, System.EventArgs e)
		{
			this.txtDiferencia.SelectAll();
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				this.cmbNombre.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaEmpleadosAnticiposPrestamos '{0}', 8, 0", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDiferencia.Focus();
		}		
	}
}

