using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_EvaluacionPersonalVentas.
	/// </summary>
	public class frmCom_EvaluacionPersonalVentas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblContador;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridVendedores;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargos;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbColores;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesde;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtHasta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_EvaluacionPersonalVentas()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_EvaluacionPersonalVentas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Faltas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Mensual");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Carpa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiarioGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MensualGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VendidoGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MetaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiferenciaDiaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeDiario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ComisionTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Contador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idColor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("btnObservaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			this.label1 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVendedores = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbColores = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.txtDesde = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVendedores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 106);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(39, 16);
			this.label1.TabIndex = 651;
			this.label1.Text = "Buscar";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(168, 10);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 650;
			this.label4.Text = "Hasta";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 10);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(37, 16);
			this.label5.TabIndex = 649;
			this.label5.Text = "Desde";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(496, 104);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 652;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(328, 103);
			this.btnVer.Name = "btnVer";
			this.btnVer.TabIndex = 645;
			this.btnVer.Text = "&Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// txtBuscar
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance1;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(56, 104);
			this.txtBuscar.MaxLength = 100;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(264, 21);
			this.txtBuscar.TabIndex = 644;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// dtHasta
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance2;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(216, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 648;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// dtDesde
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance3;
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
			this.dtDesde.TabIndex = 647;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDesde_KeyDown);
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(8, 34);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(31, 16);
			this.lblBodega.TabIndex = 655;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
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
			ultraDataColumn24.DataType = typeof(System.Decimal);
			ultraDataColumn24.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn25.DefaultValue = 0;
			ultraDataColumn26.DataType = typeof(int);
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
																																 ultraDataColumn26});
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
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 32);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(264, 21);
			this.cmbBodega.TabIndex = 653;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29});
			// 
			// uGridVendedores
			// 
			this.uGridVendedores.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVendedores.Cursor = System.Windows.Forms.Cursors.Default;
			appearance4.BackColor = System.Drawing.Color.White;
			this.uGridVendedores.DisplayLayout.Appearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVendedores.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.uGridVendedores.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridVendedores.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridVendedores.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridVendedores.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 8.5F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVendedores.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.uGridVendedores.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVendedores.DisplayLayout.Override.RowAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVendedores.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			this.uGridVendedores.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVendedores.Location = new System.Drawing.Point(8, 152);
			this.uGridVendedores.Name = "uGridVendedores";
			this.uGridVendedores.Size = new System.Drawing.Size(1000, 208);
			this.uGridVendedores.TabIndex = 654;
			this.uGridVendedores.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.uGridVendedores_InitializeRow);
			this.uGridVendedores.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVendedores_KeyDown);
			this.uGridVendedores.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVendedores_ClickCellButton);
			this.uGridVendedores.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVendedores_AfterCellUpdate);
			this.uGridVendedores.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVendedores_InitializeLayout);
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
			// ultraDataSource3
			// 
			ultraDataColumn30.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// cmbCargos
			// 
			this.cmbCargos.CausesValidation = false;
			this.cmbCargos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargos.DataSource = this.ultraDataSource3;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 288;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargos.Location = new System.Drawing.Point(56, 56);
			this.cmbCargos.MaxDropDownItems = 30;
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(264, 21);
			this.cmbCargos.TabIndex = 656;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCargos_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 58);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 657;
			this.label2.Text = "Cargo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 82);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(39, 16);
			this.label3.TabIndex = 659;
			this.label3.Text = "Estado";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstados
			// 
			this.cmbEstados.CausesValidation = false;
			this.cmbEstados.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource4;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayMember = "Estado";
			this.cmbEstados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstados.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(56, 80);
			this.cmbEstados.MaxDropDownItems = 30;
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(264, 21);
			this.cmbEstados.TabIndex = 658;
			this.cmbEstados.ValueMember = "idCom_Estado";
			this.cmbEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstados_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn32.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn32,
																																 ultraDataColumn33});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1016, 8);
			this.groupBox1.TabIndex = 660;
			this.groupBox1.TabStop = false;
			// 
			// cmbColores
			// 
			this.cmbColores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColores.DataSource = this.ultraDataSource5;
			this.cmbColores.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 198;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			this.cmbColores.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbColores.DisplayMember = "Color";
			this.cmbColores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbColores.Location = new System.Drawing.Point(264, 176);
			this.cmbColores.Name = "cmbColores";
			this.cmbColores.Size = new System.Drawing.Size(200, 112);
			this.cmbColores.TabIndex = 688;
			this.cmbColores.ValueMember = "idColor";
			this.cmbColores.Visible = false;
			// 
			// ultraDataSource5
			// 
			ultraDataColumn34.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn34,
																																 ultraDataColumn35});
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.txtDesde);
			this.groupBox2.Controls.Add(this.txtHasta);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.label7);
			this.groupBox2.Location = new System.Drawing.Point(728, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(280, 56);
			this.groupBox2.TabIndex = 689;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Porcentajes";
			// 
			// txtDesde
			// 
			this.txtDesde.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDesde.Location = new System.Drawing.Point(64, 24);
			this.txtDesde.MaxValue = 999;
			this.txtDesde.MinValue = 0;
			this.txtDesde.Name = "txtDesde";
			this.txtDesde.PromptChar = ' ';
			this.txtDesde.Size = new System.Drawing.Size(72, 21);
			this.txtDesde.TabIndex = 671;
			this.txtDesde.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDesde_KeyDown);
			// 
			// txtHasta
			// 
			this.txtHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtHasta.Location = new System.Drawing.Point(200, 24);
			this.txtHasta.MaxValue = 999;
			this.txtHasta.MinValue = 0;
			this.txtHasta.Name = "txtHasta";
			this.txtHasta.PromptChar = ' ';
			this.txtHasta.Size = new System.Drawing.Size(72, 21);
			this.txtHasta.TabIndex = 673;
			this.txtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHasta_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(152, 26);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 674;
			this.label6.Text = "Hasta";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(16, 26);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(37, 16);
			this.label7.TabIndex = 672;
			this.label7.Text = "Desde";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource6
			// 
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn36});
			// 
			// btnExcel
			// 
			this.btnExcel.CausesValidation = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(408, 103);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.TabIndex = 690;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Visible = false;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(bool);
			ultraDataColumn39.DefaultValue = false;
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39});
			// 
			// frmCom_EvaluacionPersonalVentas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1016, 366);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.cmbColores);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.uGridVendedores);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbCargos);
			this.KeyPreview = true;
			this.Name = "frmCom_EvaluacionPersonalVentas";
			this.Text = "Evaluación Personal de Ventas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_EvaluacionPersonalVentas_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_EvaluacionPersonalVentas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVendedores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
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

		int iBodega =0;
		private void frmCom_EvaluacionPersonalVentas_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200541");
			
			if (!Funcion.Permiso("2067", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Evaluación por Personal", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Visible = true;

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodega = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
			this.Cargos ();
			//this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCargo, Cargo From Com_CargosDeVentas Where Comisiona = 1 And idCargo <> 7 Order By Orden");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idCom_Estado, Estado From Com_Estado Where idCom_Estado <> 2");
			this.cmbColores.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idColor, Color From ArticuloColor Where idColor In (1, 13, 15)");

			if (iBodega > 0) this.cmbBodega.Value = iBodega; else this.cmbBodega.Enabled = true;

			DateTime dtFecha = DateTime.Parse("01/06/2016");

			this.dtDesde.CalendarInfo.MinDate = dtFecha;
			this.dtHasta.CalendarInfo.MinDate = dtFecha;

			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.dtHasta.Value = DateTime.Today;

			this.dtDesde.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);
			this.dtHasta.CalendarInfo.MaxDate = Convert.ToDateTime(this.dtHasta.Value);

			this.txtDesde.Value = 0;
			this.txtHasta.Value = 999;
			this.Locales ();
			
			this.txtBuscar.Focus();
		}

		int iSemana = 0;

		private void Locales()
		{

			string sSQL = string.Format("Exec Com_ResumenComisionesListaLocales {0}, {1}, {2}, '{3}'", 
				MenuLatinium.IdUsuario, iSemana, 0, Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridVendedores.Rows.Count > 0)
			{
				#region Informacion Basica
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["idNomina"].Hidden = true;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["idPersonal"].Hidden = true;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Bodega"].Hidden = true;
								
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Local"].Width = 130;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Local"].CellActivation = Activation.ActivateOnly;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Codigo"].Width = 100;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Codigo"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Codigo"].Hidden = true;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendedor"].Width = 220;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendedor"].CellActivation = Activation.ActivateOnly;
				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Entrada"].Width = 75;				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Entrada"].CellActivation = Activation.ActivateOnly;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vence"].Width = 75;				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vence"].CellActivation = Activation.ActivateOnly;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Tiempo"].Width = 100;				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Tiempo"].CellActivation = Activation.ActivateOnly;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Cargo"].Width = 150;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Cargo"].CellActivation = Activation.ActivateOnly;
				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estado"].Width = 75;				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estado"].CellActivation = Activation.ActivateOnly;	

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Sector"].Width = 180;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Sector"].CellActivation = Activation.ActivateOnly;
				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Faltas"].Width = 150;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Faltas"].CellActivation = Activation.ActivateOnly;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vacaciones"].Width = 150;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vacaciones"].CellActivation = Activation.ActivateOnly;
				#endregion Informacion Basica

				#region Columnas de Comisiones
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Mensual"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Mensual"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Mensual"].Format = "##0.00";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Mensual"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendido"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendido"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendido"].Format = "##0.00";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Vendido"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Carpa"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Carpa"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Carpa"].Format = "##0.00";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Carpa"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["VendidoGE"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["VendidoGE"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["VendidoGE"].Format = "##0.00";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["VendidoGE"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Diferencia"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Diferencia"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Diferencia"].Format = "##0.00";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Diferencia"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Porcentaje"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Porcentaje"].CellActivation = Activation.ActivateOnly;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Porcentaje"].Format = "##0.00";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Porcentaje"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				#endregion Columnas de Comisiones
				
				#region Columnas de Edicion
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Ultima_Gestion"].Width = 100;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Ultima_Gestion"].Format = "dd/MM/yyyy HH:mm";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Ultima_Gestion"].Header.Caption = "Ult. Gestión";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Ultima_Gestion"].CellActivation = Activation.ActivateOnly;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["idColor"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["idColor"].Style = ColumnStyle.DropDown;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["idColor"].ValueList = this.cmbColores;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["idColor"].Header.Caption = "Color";

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Observaciones"].Width = 280;
			
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["FechaReemplazo"].Width = 80;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["FechaReemplazo"].Format = "dd/MM/yyyy";
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["FechaReemplazo"].Header.Caption = "Fec. Reempl";
				#endregion Columnas de Edicion

				#region Columnas del Historial
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estrellas"].Width = 60;
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estrellas"].CellActivation = Activation.ActivateOnly;				
				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Estrellas"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;

				this.uGridVendedores.DisplayLayout.Bands[0].Columns["Orden"].Hidden = true;

				if (this.uGridVendedores.Rows.Count > 0)
				{
					string nombrCol = "";
							
					int iColumnas = (int)this.uGridVendedores.DisplayLayout.Bands[0].Columns.Count;
				
					for (int j = 26; j < iColumnas; j++)
					{
						nombrCol = this.uGridVendedores.DisplayLayout.Bands[0].Columns[j].Header.Caption;
						
						this.uGridVendedores.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.ActivateOnly;
						this.uGridVendedores.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
						this.uGridVendedores.DisplayLayout.Bands[0].Columns[nombrCol].Width = 50;
					}
				}
				#endregion Columnas del Historial

				if(miAcceso.Nuevo)
				{
					this.uGridVendedores.DisplayLayout.Bands[0].Columns.Add("btnObservaciones");
					this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnObservaciones"].Width = 40;
					this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnObservaciones"].Style = ColumnStyle.Button;
					this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnObservaciones"].ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
					this.uGridVendedores.DisplayLayout.Bands[0].Columns["btnObservaciones"].Header.Caption = "...";
				}
			}
			#endregion Diseño Grid
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/06/2016");
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

		private void uGridVendedores_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idColor")
			{
				int idColor = 0;
				if (e.Cell.Row.Cells["idColor"].Value != System.DBNull.Value) idColor = (int)e.Cell.Row.Cells["idColor"].Value;

				if (idColor == 0) e.Cell.Row.Appearance.BackColor = Color.White;
				if (idColor == 1) e.Cell.Row.Appearance.BackColor = Color.Yellow;
				if (idColor == 13) e.Cell.Row.Appearance.BackColor = Color.Red;
				if (idColor == 15) e.Cell.Row.Appearance.BackColor = Color.Green;

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Nomina Set idColor = {0} Where idNomina = {1}", 
					idColor, (int)e.Cell.Row.Cells["idNomina"].Value));
			}

			if (e.Cell.Column.ToString() == "Observaciones")
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Nomina Set Observaciones = '{0}' Where idNomina = {1}", 
					e.Cell.Row.Cells["Observaciones"].Value.ToString(), (int)e.Cell.Row.Cells["idNomina"].Value));
			}

			if (e.Cell.Column.ToString() == "FechaReemplazo")
			{
				if (e.Cell.Row.Cells["FechaReemplazo"].Value == System.DBNull.Value) 
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Nomina Set FechaReemplazo = NULL Where idNomina = {0}", 
						(int)e.Cell.Row.Cells["idNomina"].Value));
				else
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Nomina Set FechaReemplazo = '{0}' Where idNomina = {1}", 
						Convert.ToDateTime(e.Cell.Row.Cells["FechaReemplazo"].Value).ToString("yyyyMMdd"), (int)e.Cell.Row.Cells["idNomina"].Value));
			}
		}

		private void uGridVendedores_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnObservaciones")
			{	
				using (frmCom_Observaciones CO = new frmCom_Observaciones((int)e.Cell.Row.Cells["idPersonal"].Value, (int)e.Cell.Row.Cells["Bodega"].Value))
				{
					if (DialogResult.OK == CO.ShowDialog())
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Com_ActualizaObservaciones '{0}', {1}", 
							CO.txtObservaciones.Text.ToString(), (int)e.Cell.Row.Cells["idPersonal"].Value));
					}
				}
			}
		}

		private void uGridVendedores_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			int idColor = 0;
			if (e.Row.Cells["idColor"].Value != System.DBNull.Value) idColor = (int)e.Row.Cells["idColor"].Value;

			if (idColor == 0) e.Row.Appearance.BackColor = Color.White;
			if (idColor == 1) e.Row.Appearance.BackColor = Color.Yellow;
			if (idColor == 13) e.Row.Appearance.BackColor = Color.Red;
			if (idColor == 15) e.Row.Appearance.BackColor = Color.Green;
		}

		private void uGridVendedores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVendedores);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridVendedores);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{

			//			if( MenuLatinium.iNivel == 23) if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la bodega")) return;
			//      if( MenuLatinium.iNivel == 25) if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione la bodega")) return;
			int iBodegas = 0;
			if (this.cmbBodega.ActiveRow != null) iBodegas = (int)this.cmbBodega.Value;
			int idCargo = 0;
			if (this.cmbCargos.ActiveRow != null) idCargo = (int)this.cmbCargos.Value;
			int idEstado = 0;
			if (this.cmbEstados.ActiveRow != null) idEstado = (int)this.cmbEstados.Value;

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

			this.uGridVendedores.DataSource = null;

			string sSQL = string.Format("Exec Com_EvaluacionPersonal '{0}', '{1}', {2}, {3}, {4}, '{5}', {6}, {7}, {8}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				iBodegas, idCargo, idEstado, this.txtBuscar.Text.ToString(), (int)this.txtDesde.Value, (int)this.txtHasta.Value, MenuLatinium.IdUsuario);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridVendedores);

			this.lblContador.Text = this.uGridVendedores.Rows.Count + " Registros Encontrados";

			DiseñoGrid();
		}

		private void Cargos()
		{
			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Com_ResumenComisionesListaCargos {0}, {1}, {2}", 
				MenuLatinium.IdUsuario, MenuLatinium.iNivel, 0));
		}

		private void uGridVendedores_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void frmCom_EvaluacionPersonalVentas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCargos.Focus();
		}

		private void dtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtHasta.Focus();
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbCargos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEstados.Focus();
		}

		private void cmbEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void txtDesde_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtHasta.Focus();
		}

		private void txtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}
	}
}
