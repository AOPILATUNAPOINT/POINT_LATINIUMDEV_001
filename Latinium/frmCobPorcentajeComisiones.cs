using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCobPorcentajeComisiones.
	/// </summary>
	public class frmCobPorcentajeComisiones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnGuardar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargoCarpa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcReportes;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl TbBasico;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl10;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTramo;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label4;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label1;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.Button btnCopiar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCobPorcentajeComisiones()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCobPorcentajeComisiones));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCobPorcentajeComisiones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargoCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Descending, false);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCobPorcentajeComisiones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iCob_Penalizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargoCarpa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasDesde");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasHasta");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDesde");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentanjeHasta");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Penalizacion");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Penalizacion", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Penalizacion", 7, true);
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 6, true);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PorcentanjeHasta", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PorcentanjeHasta", 5, true);
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "PorcDesde", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "PorcDesde", 4, true);
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iCob_Penalizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcDesde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentanjeHasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Penalizacion");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.TbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnComandos = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargoCarpa = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnExcel = new System.Windows.Forms.Button();
			this.ultraTabPageControl10 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnCopiar = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.cmbTramo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.utcReportes = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btnSalir = new System.Windows.Forms.Button();
			this.TbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoCarpa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.ultraTabPageControl10.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTramo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).BeginInit();
			this.utcReportes.SuspendLayout();
			this.SuspendLayout();
			// 
			// TbBasico
			// 
			this.TbBasico.Controls.Add(this.btnComandos);
			this.TbBasico.Controls.Add(this.label4);
			this.TbBasico.Controls.Add(this.dtDesde);
			this.TbBasico.Controls.Add(this.cmbCargos);
			this.TbBasico.Controls.Add(this.uGridInformacion);
			this.TbBasico.Controls.Add(this.cmbCargoCarpa);
			this.TbBasico.Controls.Add(this.btnGuardar);
			this.TbBasico.Controls.Add(this.btnExcel);
			this.TbBasico.Location = new System.Drawing.Point(-10000, -10000);
			this.TbBasico.Name = "TbBasico";
			this.TbBasico.Size = new System.Drawing.Size(652, 337);
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(416, 8);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.TabIndex = 977;
			this.btnComandos.Text = "Copiar ";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 8);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 963;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDesde
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance1;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton1);
			this.dtDesde.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtDesde.Format = "MM/yyyy ";
			this.dtDesde.Location = new System.Drawing.Point(72, 8);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 962;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			this.dtDesde.ValueChanged += new System.EventHandler(this.dtDesde_ValueChanged);
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(752, 448);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(24, 8);
			this.cmbCargos.TabIndex = 763;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInformacion.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance2;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn4.Header.Caption = "Cargo";
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 250;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Width = 80;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 80;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance5;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellButtonAppearance = appearance6;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Width = 80;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance7;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellButtonAppearance = appearance8;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 120;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			ultraGridBand2.Header.Caption = "Periodo A";
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 40);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(640, 288);
			this.uGridInformacion.TabIndex = 957;
			this.uGridInformacion.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridInformacion_AfterRowInsert);
			this.uGridInformacion.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridInformacion_BeforeRowsDeleted);
			this.uGridInformacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridInformacion_KeyDown);
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// cmbCargoCarpa
			// 
			this.cmbCargoCarpa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargoCarpa.DataSource = this.ultraDataSource2;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 248;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbCargoCarpa.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCargoCarpa.DisplayMember = "Descripcion";
			this.cmbCargoCarpa.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargoCarpa.Location = new System.Drawing.Point(248, 144);
			this.cmbCargoCarpa.Name = "cmbCargoCarpa";
			this.cmbCargoCarpa.Size = new System.Drawing.Size(200, 112);
			this.cmbCargoCarpa.TabIndex = 961;
			this.cmbCargoCarpa.ValueMember = "idCargo";
			this.cmbCargoCarpa.Visible = false;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(584, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 960;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(496, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 959;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// ultraTabPageControl10
			// 
			this.ultraTabPageControl10.Controls.Add(this.btnCopiar);
			this.ultraTabPageControl10.Controls.Add(this.label1);
			this.ultraTabPageControl10.Controls.Add(this.dtHasta);
			this.ultraTabPageControl10.Controls.Add(this.button1);
			this.ultraTabPageControl10.Controls.Add(this.button2);
			this.ultraTabPageControl10.Controls.Add(this.cmbTramo);
			this.ultraTabPageControl10.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl10.Location = new System.Drawing.Point(2, 21);
			this.ultraTabPageControl10.Name = "ultraTabPageControl10";
			this.ultraTabPageControl10.Size = new System.Drawing.Size(652, 337);
			this.ultraTabPageControl10.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl10_Paint);
			// 
			// btnCopiar
			// 
			this.btnCopiar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCopiar.Image = ((System.Drawing.Image)(resources.GetObject("btnCopiar.Image")));
			this.btnCopiar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCopiar.Location = new System.Drawing.Point(416, 8);
			this.btnCopiar.Name = "btnCopiar";
			this.btnCopiar.TabIndex = 977;
			this.btnCopiar.Text = "Copiar";
			this.btnCopiar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCopiar.Click += new System.EventHandler(this.btnCopiar_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 966;
			this.label1.Text = "Periodo";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance14;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton2);
			this.dtHasta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.dtHasta.Format = "MM/yyyy ";
			this.dtHasta.Location = new System.Drawing.Point(72, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 965;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.ValueChanged += new System.EventHandler(this.dtHasta_ValueChanged);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.Enabled = false;
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(576, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(76, 23);
			this.button1.TabIndex = 964;
			this.button1.Text = "Guardar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.Enabled = false;
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(496, 8);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(72, 24);
			this.button2.TabIndex = 963;
			this.button2.Text = "Excel";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// cmbTramo
			// 
			this.cmbTramo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTramo.DataSource = this.ultraDataSource2;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 248;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbTramo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTramo.DisplayMember = "Descripcion";
			this.cmbTramo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbTramo.Location = new System.Drawing.Point(226, 176);
			this.cmbTramo.Name = "cmbTramo";
			this.cmbTramo.Size = new System.Drawing.Size(200, 48);
			this.cmbTramo.TabIndex = 962;
			this.cmbTramo.ValueMember = "idCargo";
			this.cmbTramo.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource3;
			appearance15.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance15;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn14.Header.Caption = "Tramo";
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn14.Width = 200;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn15.CellAppearance = appearance16;
			ultraGridColumn15.Header.Caption = "Desde";
			ultraGridColumn15.Header.VisiblePosition = 2;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance17;
			ultraGridColumn16.Header.Caption = "Hasta";
			ultraGridColumn16.Header.VisiblePosition = 3;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance18;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "% Desde";
			ultraGridColumn17.Header.VisiblePosition = 4;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance19;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "% Hasta";
			ultraGridColumn18.Header.VisiblePosition = 5;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance20;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 6;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance21;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			ultraGridBand5.Header.Caption = "Periodo A";
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance22;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance23;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance24;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance25;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand5.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand5.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(640, 280);
			this.ultraGrid1.TabIndex = 958;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16});
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
			// utcReportes
			// 
			this.utcReportes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcReportes.Appearance = appearance31;
			this.utcReportes.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcReportes.Controls.Add(this.TbBasico);
			this.utcReportes.Controls.Add(this.ultraTabPageControl10);
			this.utcReportes.Location = new System.Drawing.Point(16, 8);
			this.utcReportes.Name = "utcReportes";
			this.utcReportes.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcReportes.Size = new System.Drawing.Size(656, 360);
			this.utcReportes.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.PropertyPage2003;
			this.utcReportes.TabButtonStyle = Infragistics.Win.UIElementButtonStyle.OfficeXPToolbarButton;
			this.utcReportes.TabIndex = 962;
			ultraTab1.TabPage = this.TbBasico;
			ultraTab1.Text = "Porcentaje Comisiones";
			ultraTab2.TabPage = this.ultraTabPageControl10;
			ultraTab2.Text = "Penalización";
			this.utcReportes.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													ultraTab1,
																																													ultraTab2});
			this.utcReportes.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(652, 337);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(688, 8);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 963;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCobPorcentajeComisiones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(768, 374);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.utcReportes);
			this.KeyPreview = true;
			this.Name = "frmCobPorcentajeComisiones";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Porcentaje Comisiones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCobPorcentajeComisiones_KeyDown);
			this.Load += new System.EventHandler(this.frmCobPorcentajeComisiones_Load);
			this.TbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargoCarpa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ultraTabPageControl10.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTramo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcReportes)).EndInit();
			this.utcReportes.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCobPorcentajeComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private Acceso miAcceso;	
		private void UnloadMe()
		{
			this.Close();
		}

		private void frmCobPorcentajeComisiones_Load(object sender, System.EventArgs e)
		{
			
			miAcceso = new Acceso(cdsSeteoF, "083603");
			if (!Funcion.Permiso("1191", cdsSeteoF))
			{		
				MessageBox.Show("No tiene Acceso a Porcentaje Comisiones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;
			if (miAcceso.Imprimir) this.btnGuardar.Enabled =  true;
			if (miAcceso.BExportar) this.button2.Enabled =  true;
			if (miAcceso.Imprimir) this.button1.Enabled =  true;

			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");
			this.dtDesde.CalendarInfo.MaxDate = DateTime.Now.AddDays(+5 );		
			this.dtDesde.Value = DateTime.Today ;

			this.dtHasta.CalendarInfo.MinDate = DateTime.Parse("01/03/2019");
			this.dtHasta.CalendarInfo.MaxDate = DateTime.Now.AddDays(+5 );		
			this.dtHasta.Value = DateTime.Today ;
	
			this.cmbCargoCarpa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select Tramo AS idCargo, Descripcion from TramosOperadoras");
			this.cmbTramo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "select Tramo AS idCargo, Descripcion from TramosOperadoras");

			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 0, '{0}'",Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd") ));
//				this.cmbDDBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, '{0}'", (int) this.cmbDDParroquia.Value));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 1, '{0}' ",Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")) );
		}

		private void uGridInformacion_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCobPorcentajeComisiones"].Value = 0;
			e.Row.Cells["idCargo"].Value = 0;
			e.Row.Cells["Hasta"].Value = 0.00m;
			e.Row.Cells["Desde"].Value = 0.00m;
			e.Row.Cells["Porcentaje"].Value = 0.00m;
			e.Row.Cells["Valor"].Value = 0.00m;

		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
		FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void uGridInformacion_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["idCobPorcentajeComisiones"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["idCobPorcentajeComisiones"].Value > 0)
						{
							string sSQL = string.Format("Delete From Cob_PorcentajeComisiones Where idCobPorcentajeComisiones = {0}", (int)e.Rows[i].Cells["idCobPorcentajeComisiones"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
					}			
				}
			}
			e.DisplayPromptMsg = false;
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione El Periodo  ", false, "", cdsSeteoF)) return;

			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
			{
				#region Porcentajes Desde C
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
					
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["Desde"].Value) > Convert.ToDecimal(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridInformacion.ActiveRow = dr;
					this.uGridInformacion.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				#region Porcentajes repetidos
//				string sSQL = string.Format("Exec Cob_ValidaPorcentajes {0}, {1}, {2}, {3}", 
//					(int)dr.Cells["idCobPorcentajeComisiones"].Value, (int)dr.Cells["idCargo"].Value, Convert.ToDecimal(dr.Cells["Desde"].Value), Convert.ToDecimal(dr.Cells["Hasta"].Value));
//
//				if (Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
//				{
//					MessageBox.Show("Esta ingresando valores repetidos en Porcentajes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.uGridInformacion.ActiveRow = dr;
//					this.uGridInformacion.ActiveRow.Selected = true;
//
//					return;
//				}
				#endregion Porcentajes repetidos				
		
			}
				
			#endregion Valida Comisiones

			#region Guardar
			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
				{
					string sSQLDetalle = string.Format("Exec GuardaCob_PorcentajeComisiones {0}, {1}, {2}, {3}, {4}, {5}, '{6}'",
						(int)dr.Cells["idCobPorcentajeComisiones"].Value, 
						(int)dr.Cells["idCargo"].Value, 
						(decimal)dr.Cells["Desde"].Value, 
						(decimal)dr.Cells["Hasta"].Value, 
						(decimal)dr.Cells["Porcentaje"].Value, 
						(decimal)dr.Cells["Valor"].Value,
						Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}
			}
			#endregion guardar


			this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 0, '{0}'",Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd") ));
	

		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Desde", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Hasta", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Porcentaje", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void uGridInformacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridInformacion);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione El Periodo  ", false, "", cdsSeteoF)) return;

			#region Valida Comisiones
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				#region Porcentajes Desde C
				if (Convert.ToDecimal(dr.Cells["PorcDesde"].Value) > Convert.ToDecimal(dr.Cells["PorcentanjeHasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final
					
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Porcentajes Inicial Final
				if (Convert.ToDecimal(dr.Cells["DiasDesde"].Value) > Convert.ToDecimal(dr.Cells["DiasHasta"].Value))
				{
					MessageBox.Show("En Porcentajes, el valor Inicial no puede ser mayor al valor Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				if (Convert.ToInt32(dr.Cells["DiasDesde"].Value) == 0)
				{
					MessageBox.Show("Ingrese el Dia Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				if (Convert.ToInt32(dr.Cells["DiasHasta"].Value) == 0)
				{
					MessageBox.Show("Ingrese el Dia Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["DiasDesde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Dia Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				if (dr.Cells["DiasHasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Dia Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

				if (dr.Cells["Valor"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;

					return;
				}

//				if ((int)dr.Cells["Valor"].Value == 0)
//				{
//					MessageBox.Show("Ingreso el Valor ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.ultraGrid1.ActiveRow = dr;
//					this.ultraGrid1.ActiveRow.Selected = true;
//
//					return;
//				}
				
		
			}
				
			#endregion Valida Comisiones

			#region Guardar
			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				{
					string sSQLDetalle = string.Format("Exec [GuardaCob_PorcentajeComisionesPenalizacion] {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}'",
						(int)dr.Cells["iCob_Penalizacion"].Value, 
						(int)dr.Cells["idCargo"].Value, 
						(int)dr.Cells["DiasDesde"].Value, 
						(int)dr.Cells["DiasHasta"].Value, 
						(decimal)dr.Cells["PorcDesde"].Value, 
						(decimal)dr.Cells["PorcentanjeHasta"].Value,
						(decimal)dr.Cells["Valor"].Value,
						(decimal)dr.Cells["Penalizacion"].Value,
						Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}
			}
			#endregion guardar

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 1, '{0}' ",Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")) );
	

		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["iCob_Penalizacion"].Value = 0;
			e.Row.Cells["idCargo"].Value = 0;
			e.Row.Cells["DiasDesde"].Value = 0;
			e.Row.Cells["DiasHasta"].Value = 0;
			e.Row.Cells["PorcDesde"].Value = 0;
			e.Row.Cells["PorcentanjeHasta"].Value = 0;
			e.Row.Cells["Valor"].Value = 0;
			e.Row.Cells["Penalizacion"].Value = 0;



		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))				
			{
				for (int i = 0; i < e.Rows.Length; i++)
				{
					if (e.Rows[i].Cells["iCob_Penalizacion"].Value != System.DBNull.Value)
					{									
						if ((int)e.Rows[i].Cells["iCob_Penalizacion"].Value > 0)
						{
							string sSQL = string.Format("Delete From Cob_Penalizacion Where iCob_Penalizacion = {0}", (int)e.Rows[i].Cells["iCob_Penalizacion"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						}
					}			
				}
			}
			e.DisplayPromptMsg = false;
		}

		private void dtDesde_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 0, '{0}' ",Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd")) );
			}
		}

		private void dtHasta_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtDesde.Value != System.DBNull.Value)
			{
				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 1, '{0}'",Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd") ));
			}
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
			if (!Validacion.vbFechaEnDocumentos(this.dtDesde, "Seleccione El Periodo que desea Actualizar  ", false, "", cdsSeteoF)) return;

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select Top 1  convert (varchar ( 20), DatePart ( year, Periodo))+'/'+ convert (varchar ( 20), DatePart ( MONTH, Periodo)) From Cob_PorcentajeComisiones Order By idCobPorcentajeComisiones Desc  ");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [CopiaPorcentajeComisiones] '{0}'", Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd")));
				//				this.uGridcobradores.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
				if ( iRetorna == 1)
				{

					MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 0, '{0}'",Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd") ));
					//		
				}
				else
				{
					MessageBox.Show(string.Format("No existen Datos en el Mes anterior {0} o ya se genero una copia en el mes actual ", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.uGridInformacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 0, '{0}'",Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd") ));
					//		

				}
			}
		}

		private void btnCopiar_Click(object sender, System.EventArgs e)
		{
			int iRetorna = 0;
			if (!Validacion.vbFechaEnDocumentos(this.dtHasta, "Seleccione El Periodo que desea Actualizar  ", false, "", cdsSeteoF)) return;

			string sPeriodo = Funcion.sEscalarSQL(cdsSeteoF, "Select Top 1  convert (varchar ( 20), DatePart ( year, Periodo))+'/'+ convert (varchar ( 20), DatePart ( MONTH, Periodo)) From Cob_Penalizacion Order By iCob_Penalizacion Desc  ");

			if (DialogResult.Yes == MessageBox.Show("¿Esta Seguro De Copiar La Información del anterior Mes?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{	

				iRetorna = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [CopiaPorcentajePenalización] '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));
				//				this.uGridcobradores.DataSource = FuncionesProcedimientos.dtGeneral(sSQLPa);
				if ( iRetorna == 1)
				{

					MessageBox.Show(string.Format("Se copio la Información del Periodo {0} correctamente", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 1, '{0}' ",Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")) );
					//		
				}
				else
				{
					MessageBox.Show(string.Format("No existen Datos en el Mes anterior {0} o ya se genero una copia en el mes actual ", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cob_ConsultaPorcentajesComisiones 1, '{0}' ",Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")) );
					//		

				}
			}
		}

		private void ultraTabPageControl10_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}
	}
}
