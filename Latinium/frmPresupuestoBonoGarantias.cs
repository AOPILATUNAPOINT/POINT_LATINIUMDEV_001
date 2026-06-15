using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPresupuestoBonoGarantias.
	/// </summary>
	public class frmPresupuestoBonoGarantias : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDepositos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCargos;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbFormasPago;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPresupuestoBonoGarantias()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPresupuestoBonoGarantias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo", -1, "cmbCargos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dGex");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCP");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dPM");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dPJ");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dPro");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dPR");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dSDT");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago", -1, "cmbFormasPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dSDF");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dGex", 2, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dGex", 2, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dCP", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dCP", 3, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dPM", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dPM", 4, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dPJ", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dPJ", 5, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dPro", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dPro", 8, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dPR", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dPR", 9, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dSDT", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dSDT", 10, true);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "dSDF", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "dSDF", 12, true);
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPresupuestoBonoGarantias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dGex");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dCP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dPM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dPJ");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dPro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dPR");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dSDT");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("dSDF");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPresupuestoBonoGarantias));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			this.uGridDepositos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCargos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.btnExportar = new System.Windows.Forms.Button();
			this.cmbFormasPago = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormasPago)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDepositos
			// 
			this.uGridDepositos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDepositos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDepositos.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridDepositos.DisplayLayout.Appearance = appearance1;
			this.uGridDepositos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 127;
			ultraGridColumn2.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn2.Header.Caption = "Cargo";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn2.Width = 99;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance2;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.Caption = "Garantia Extendida";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 66;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "Compra Protegida";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 68;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.Caption = "Point Movil";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 69;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Pay Joy";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 68;
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Width = 66;
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.Width = 66;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance6;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Proyector";
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 63;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance7;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Point Respaldo";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 67;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance8;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "Salud D. Tit.";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 67;
			ultraGridColumn12.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn12.Header.Caption = "Tipo Precio";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn12.Width = 101;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance9;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Salud D. Fam.";
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Width = 67;
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
																										 ultraGridColumn13});
			ultraGridBand1.Header.Caption = "Intangible";
			ultraGridBand1.HeaderVisible = true;
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance10;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance11;
			summarySettings2.DisplayFormat = "{0:#,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance12;
			summarySettings3.DisplayFormat = "{0:#,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance13;
			summarySettings4.DisplayFormat = "{0:#,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance14;
			summarySettings5.DisplayFormat = "{0:#,##0.00}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings6.Appearance = appearance15;
			summarySettings6.DisplayFormat = "{0:#,##0.00}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings7.Appearance = appearance16;
			summarySettings7.DisplayFormat = "{0:#,##0.00}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings8.Appearance = appearance17;
			summarySettings8.DisplayFormat = "{0:#,##0.00}";
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
			ultraGridBand1.SummaryFooterCaption = "Total";
			this.uGridDepositos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridDepositos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.uGridDepositos.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDepositos.DisplayLayout.Override.CellAppearance = appearance19;
			this.uGridDepositos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance20.FontData.BoldAsString = "True";
			appearance20.FontData.Name = "Arial";
			appearance20.FontData.SizeInPoints = 8F;
			appearance20.ForeColor = System.Drawing.Color.White;
			appearance20.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDepositos.DisplayLayout.Override.HeaderAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDepositos.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.uGridDepositos.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDepositos.Location = new System.Drawing.Point(8, 56);
			this.uGridDepositos.Name = "uGridDepositos";
			this.uGridDepositos.Size = new System.Drawing.Size(888, 296);
			this.uGridDepositos.TabIndex = 252;
			this.uGridDepositos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDepositos_AfterRowInsert);
			this.uGridDepositos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDepositos_BeforeRowsDeleted);
			this.uGridDepositos.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDepositos_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.Decimal);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn13});
			// 
			// cmbCargos
			// 
			this.cmbCargos.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbCargos.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCargos.DisplayMember = "Cargo";
			this.cmbCargos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbCargos.Location = new System.Drawing.Point(48, 176);
			this.cmbCargos.Name = "cmbCargos";
			this.cmbCargos.Size = new System.Drawing.Size(200, 112);
			this.cmbCargos.TabIndex = 768;
			this.cmbCargos.ValueMember = "idCargo";
			this.cmbCargos.Visible = false;
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
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(912, 8);
			this.groupBox1.TabIndex = 773;
			this.groupBox1.TabStop = false;
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button1.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(752, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(76, 23);
			this.button1.TabIndex = 771;
			this.button1.Text = "Guardar";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.button2.CausesValidation = false;
			this.button2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.button2.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
			this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button2.Location = new System.Drawing.Point(832, 8);
			this.button2.Name = "button2";
			this.button2.TabIndex = 772;
			this.button2.Text = "Salir";
			this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(672, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 24);
			this.btnExportar.TabIndex = 774;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Visible = false;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// cmbFormasPago
			// 
			this.cmbFormasPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 200;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbFormasPago.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbFormasPago.DisplayMember = "FormaPago";
			this.cmbFormasPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
			this.cmbFormasPago.Location = new System.Drawing.Point(416, 176);
			this.cmbFormasPago.Name = "cmbFormasPago";
			this.cmbFormasPago.Size = new System.Drawing.Size(200, 112);
			this.cmbFormasPago.TabIndex = 775;
			this.cmbFormasPago.ValueMember = "idFormaPago";
			this.cmbFormasPago.Visible = false;
			// 
			// frmPresupuestoBonoGarantias
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(912, 358);
			this.Controls.Add(this.cmbFormasPago);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.cmbCargos);
			this.Controls.Add(this.uGridDepositos);
			this.KeyPreview = true;
			this.Name = "frmPresupuestoBonoGarantias";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Configuración Intangible";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPresupuestoBonoGarantias_KeyDown);
			this.Load += new System.EventHandler(this.frmPresupuestoBonoGarantias_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDepositos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormasPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPresupuestoBonoGarantias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
	 this.Close();		
		}

		private void uGridDepositos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idPresupuestoBonoGarantias"].Value = 0;
			e.Row.Cells["idCargo"].Value = 0.00m;
			e.Row.Cells["dGex"].Value = 0.00m;			
			e.Row.Cells["dCP"].Value = 0.00m;
			e.Row.Cells["dPM"].Value = 0.00m;			
			e.Row.Cells["dPJ"].Value = 0.00m;
			e.Row.Cells["dPro"].Value = 0.00m;
			e.Row.Cells["dPR"].Value = 0.00m;
			e.Row.Cells["dSDT"].Value = 0.00m;
			e.Row.Cells["dSDF"].Value = 0.00m;
			e.Row.Cells["idFormaPago"].Value = 0.00m;
		}

		private void frmPresupuestoBonoGarantias_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "052114");

			if (!Funcion.Permiso("1750", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Configuración de Intangible", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.BExportar) this.btnExportar.Visible = true;
			

			this.cmbCargos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "	 Select idCargo, Cargo From dbo.Com_CargosDeVentas Where Comisiona = 1 and idCargo in (1,3) Order By Orden");
			this.cmbFormasPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, 
				@"Select idFormaPago, FormaPago From dbo.CompraFormaPago Where Activo = 1 and idFormaPago in (1,3) 
				union
				Select 0 as idFormaPago, 'TODOS' as FormaPago
				Order By idFormaPago");
	    Consulta ();
		}
		private void Consulta()
		{
			this.uGridDepositos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaPresupuestoBonoGarantias");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}

		bool bActualiza = false;
		private void uGridDepositos_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bActualiza) return;

			if (e.Cell.Column.ToString() == "dGex")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dGex"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dGex"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dGex"].Value < 0 )
				{					
					e.Cell.Row.Cells["dGex"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}


			if (e.Cell.Column.ToString() == "dCP")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dCP"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dCP"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dCP"].Value < 0 )
				{					
					e.Cell.Row.Cells["dCP"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}


			if (e.Cell.Column.ToString() == "dPM")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dPM"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dPM"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dPM"].Value < 0 )
				{					
					e.Cell.Row.Cells["dPM"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "dPJ")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dPJ"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dPJ"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dPJ"].Value < 0 )
				{					
					e.Cell.Row.Cells["dPJ"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "dPro")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dPro"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dPro"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dPro"].Value < 0 )
				{					
					e.Cell.Row.Cells["dPro"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "dPR")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dPR"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dPR"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dPR"].Value < 0 )
				{					
					e.Cell.Row.Cells["dPR"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "dSDT")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dSDT"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dSDT"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dSDT"].Value < 0 )
				{					
					e.Cell.Row.Cells["dSDT"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}

			if (e.Cell.Column.ToString() == "dSDF")
			{	
				bActualiza = true;

				if (e.Cell.Row.Cells["dSDF"].Value == System.DBNull.Value)
				{					
					e.Cell.Row.Cells["dSDF"].Value = (decimal)e.Cell.OriginalValue;	
				}	
				if ((decimal ) e.Cell.Row.Cells["dSDF"].Value < 0 )
				{					
					e.Cell.Row.Cells["dSDF"].Value = (decimal)e.Cell.OriginalValue;	
				}
				bActualiza = false;
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
			{
				#region Cargo
				if (dr.Cells["idCargo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}

				if ((int)dr.Cells["idCargo"].Value == 0)
				{
					MessageBox.Show("Seleccione un Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion Cargo

				#region Forma Pago
				if (dr.Cells["idFormaPago"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione una Forma de Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}

				#endregion Forma Pago


			
				if (dr.Cells["Desde"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un la Fecha inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}

				if (dr.Cells["Hasta"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione un la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}

				
				#region Porcentajes Inicial Final
				if (Convert.ToDateTime(dr.Cells["Desde"].Value) > Convert.ToDateTime(dr.Cells["Hasta"].Value))
				{
					MessageBox.Show("La fecha inicial no puede ser mayor a la fecha final.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

					this.uGridDepositos.ActiveRow = dr;
					this.uGridDepositos.ActiveRow.Selected = true;

					return;
				}
				#endregion Porcentajes Inicial Final

			}

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de Guardar los cambios", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDepositos.Rows.All)
				{
					string sSQLDetalle = string.Format("Exec [GuardaPresupuestoBonoGarantias] {0}, {1}, {2}, {3}, {4}, {5}, '{6}', '{7}', {8}, {9}, {10}, {11}, {12}",
						(int)dr.Cells["idPresupuestoBonoGarantias"].Value, 
						(int)dr.Cells["idCargo"].Value,
						(decimal)dr.Cells["dGex"].Value, 
						(decimal)dr.Cells["dCP"].Value,
						(decimal)dr.Cells["dPM"].Value, 
						(decimal)dr.Cells["dPJ"].Value,
						Convert.ToDateTime(dr.Cells["Desde"].Value).ToString("yyyyMMdd"), 
						Convert.ToDateTime(dr.Cells["Hasta"].Value).ToString("yyyyMMdd"),
						(decimal)dr.Cells["dPro"].Value,
						(decimal)dr.Cells["dPR"].Value,
						(decimal)dr.Cells["dSDT"].Value,
						(decimal)dr.Cells["dSDF"].Value,
						(int)dr.Cells["idFormaPago"].Value);

					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}

				Consulta();
			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			 this.Close();	
		}

		private void uGridDepositos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			#region Elimina

			if (miAcceso.Eliminar)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Eliminar el registro?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{				
					for (int i=0; i < e.Rows.Length; i++)
					{	
						if ((int)e.Rows[i].Cells["idPresupuestoBonoGarantias"].Value > 0)
						{
							string sSQL = string.Format("Delete From PresupuestoBonoGarantias where idPresupuestoBonoGarantias = {0} ", (int)e.Rows[i].Cells["idPresupuestoBonoGarantias"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}					
				}
				else e.Cancel = true;
				#endregion Elimina
			}
			else
			{
				e.DisplayPromptMsg = false;
				e.Cancel = true;
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridDepositos);
		}
	}
}
