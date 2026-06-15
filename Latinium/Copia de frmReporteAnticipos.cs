using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteAnticipos.
	/// </summary>
	public class frmReporteAnticipos : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSuma;
		private Infragistics.Win.Misc.UltraLabel lblContador;
		private Infragistics.Win.Misc.UltraLabel lblSeleccionados;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProcesados;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btProcesar;
		private System.Windows.Forms.Button btBuscar;
		private C1DataRow drSeteoF;

		public frmReporteAnticipos()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCruceAnticiposSolidario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAnticipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Temporal");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 7, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 8, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Pagado", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Pagado", 9, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 10, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 10, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCruceAnticiposSolidario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Detalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Temporal");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraLabel3 = new Infragistics.Win.Misc.UltraLabel();
			this.txtBuscador = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtSuma = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblContador = new Infragistics.Win.Misc.UltraLabel();
			this.lblSeleccionados = new Infragistics.Win.Misc.UltraLabel();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.chkProcesados = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btProcesar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btBuscar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSuma)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
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
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 111;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 65;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.Caption = "Local";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 50;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 90;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 250;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 250;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance2;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.PromptChar = ' ';
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance3;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 75;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 75;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 75;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 87;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.Header.Caption = "Cuenta";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 100;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn14.Header.Caption = "Procesado";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 75;
			ultraGridColumn15.Header.Caption = "...";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 26;
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
																										 ultraGridColumn15});
			ultraGridBand1.Header.Caption = "";
			ultraGridBand1.HeaderVisible = true;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance6;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance7;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance8;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.Hidden = true;
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance9;
			summarySettings4.DisplayFormat = "{0: #,##0.00}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.RowAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(7, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1152, 272);
			this.ultraGrid1.TabIndex = 6;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
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
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(System.DateTime);
			ultraDataColumn14.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
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
																																 ultraDataColumn15});
			// 
			// ultraLabel3
			// 
			this.ultraLabel3.AutoSize = true;
			this.ultraLabel3.Location = new System.Drawing.Point(8, 11);
			this.ultraLabel3.Name = "ultraLabel3";
			this.ultraLabel3.Size = new System.Drawing.Size(39, 14);
			this.ultraLabel3.TabIndex = 159;
			this.ultraLabel3.Text = "Buscar";
			// 
			// txtBuscador
			// 
			this.txtBuscador.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscador.Location = new System.Drawing.Point(72, 8);
			this.txtBuscador.Name = "txtBuscador";
			this.txtBuscador.Size = new System.Drawing.Size(304, 21);
			this.txtBuscador.TabIndex = 0;
			this.txtBuscador.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtBuscador_KeyPress);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(990, 344);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 233;
			this.label1.Text = "TOTAL";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSuma
			// 
			this.txtSuma.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSuma.Appearance = appearance16;
			this.txtSuma.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSuma.Enabled = false;
			this.txtSuma.FormatString = "#,##0.00";
			this.txtSuma.Location = new System.Drawing.Point(1046, 342);
			this.txtSuma.Name = "txtSuma";
			this.txtSuma.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSuma.PromptChar = ' ';
			this.txtSuma.Size = new System.Drawing.Size(112, 21);
			this.txtSuma.TabIndex = 234;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Location = new System.Drawing.Point(560, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 0);
			this.lblContador.TabIndex = 235;
			// 
			// lblSeleccionados
			// 
			this.lblSeleccionados.AutoSize = true;
			this.lblSeleccionados.Location = new System.Drawing.Point(832, 8);
			this.lblSeleccionados.Name = "lblSeleccionados";
			this.lblSeleccionados.Size = new System.Drawing.Size(0, 0);
			this.lblSeleccionados.TabIndex = 236;
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
			// chkProcesados
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProcesados.Appearance = appearance17;
			this.chkProcesados.BackColor = System.Drawing.Color.Transparent;
			this.chkProcesados.Location = new System.Drawing.Point(512, 7);
			this.chkProcesados.Name = "chkProcesados";
			this.chkProcesados.Size = new System.Drawing.Size(88, 22);
			this.chkProcesados.TabIndex = 2;
			this.chkProcesados.Text = "Procesados";
			this.chkProcesados.CheckedChanged += new System.EventHandler(this.chkProcesados_CheckedChanged);
			// 
			// dtFecha
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance18;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(656, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 3;
			this.dtFecha.Value = new System.DateTime(2014, 3, 22, 0, 0, 0, 0);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(608, 10);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(36, 16);
			this.label19.TabIndex = 237;
			this.label19.Text = "Fecha";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1174, 7);
			this.groupBox1.TabIndex = 240;
			this.groupBox1.TabStop = false;
			// 
			// btProcesar
			// 
			this.btProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btProcesar.CausesValidation = false;
			this.btProcesar.Location = new System.Drawing.Point(8, 336);
			this.btProcesar.Name = "btProcesar";
			this.btProcesar.Size = new System.Drawing.Size(80, 23);
			this.btProcesar.TabIndex = 285;
			this.btProcesar.Text = "Procesar";
			this.btProcesar.Click += new System.EventHandler(this.btProcesar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnSalir.Location = new System.Drawing.Point(96, 336);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 23);
			this.btnSalir.TabIndex = 363;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btBuscar
			// 
			this.btBuscar.CausesValidation = false;
			this.btBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btBuscar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btBuscar.Location = new System.Drawing.Point(408, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 23);
			this.btBuscar.TabIndex = 364;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// frmReporteAnticipos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1166, 366);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btProcesar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkProcesados);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSeleccionados);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtSuma);
			this.Controls.Add(this.ultraLabel3);
			this.Controls.Add(this.txtBuscador);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "frmReporteAnticipos";
			this.Text = "Captura de Anticipos Solidario";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteAnticipos_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmReporteAnticipos_Closing);
			this.Load += new System.EventHandler(this.frmReporteAnticipos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSuma)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void txtBuscador_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.btBuscar_Click(sender, e);
		}

		private void Suma()
		{
			this.txtSuma.Value = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Temporal"].Value) this.txtSuma.Value = Convert.ToDecimal(this.txtSuma.Value) + Convert.ToDecimal(dr.Cells["Valor"].Value);
			}
		}

		private void frmReporteAnticipos_Load(object sender, System.EventArgs e)
		{
			this.txtBuscador.Focus();

			this.dtFecha.Value = DateTime.Today;
		}

		private void ultraGrid1_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Temporal")
			{
				this.ultraGrid1.UpdateData();
					
				int bTemp = 0; /* ESTA COMO ENTERO NO RECONOCE COMO BIT AL CAMPO TEMPORAL */
				if ((bool)e.Cell.Row.Cells["Temporal"].Value) bTemp = 1;
				string sSQLAct = string.Format("Update Anticipos Set Temporal = {1} Where idAnticipo = {0}", (int)e.Cell.Row.Cells["idAnticipo"].Value, bTemp);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);

				if (bTemp == 1)
				{
					string sSQL = string.Format("Update Anticipos Set Valor = {0} Where idAnticipo = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value), (int)e.Cell.Row.Cells["idAnticipo"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
				}

				Suma();
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["Valor"].Format = stFormato; 
			e.Layout.Bands[0].Columns["Valor"].MaskInput = stInput;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private bool ValidaValor(Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Row.Cells["idAnticipo"].Value == DBNull.Value) return false;
			
			if (Convert.ToDecimal(e.Cell.Value) > Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value))
			{
				e.Cell.Value = Convert.ToDecimal(e.Cell.OriginalValue);
				MessageBox.Show(string.Format("No puede ser mayor a {0}.", Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
		
			return true;
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Valor")
			{
				//if (!ValidaValor(e)) return;

				if (e.Cell.Row.Cells["idAnticipo"].Value == DBNull.Value) return;
			
				if (Convert.ToDecimal(e.Cell.Value) > Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value)) 
				{
					e.Cell.Value = Convert.ToDecimal(e.Cell.OriginalValue);
					MessageBox.Show(string.Format("No puede ser mayor a {0}.", Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				
				string sSQL = string.Format("Update Anticipos Set Valor = {0} Where idAnticipo = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Valor"].Value), (int)e.Cell.Row.Cells["idAnticipo"].Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

				Suma();
			}

			if (e.Cell.Column.ToString() == "Temporal")
			{
//				int bTemp = 0; /* ESTA COMO ENTERO NO RECONOCE COMO BIT AL CAMPO TEMPORAL */
//				if ((bool)e.Cell.Row.Cells["Temporal"].Value) bTemp = 1;
//				string sSQLAct = string.Format("Update Anticipos Set Temporal = {1} Where idAnticipo = {0}", (int)e.Cell.Row.Cells["idAnticipo"].Value, bTemp);
//				Funcion.EjecutaSQL(cdsSeteoF, sSQLAct, true);
//
//				if (bTemp == 1)
//				{
//					string sSQL = string.Format("Update Anticipos Set Valor = {0} Where idAnticipo = {1}", Convert.ToDecimal(e.Cell.Row.Cells["Saldo"].Value, (int)e.Cell.Row.Cells["idAnticipo"].Value);
//					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
//				}
//
//				Suma();
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{				
				switch(e.KeyValue)
				{		
					#region F3
					case (int) Keys.F3:
					#region Cuenta contable
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "DESCRIPCION")
						{															
							using (BuscaCuentas miBusqueda = new BuscaCuentas())
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									this.ultraGrid1.ActiveRow.Cells["idCuenta"].Value = (int)miBusqueda.ultraGrid1.ActiveRow.Cells["idCuenta"].Value;
									this.ultraGrid1.ActiveRow.Cells["Descripcion"].Value = miBusqueda.ultraGrid1.ActiveRow.Cells["Descripcion"].Value.ToString();
								}
							}
						}
						#endregion Cuenta contable						
					break;
					#endregion F3						
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
						if(ultraGrid1.ActiveCell.DroppedDown == false) ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
						#endregion Enter
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Tecla");
			}
		}

		private void frmReporteAnticipos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			
		}

		private void frmReporteAnticipos_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (Funcion.iEscalarSQL(cdsSeteoF, "Select Count(*) From Anticipos Where Temporal = 1 And Capturado = 0 And Tipo = 5") > 0)
			{
				if (DialogResult.No == MessageBox.Show(string.Format("Ha seleccionado anticipos para procesar. ¿Esta seguro de salir?."), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					e.Cancel = true;
				}
				else
				{
					Funcion.EjecutaSQL(cdsSeteoF, "Update Anticipos Set Temporal = 0 Where Temporal = 1 And Tipo = 5");
				}
			}
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{			
			if ((bool)e.Row.Cells["Temporal"].Value)
			{
				e.Row.Appearance.BackColor = Color.LightGreen;
				e.Row.Appearance.BackColor2 = Color.LightGreen;	
			}
			else if (!(bool)e.Row.Cells["Temporal"].Value)
			{
				e.Row.Appearance.BackColor = Color.White;
				e.Row.Appearance.BackColor2 = Color.White;	
			}
		}

		private void chkProcesados_CheckedChanged(object sender, System.EventArgs e)
		{
			if ((bool)this.chkProcesados.Checked) this.btProcesar.Enabled = false;
			else this.btProcesar.Enabled = true;
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btProcesar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ReporteAnticipos '', 0, '{0}'", dtFecha.ToString("yyyyMMdd")));

			int iContador = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool)dr.Cells["Temporal"].Value) iContador = iContador + 1;
			}

			if (iContador == 0)
			{
				MessageBox.Show("No ha seleccionado ningun registro para procesar.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscador.Focus();
				return;
			}

			if (Convert.ToDecimal(this.txtSuma.Value) == 0.00m)
			{
				MessageBox.Show("No puede procesar en 0.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtBuscador.Focus();
				return;
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("¿Esta seguro de procesar la suma de {0}.?", Convert.ToDecimal(this.txtSuma.Value)), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{	
				using (frmIngresoBeneficiario miPrestamo = new frmIngresoBeneficiario (8, Convert.ToDecimal(this.txtSuma.Value), 0))
				{
					if (DialogResult.OK == miPrestamo.ShowDialog())
					{	
						this.Cursor = Cursors.WaitCursor;
		
						#region Autorización de Pagos
						decimal dValor = Convert.ToDecimal(this.txtSuma.Value) + Convert.ToDecimal(miPrestamo.txtDiferencia.Value);

						string sSQL = string.Format("Exec CreaAutorizacionesDePago 8, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'CRUCE ANTICIPOS', 0", 
							(int)miPrestamo.txtIdRolPrestamo.Value, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"), dValor);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
						#endregion Autorización de Pagos

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
						{
							int idCuenta = 0;
							if (dr.Cells["idCuenta"].Value != System.DBNull.Value) idCuenta = (int)dr.Cells["idCuenta"].Value;

							if ((bool)dr.Cells["Temporal"].Value)
							{
								/*string sSQLAsiento = string.Format("Exec AsientoReposicionSolidario 0, '', 19, '{0}', '{1}', {2}", dr.Cells["Nombre"].Value.ToString(), "");
								int IdAsiento = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento);*/

								string sSQLCruce = string.Format("Exec GCruceAnticipos {0}, {1}, {2}, {3}, {4}, ''", 
									(int)dr.Cells["idCruceAnticiposSolidario"].Value, (int)dr.Cells["idAnticipo"].Value, miPrestamo.txtIdRolPrestamo.Value, idCuenta, Convert.ToDecimal(dr.Cells["Valor"].Value));
								Funcion.EjecutaSQL(cdsSeteoF, sSQLCruce, true);

								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec SaldoAnticipoSolidario {0}", (int)dr.Cells["idAnticipo"].Value));	

								Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Anticipos Set Temporal = 0 Where idAnticipo = {0}", (int)dr.Cells["idAnticipo"].Value));
							}
						}

						this.Cursor = Cursors.Default;

						MessageBox.Show(string.Format("Se Proceso la suma de {0}.", Convert.ToDecimal(this.txtSuma.Value)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						
						dtFecha = (DateTime)this.dtFecha.Value;

						this.txtSuma.Value = 0.00m;

						this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ReporteAnticipos '', 0, '{0}'", dtFecha.ToString("yyyyMMdd")));
					}
				}
			}
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = "";
			DateTime dtFecha = (DateTime)this.dtFecha.Value;

			this.ultraGrid1.DataSource = ultraDataSource1;

			string sSQL = string.Format("Exec ReporteAnticipos '{0}', {1}, '{2}'", this.txtBuscador.Text.ToString(), (bool)this.chkProcesados.Checked, dtFecha.ToString("yyyyMMdd"));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			this.ultraGrid1.DisplayLayout.Bands[0].Header.Caption = this.ultraGrid1.Rows.Count.ToString() + " Registros encontrados.";

			Suma();
		}
	}
}
