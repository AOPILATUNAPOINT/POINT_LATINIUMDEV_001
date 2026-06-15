using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmFondoGarantiaPichincha.
	/// </summary>
	public class frmFondoGarantiaPichincha : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btBuscar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private System.Windows.Forms.Button btnContabilizar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Button btnImportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFondoGarantiaPichincha()
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
		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFondoGarantiaPichincha));
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroLote");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeOperacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Vendidas");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapitalVendido");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias_Mora");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDebito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDebito");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Monto", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Monto", 6, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CapitalVendido", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CapitalVendido", 11, true);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 12, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorDebito", 15, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorDebito", 15, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroLote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeOperacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Vendidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CapitalVendido");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Dias_Mora");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDebito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDebito");
			this.label2 = new System.Windows.Forms.Label();
			this.btBuscar = new System.Windows.Forms.Button();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.btnImportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(81, 17);
			this.label2.TabIndex = 639;
			this.label2.Text = "Fecha De Corte";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btBuscar
			// 
			this.btBuscar.CausesValidation = false;
			this.btBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btBuscar.Image")));
			this.btBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btBuscar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btBuscar.Location = new System.Drawing.Point(248, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(88, 23);
			this.btBuscar.TabIndex = 640;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(112, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(120, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 638;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// uGridLotes
			// 
			this.uGridLotes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridLotes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridLotes.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridLotes.DisplayLayout.Appearance = appearance1;
			this.uGridLotes.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "SeteoF";
			ultraGridColumn1.Header.Caption = "...";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 20;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 40;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 65;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 65;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Cédula";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 75;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 250;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 60;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Header.Caption = "Lote";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 40;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Operacion";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 70;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.Caption = "Venta";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 65;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance4;
			ultraGridColumn11.Header.Caption = "C Vendidas";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 69;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance5;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Cap Vendido";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 80;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 60;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Header.Caption = "D. Mora";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 50;
			ultraGridColumn15.Header.Caption = "F. Debito";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 80;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance8;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Debitado";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 70;
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
																										 ultraGridColumn16});
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
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridLotes.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.Appearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.BandLabelAppearance = appearance14;
			this.uGridLotes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridLotes.DisplayLayout.GroupByBox.Prompt = " ";
			appearance15.BackColor = System.Drawing.SystemColors.Control;
			appearance15.BackColor2 = System.Drawing.SystemColors.Control;
			appearance15.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance15.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.PromptAppearance = appearance15;
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.uGridLotes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.uGridLotes.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.ForeColor = System.Drawing.Color.Black;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.CellAppearance = appearance18;
			this.uGridLotes.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 8F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLotes.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.uGridLotes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLotes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLotes.Location = new System.Drawing.Point(8, 56);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(1140, 240);
			this.uGridLotes.TabIndex = 637;
			this.uGridLotes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridLotes_KeyDown);
			this.uGridLotes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridLotes_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(bool);
			ultraDataColumn1.DefaultValue = false;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn4.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
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
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn14.DefaultValue = 0;
			ultraDataColumn15.DataType = typeof(System.DateTime);
			ultraDataColumn15.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn16});
			// 
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Image = ((System.Drawing.Image)(resources.GetObject("btnContabilizar.Image")));
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnContabilizar.Location = new System.Drawing.Point(1056, 8);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 23);
			this.btnContabilizar.TabIndex = 641;
			this.btnContabilizar.Text = "&Contabilizar";
			this.btnContabilizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(472, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 23);
			this.lblContador.TabIndex = 642;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1160, 8);
			this.groupBox1.TabIndex = 643;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(344, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 756;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(960, 8);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(88, 23);
			this.btnImportar.TabIndex = 757;
			this.btnImportar.Text = "Importar";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// frmFondoGarantiaPichincha
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1154, 302);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.uGridLotes);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmFondoGarantiaPichincha";
			this.Text = "Debito Fondo de Garantia Pichincha";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFondoGarantiaPichincha_KeyDown);
			this.Load += new System.EventHandler(this.frmFondoGarantiaPichincha_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}

			string sSQL = string.Format("Exec ListaClientesFondosGarantia '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridLotes);
	
			this.lblContador.Text = this.uGridLotes.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void frmFondoGarantiaPichincha_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320203");

			if (!Funcion.Permiso("963", cdsSeteoF))
			{				
				MessageBox.Show("No tiene Acceso PAGOS", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.dtFecha.CalendarInfo.MinDate = DateTime.Parse("01/05/2016");
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;
		}

		private void uGridLotes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "ValorDebito", 2);
		}

		private void uGridLotes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridLotes);
		}

		public bool ActualizaValores(string stArchivo) 
		{ 	
			int iContTotal = 0;
			int iContFDG = 0;

			if (File.Exists(stArchivo))
			{
				OleDbConnection oConn = new OleDbConnection(); 
				OleDbCommand oCmd = new OleDbCommand(); 
				OleDbDataAdapter oDa = new OleDbDataAdapter(); 
				DataSet oDs = new DataSet();
				oConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + stArchivo + "; Extended Properties= Excel 8.0;"; 
				oConn.Open(); 
				oCmd.CommandText = "SELECT * FROM [FDG$]"; 
				oCmd.Connection = oConn; 
				oDa.SelectCommand = oCmd; 
				oDa.Fill(oDs, "Lote");
 
				foreach (DataRow row in oDs.Tables[0].Rows)
				{					
					string sNumeroDeOperacion = "";
					DateTime dtFecha = DateTime.Today;
					decimal dPago = 0.00m;
										
					if (row["OPERACION"] != System.DBNull.Value) sNumeroDeOperacion = row["OPERACION"].ToString();
					if (row["FECHA_CANCELACION"] != System.DBNull.Value) dtFecha = DateTime.Parse(row["FECHA_CANCELACION"].ToString());
					if (row["TOTAL_PAGO"] != System.DBNull.Value) dPago = decimal.Parse(row["TOTAL_PAGO"].ToString());
					
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
					{												
						if (sNumeroDeOperacion.Trim() == dr.Cells["NumeroDeOperacion"].Value.ToString().Trim())
						{							
							dr.Cells["Sel"].Value = true;
							dr.Cells["FechaDebito"].Value = dtFecha;
							dr.Cells["ValorDebito"].Value = dPago;

							iContFDG++;

							break;
						}
					}

					iContTotal++;
				}

				MessageBox.Show(string.Format("{0} operaciones seleccionadas de {1} operaciones con mas de 99 dias en mora", iContFDG, iContTotal), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				return true; 
			}
			else
			{
				return false; 
			}
		}

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{	
			#region Validacion
			if (this.dtFecha.Value == System.DBNull.Value)
			{
				MessageBox.Show("Ingrese la fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				return;
			}

			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No existen operaciones en la lista", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int iContador = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)			
				if ((bool)dr.Cells["Sel"].Value) iContador++;

			if (iContador == 0)
			{
				MessageBox.Show("No existen operaciones seleccionadas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value)
				{
					#region Fecha
					if (dr.Cells["FechaDebito"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha de debito de la operacion seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridLotes.ActiveRow = dr;
						this.uGridLotes.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDateTime(dr.Cells["FechaDebito"].Value) > Convert.ToDateTime(this.dtFecha.Value))
					{
						MessageBox.Show("La fecha de debito no puede ser mayor a la fecha de corte", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridLotes.ActiveRow = dr;
						this.uGridLotes.ActiveRow.Selected = true;
						return;
					}
					#endregion Fecha	
				
					#region Valor
					if (dr.Cells["ValorDebito"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese el valor del debito de la operacion seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridLotes.ActiveRow = dr;
						this.uGridLotes.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDecimal(dr.Cells["ValorDebito"].Value) == 0.00m)
					{
						MessageBox.Show("El valor debitado no puede ser CERO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.uGridLotes.ActiveRow = dr;
						this.uGridLotes.ActiveRow.Selected = true;
						return;
					}

//					if (Convert.ToDecimal(dr.Cells["ValorDebito"].Value) <= Convert.ToDecimal(dr.Cells["Saldo"].Value))
//					{
//						MessageBox.Show("El valor debitado no puede ser menor al Saldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.uGridLotes.ActiveRow = dr;
//						this.uGridLotes.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valor
				}
			}
			#endregion Validacion	
			
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de contabilizar fondo de garantia?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{	
				this.Cursor = Cursors.WaitCursor;

				iContador = 0;

				#region Contabiliza
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
				{
					if ((bool)dr.Cells["Sel"].Value)
					{
						if (dr.Cells["FechaDebito"].Value != System.DBNull.Value)
						{								
							using (SqlConnection oConCobranzas = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
							{
								oConCobranzas.Open();

								SqlTransaction oTransaction = null;
								SqlCommand oCmdActualiza = oConCobranzas.CreateCommand();
								oCmdActualiza.Connection = oConCobranzas;
								oCmdActualiza.CommandType = CommandType.Text;
								oCmdActualiza.CommandTimeout = 120;
						
								try
								{							
									oTransaction = oConCobranzas.BeginTransaction();
									oCmdActualiza.Transaction = oTransaction;

									string sSQL = string.Format("Exec CreaNotaDeDebitoFondoGarantiaPichincha {0}, '{1}', {2}", 
										(int)dr.Cells["idCompra"].Value, Convert.ToDateTime(dr.Cells["FechaDebito"].Value).ToString("yyyyMMdd"), Convert.ToDecimal(dr.Cells["ValorDebito"].Value));
									oCmdActualiza.CommandText = sSQL;
									oCmdActualiza.ExecuteNonQuery(); 											

									oTransaction.Commit();	

									oConCobranzas.Close();

									iContador++;
								}
								catch (Exception ex)
								{
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							

									oConCobranzas.Close();	
									
									this.Cursor = Cursors.Default;

									continue;
								}									
							}
						}
					}
				}
				#endregion Contabiliza

				MessageBox.Show(string.Format("Se Contabilizarón {0} operaciones de fondo de garantia correctamente", iContador), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				
				this.btBuscar_Click(sender, e);

				this.Cursor = Cursors.Default;
			}
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog1.Filter = "Microsoft Excel (*.xls)|*.xls";

			if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName.Length > 0) 
			{									
				this.ActualizaValores(this.openFileDialog1.FileName); 
			}
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value == System.DBNull.Value) this.dtFecha.Value = DateTime.Today;
		}

		private void frmFondoGarantiaPichincha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridLotes);
		}		
	}
}
