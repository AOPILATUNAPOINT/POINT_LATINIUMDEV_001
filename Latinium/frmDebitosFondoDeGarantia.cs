using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDebitosFondoDeGarantia.
	/// </summary>
	public class frmDebitosFondoDeGarantia : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridLotes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btBuscar;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnContabilizar;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdAsiento;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnAsiento;
		private System.Windows.Forms.Button btnImportar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label lblContador;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmDebitosFondoDeGarantia()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDebitosFondoDeGarantia));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDebitosFondoDeGarantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNotaDeDebito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorDebito");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorDebito", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorDebito", 6, true);
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDebitosFondoDeGarantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNotaDeDebito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAsiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proceso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cedula");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorDebito");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.label2 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridLotes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btBuscar = new System.Windows.Forms.Button();
			this.btnImportar = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnContabilizar = new System.Windows.Forms.Button();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnAsiento = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.lblContador = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(34, 16);
			this.label2.TabIndex = 635;
			this.label2.Text = "Fecha";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(120, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 634;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(696, 8);
			this.groupBox1.TabIndex = 633;
			this.groupBox1.TabStop = false;
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
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 75;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 390;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Debito";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 90;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridLotes.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.Appearance = appearance4;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.BandLabelAppearance = appearance5;
			this.uGridLotes.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridLotes.DisplayLayout.GroupByBox.Prompt = " ";
			appearance6.BackColor = System.Drawing.SystemColors.Control;
			appearance6.BackColor2 = System.Drawing.SystemColors.Control;
			appearance6.BackColorDisabled = System.Drawing.SystemColors.Control;
			appearance6.BackColorDisabled2 = System.Drawing.SystemColors.Control;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.GroupByBox.PromptAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.ActiveRowAppearance = appearance7;
			this.uGridLotes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.uGridLotes.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridLotes.DisplayLayout.Override.CellAppearance = appearance9;
			this.uGridLotes.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridLotes.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridLotes.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridLotes.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridLotes.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridLotes.Location = new System.Drawing.Point(8, 88);
			this.uGridLotes.Name = "uGridLotes";
			this.uGridLotes.Size = new System.Drawing.Size(676, 440);
			this.uGridLotes.TabIndex = 632;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
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
																																 ultraDataColumn7});
			// 
			// btBuscar
			// 
			this.btBuscar.CausesValidation = false;
			this.btBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btBuscar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btBuscar.Location = new System.Drawing.Point(192, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(80, 23);
			this.btBuscar.TabIndex = 636;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Enabled = false;
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnImportar.Location = new System.Drawing.Point(328, 8);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(88, 23);
			this.btnImportar.TabIndex = 637;
			this.btnImportar.Text = "&Importar TXT";
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
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
			// btnContabilizar
			// 
			this.btnContabilizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnContabilizar.CausesValidation = false;
			this.btnContabilizar.Enabled = false;
			this.btnContabilizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnContabilizar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnContabilizar.Location = new System.Drawing.Point(418, 8);
			this.btnContabilizar.Name = "btnContabilizar";
			this.btnContabilizar.Size = new System.Drawing.Size(88, 23);
			this.btnContabilizar.TabIndex = 638;
			this.btnContabilizar.Text = "&Contabilizar";
			this.btnContabilizar.Click += new System.EventHandler(this.btnContabilizar_Click);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdAsiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdAsiento.Enabled = false;
			this.txtIdAsiento.Location = new System.Drawing.Point(288, 8);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.PromptChar = ' ';
			this.txtIdAsiento.Size = new System.Drawing.Size(16, 22);
			this.txtIdAsiento.TabIndex = 639;
			this.txtIdAsiento.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(8, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 640;
			// 
			// btnAsiento
			// 
			this.btnAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnAsiento.CausesValidation = false;
			this.btnAsiento.Enabled = false;
			this.btnAsiento.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAsiento.ImageIndex = 36;
			this.btnAsiento.Location = new System.Drawing.Point(508, 8);
			this.btnAsiento.Name = "btnAsiento";
			this.btnAsiento.Size = new System.Drawing.Size(88, 23);
			this.btnAsiento.TabIndex = 641;
			this.btnAsiento.Text = "Ver Asiento";
			this.btnAsiento.Click += new System.EventHandler(this.btnAsiento_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(598, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(88, 23);
			this.btnImprimir.TabIndex = 642;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtValor
			// 
			this.txtValor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValor.Appearance = appearance13;
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.Enabled = false;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(598, 40);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(88, 22);
			this.txtValor.TabIndex = 643;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(544, 43);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(30, 16);
			this.label1.TabIndex = 644;
			this.label1.Text = "Valor";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(328, 40);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 19);
			this.lblContador.TabIndex = 645;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmDebitosFondoDeGarantia
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(692, 534);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnAsiento);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btnContabilizar);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.uGridLotes);
			this.Name = "frmDebitosFondoDeGarantia";
			this.Text = "Debitos Fondo De Garantia";
			this.Load += new System.EventHandler(this.frmDebitosFondoDeGarantia_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridLotes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmDebitosFondoDeGarantia_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "08320101");

			if (!Funcion.Permiso("955", cdsSeteoF))
			{		
		
				MessageBox.Show("No tiene Acceso DEBITOS FONDOS DE GARANTIA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;

			}
			this.dtFecha.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			this.lblEstado.Text = "";
			this.lblContador.Text = "";
			this.btnContabilizar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnImportar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.txtValor.Value = 0;

			DateTime dtFecha = (DateTime)this.dtFecha.Value;
			string sSQL = string.Format("Exec ListaDebitosFondoDeGarantia '{0}'", dtFecha.ToString("yyyyMMdd"));
			this.uGridLotes.DataSource = Funcion.dvProcedimiento (cdsSeteoF, sSQL);

			string sSQLBuscaId = string.Format("Declare @RetornaIdAsiento Int = 0 If Exists(Select idAsiento From Asiento Where CONVERT(Date, Fecha) = '{0}' And idOrigenAsiento = 46) Select @RetornaIdAsiento = idAsiento From Asiento Where CONVERT(Date, Fecha) = '{0}' And idOrigenAsiento = 46 Select @RetornaIdAsiento", 
				dtFecha.ToString("yyyyMMdd"));
			this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLBuscaId);

			this.lblContador.Text = this.uGridLotes.Rows.Count + " Registros";

			this.txtValor.Value = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridLotes.Rows.All)
			{
				this.txtValor.Value = Convert.ToDecimal(this.txtValor.Value) + Convert.ToDecimal(dr.Cells["ValorDebito"].Value);
			}

			if (this.uGridLotes.Rows.Count == 0)
			{				
				this.lblEstado.Text = "PENDIENTE IMPORTAR TXT";
				this.btnImportar.Enabled = true;
			}
			if (this.uGridLotes.Rows.Count > 0 && (int)this.txtIdAsiento.Value == 0) 
			{
				this.lblEstado.Text = "PENDIENTE CONTABILIZAR";
				this.btnContabilizar.Enabled = true;
			}
			if (this.uGridLotes.Rows.Count > 0 && (int)this.txtIdAsiento.Value > 0)
			{
				this.lblEstado.Text = "CONTABILIZADO";
				this.btnAsiento.Enabled = true;
				this.btnImprimir.Enabled = true;
			}
		}

		private void btnContabilizar_Click(object sender, System.EventArgs e)
		{
			if (this.uGridLotes.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros para Contabilizar", "Point Techynology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			}
			else
			{
				this.Cursor = Cursors.WaitCursor;
				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				string sSQLAsiento = string.Format("Exec AsientoDebitosFondoDeGarantia '{0}'", dtFecha.ToString("yyyyMMdd"));
				this.txtIdAsiento.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQLAsiento, true);

				string sSQLNotaDebito = string.Format("Exec CreaNotaDeDebitoFondoGarantia '{0}', {1}", dtFecha.ToString("yyyyMMdd"), (int)this.txtIdAsiento.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLNotaDebito);
				
				MessageBox.Show("Se Contabilizó Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btBuscar_Click(sender, e);	
				
				this.Cursor = Cursors.Default;
			}
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{										
			this.openFileDialog1.Filter = "Documentos de Texto (*.txt)|*.txt";

			if(openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFileDialog1.FileName.Length > 0) 
			{	
				this.Cursor = Cursors.WaitCursor;
				
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				string sArchivo = @"\\192.168.1.19\Latinium\DFG\FondoDeGarantia_" + dtFecha.ToString("yyyyMMdd") + ".txt";
				File.Copy(openFileDialog1.FileName, sArchivo);
							
				string sSQL = string.Format(@"Exec SubeArchivoDebitosFondoDeGarantia '{0}', '{1}'", sArchivo, dtFecha.ToString("yyyyMMdd"));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				MessageBox.Show("Importación Finalizada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btBuscar_Click(sender, e);				
							
				this.Cursor = Cursors.Default;
			}
		}

		private void btnAsiento_Click(object sender, System.EventArgs e)
		{
			if ((int) this.txtIdAsiento.Value == 0) MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			else
			{					
				Cursor = Cursors.WaitCursor;
				frmAsientos Asientos = new frmAsientos((int)this.txtIdAsiento.Value);
				Asientos.MdiParent = MdiParent;
				Asientos.Show();
				Cursor = Cursors.Default;				
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			#region Impresion
			if ((int)this.txtIdAsiento.Value == 0) MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			else
			{				
				string strFiltro = "{Asiento.idAsiento} = " + this.txtIdAsiento.Value.ToString();
				
				Reporte miRepor = new Reporte("Asientos.rpt", strFiltro);
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Asiento Debitos Fondo de Garantia");
				miRepor.Show();					
			}
			#endregion Impresion
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			this.lblEstado.Text = "";
			this.lblContador.Text = "";
			this.btnContabilizar.Enabled = false;
			this.btnAsiento.Enabled = false;
			this.btnImportar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.txtValor.Value = 0;
		}
	}
}

