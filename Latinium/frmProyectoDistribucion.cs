using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmProyectoDistribucion.
	/// </summary>
	public class frmProyectoDistribucion : System.Windows.Forms.Form
	{
		decimal sAgencias = 0;
		decimal sAdministrativos = 0;
		decimal sGeneral = 0;
		int idProyecto = 0;
		decimal dAfter = 0;
		bool bActualiza = false;
		
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.DateTimePicker dtpPeriodo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalAgencias;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalAdministrativos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalGeneral;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmProyectoDistribucion()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Agencias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Administrativos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("General");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Agencias");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Administrativos");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("General");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Agencias", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Agencias", 4, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Administrativos", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Administrativos", 5, true);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "General", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "General", 6, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmProyectoDistribucion));
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtTotalAgencias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.dtpPeriodo = new System.Windows.Forms.DateTimePicker();
			this.txtTotalAdministrativos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalGeneral = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalAgencias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalAdministrativos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneral)).BeginInit();
			this.SuspendLayout();
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
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance1;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "BodegaPromocion";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 58;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.Caption = "Almacen";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 162;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 53;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Width = 69;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 114;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 113;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 124;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			ultraGridBand1.Header.Caption = "PROYECTOS";
			ultraGridBand1.HeaderVisible = true;
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
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance8;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 10F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance11.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.ultraGrid2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 48);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(656, 416);
			this.ultraGrid2.TabIndex = 185;
			this.ultraGrid2.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid2_BeforeCellUpdate);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid2_AfterCellUpdate);
			// 
			// txtTotalAgencias
			// 
			this.txtTotalAgencias.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalAgencias.Appearance = appearance14;
			this.txtTotalAgencias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalAgencias.Enabled = false;
			this.txtTotalAgencias.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotalAgencias.FormatString = "#,##0.00";
			this.txtTotalAgencias.Location = new System.Drawing.Point(448, 424);
			this.txtTotalAgencias.Name = "txtTotalAgencias";
			this.txtTotalAgencias.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalAgencias.PromptChar = ' ';
			this.txtTotalAgencias.Size = new System.Drawing.Size(168, 34);
			this.txtTotalAgencias.TabIndex = 186;
			this.txtTotalAgencias.Visible = false;
			this.txtTotalAgencias.ValueChanged += new System.EventHandler(this.txtTotal_ValueChanged);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(320, 16);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 851;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(592, 15);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 868;
			// 
			// btnGenerar
			// 
			this.btnGenerar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGenerar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(240, 17);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(72, 23);
			this.btnGenerar.TabIndex = 872;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// btnVer
			// 
			this.btnVer.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnVer.CausesValidation = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnVer.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(160, 17);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 871;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 20);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(43, 16);
			this.label4.TabIndex = 869;
			this.label4.Text = "Periodo";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpPeriodo
			// 
			this.dtpPeriodo.CustomFormat = "MM/yyyy";
			this.dtpPeriodo.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dtpPeriodo.Location = new System.Drawing.Point(56, 18);
			this.dtpPeriodo.MinDate = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			this.dtpPeriodo.Name = "dtpPeriodo";
			this.dtpPeriodo.ShowUpDown = true;
			this.dtpPeriodo.Size = new System.Drawing.Size(96, 20);
			this.dtpPeriodo.TabIndex = 870;
			this.dtpPeriodo.Value = new System.DateTime(2016, 5, 1, 0, 0, 0, 0);
			// 
			// txtTotalAdministrativos
			// 
			this.txtTotalAdministrativos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalAdministrativos.Appearance = appearance15;
			this.txtTotalAdministrativos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalAdministrativos.Enabled = false;
			this.txtTotalAdministrativos.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotalAdministrativos.FormatString = "#,##0.00";
			this.txtTotalAdministrativos.Location = new System.Drawing.Point(456, 424);
			this.txtTotalAdministrativos.Name = "txtTotalAdministrativos";
			this.txtTotalAdministrativos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalAdministrativos.PromptChar = ' ';
			this.txtTotalAdministrativos.Size = new System.Drawing.Size(168, 34);
			this.txtTotalAdministrativos.TabIndex = 873;
			this.txtTotalAdministrativos.Visible = false;
			// 
			// txtTotalGeneral
			// 
			this.txtTotalGeneral.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalGeneral.Appearance = appearance16;
			this.txtTotalGeneral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalGeneral.Enabled = false;
			this.txtTotalGeneral.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotalGeneral.FormatString = "#,##0.00";
			this.txtTotalGeneral.Location = new System.Drawing.Point(480, 424);
			this.txtTotalGeneral.Name = "txtTotalGeneral";
			this.txtTotalGeneral.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalGeneral.PromptChar = ' ';
			this.txtTotalGeneral.Size = new System.Drawing.Size(168, 34);
			this.txtTotalGeneral.TabIndex = 874;
			this.txtTotalGeneral.Visible = false;
			// 
			// frmProyectoDistribucion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(672, 462);
			this.Controls.Add(this.txtTotalGeneral);
			this.Controls.Add(this.txtTotalAdministrativos);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.dtpPeriodo);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.txtTotalAgencias);
			this.Controls.Add(this.ultraGrid2);
			this.MinimizeBox = false;
			this.Name = "frmProyectoDistribucion";
			this.Text = "Proyectos Distribucion";
			this.Load += new System.EventHandler(this.frmProyectoDistribucion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalAgencias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalAdministrativos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalGeneral)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void UnloadMe()
		{
			this.Close();
		}
		private void frmProyectoDistribucion_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("1716", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Proyectos Distribucion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			this.dtpPeriodo.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());
			//this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaProyectosDistribucion "));
			Consulta();
			Total();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}
		private void Total()
		{
			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (dr.Cells["Agencias"].Value != DBNull.Value)
				{
					sAgencias= sAgencias + Math.Round(Convert.ToDecimal(dr.Cells["Agencias"].Value), 2);
					this.txtTotalAgencias.Value=sAgencias;
					//this.lblEstado.Text=dSaldo.ToString();	
				}
				if (dr.Cells["Administrativos"].Value != DBNull.Value)
				{
					sAdministrativos= sAdministrativos + Math.Round(Convert.ToDecimal(dr.Cells["Administrativos"].Value), 2);
					this.txtTotalAdministrativos.Value=sAdministrativos;
					//this.lblEstado.Text=dSaldo.ToString();
				}
				if (dr.Cells["General"].Value != DBNull.Value)
				{
					sGeneral= sGeneral + Math.Round(Convert.ToDecimal(dr.Cells["General"].Value), 2);
					this.txtTotalGeneral.Value=sGeneral;
					//this.lblEstado.Text=dSaldo.ToString();
				}
			}
			sAgencias=0;
			sAdministrativos=0;
			sGeneral=0;
			#endregion Desplazamiento de grilla			
		}

		private void ultraGrid2_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			Total();
		}

		private void ultraGrid2_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
	
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			Total();
			#region valida Porcentajes
			if (Convert.ToDecimal(this.txtTotalAgencias.Value) > 100.00m  ||  Convert.ToDecimal(this.txtTotalAgencias.Value)<  100.00m)
			{
				this.txtTotalAgencias.Value=0;
				MessageBox.Show("El porcentaje debe ser igual a 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Convert.ToDecimal(this.txtTotalAdministrativos.Value) > 100.00m  ||  Convert.ToDecimal(this.txtTotalAdministrativos.Value)<  100.00m)
			{
				this.txtTotalAdministrativos.Value=0;
				MessageBox.Show("El porcentaje debe ser igual a 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (Convert.ToDecimal(this.txtTotalGeneral.Value) > 100.00m  ||  Convert.ToDecimal(this.txtTotalGeneral.Value)<  100.00m)
			{
				this.txtTotalGeneral.Value=0;
				MessageBox.Show("El porcentaje debe ser igual a 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion valida Porcentajes
			 
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
//			{	
//				//				if ((bool)dr.Cells["Sel"].Value == false && Convert.ToDecimal(dr.Cells["Porcentaje"].Value) >0.00m )
//				//				{
//				//					this.txtTotal.Value=0;
//				//					MessageBox.Show("No puede asignar porcentaje si no esta seleccionado la opcion de Distribucion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				//					return;
//				//				}	
//				if ((bool)dr.Cells["Sel"].Value == true && Convert.ToDecimal(dr.Cells["Porcentaje"].Value) ==0.00m  )
//				{
//					this.txtTotal.Value=0;
//					MessageBox.Show("El porcentaje  no puede ser 0 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}	
//				if ((bool)dr.Cells["Sel"].Value == true && Convert.ToDecimal(dr.Cells["Porcentaje"].Value) <0.00m )
//				{
//					this.txtTotal.Value=0;
//					MessageBox.Show("El porcentaje  no puede ser meno a 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}	
//				if ((bool)dr.Cells["Sel"].Value == true && Convert.ToDecimal(dr.Cells["Porcentaje"].Value) >100.00m )
//				{
//					this.txtTotal.Value=0;
//					MessageBox.Show("El porcentaje  no puede ser mayor a 100 ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					return;
//				}	
//			}
//			if((Convert.ToDecimal(this.txtTotal.Value) == 100.00m))
//			{
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 0;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
						{
							string sSQLDetalle = string.Format("Exec ActualizaDistribucion {0},{1},{2},{3},'{4}'",
								(decimal)dr.Cells["Agencias"].Value,
								(decimal)dr.Cells["Administrativos"].Value,
								(decimal)dr.Cells["General"].Value, 
								(int)dr.Cells["idProyecto"].Value,
								this.dtpPeriodo.Text.ToString());
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						oTransaction.Commit();
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
				Consulta();
				//this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaProyectosDistribucion "));
				MessageBox.Show("Información grabada correctamente...!!!", "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
//			}
//			else 
//			{
//				MessageBox.Show("El porcentaje debe ser igual a 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.btnGuardar.Enabled = false;
		}

		private void txtTotal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From ProyectoDistribucion Where Periodo = '{0}'", this.dtpPeriodo.Text.ToString())) > 0)
			{
				MessageBox.Show(string.Format("La distribución se encuentra creada para periodo {0}", this.dtpPeriodo.Text.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.btnVer_Click(sender, e);
				this.dtpPeriodo.Focus();
				return;
			}

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de Generar la Asignacion de personal para el periodo {0}", this.dtpPeriodo.Text.ToString()), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				string sSQL = string.Format("Exec GeneraProyectoDistribucion '{0}'", this.dtpPeriodo.Text.ToString());
				Funcion.EjecutaSQL(cdsSeteoF, sSQL);

				MessageBox.Show("Generación Terminada Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

				this.btnVer_Click(sender, e);
			}
		}
		private void Consulta()
		{
			if (this.dtpPeriodo.Text.ToString().Length == 0) return;			

			string sSQL = string.Format("Exec ListaProyectosDistribucion '{0}'", 
				this.dtpPeriodo.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid2);
	
			this.ultraGrid2.Rows.ExpandAll(true);

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (this.dtpPeriodo.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtpPeriodo.Focus();
				return;
			}

/*			string sSQL = string.Format("Exec GeneraProyectoDistribucion '{0}'", this.dtpPeriodo.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);*/

			//this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaProyectosDistribucion "));
			Consulta();
			Total();
		}

	}

}
