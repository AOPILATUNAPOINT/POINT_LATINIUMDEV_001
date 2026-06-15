using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReportesFacturacionEncuestas.
	/// </summary>
	public class frmReportesFacturacionEncuestas : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReporte;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReportesFacturacionEncuestas()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReportesFacturacionEncuestas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label53 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.optReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).BeginInit();
			this.SuspendLayout();
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(31, 16);
			this.label2.TabIndex = 393;
			this.label2.Text = "Local";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 180;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(56, 40);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(220, 21);
			this.cmbBodega.TabIndex = 392;
			this.cmbBodega.ValueMember = "Bodega";
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Location = new System.Drawing.Point(184, 9);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(34, 16);
			this.label53.TabIndex = 391;
			this.label53.Text = "Hasta";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.BackColor = System.Drawing.Color.Transparent;
			this.lblBodega.Location = new System.Drawing.Point(10, 9);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(37, 16);
			this.lblBodega.TabIndex = 390;
			this.lblBodega.Text = "Desde";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(368, 38);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 386;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(288, 38);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 385;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(232, 7);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 388;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance2;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 7);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 387;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance3.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance3;
			this.uGridInformacion.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.uGridInformacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 88);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(728, 256);
			this.uGridInformacion.TabIndex = 389;
			this.uGridInformacion.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridInformacion_InitializeLayout);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(744, 8);
			this.groupBox1.TabIndex = 394;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.lblContador.Location = new System.Drawing.Point(456, 42);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 395;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// optReporte
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReporte.Appearance = appearance8;
			this.optReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReporte.CheckedIndex = 0;
			this.optReporte.ItemAppearance = appearance9;
			this.optReporte.ItemOrigin = new System.Drawing.Point(12, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Detallado";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Por Local";
			this.optReporte.Items.Add(valueListItem1);
			this.optReporte.Items.Add(valueListItem2);
			this.optReporte.ItemSpacingVertical = 12;
			this.optReporte.Location = new System.Drawing.Point(368, 5);
			this.optReporte.Name = "optReporte";
			this.optReporte.Size = new System.Drawing.Size(160, 24);
			this.optReporte.TabIndex = 396;
			this.optReporte.Text = "Detallado";
			// 
			// frmReportesFacturacionEncuestas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(744, 350);
			this.Controls.Add(this.optReporte);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label53);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.uGridInformacion);
			this.KeyPreview = true;
			this.Name = "frmReportesFacturacionEncuestas";
			this.Text = "Encuestas";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReportesFacturacionEncuestas_KeyDown);
			this.Load += new System.EventHandler(this.frmReportesFacturacionEncuestas_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReportesFacturacionEncuestas_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("29/04/2017");

			this.dtDesde.Value = DateTime.Parse("01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString());			
			this.dtHasta.Value = DateTime.Today;

			this.cmbBodega.DataSource = FuncionesProcedimientos.dtGeneral("Select Bodega, Codigo, Nombre From Bodega Where Activo = 1 And Factura = 1 Order By Nombre");			
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			this.uGridInformacion.DataSource = null;

			string sSQL = string.Format("Exec ReportesFacturacionEncuestas '{0}', '{1}', {2}, {3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"),
				iBodega, (int)this.optReporte.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

			DiseñoGrid();

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " Registros encontrados";
		}

		private void DiseñoGrid()
		{
			#region Diseño Grid
			if (this.uGridInformacion.Rows.Count > 0)
			{				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Local"].Width = 160;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Local"].CellActivation = Activation.ActivateOnly;

				#region Detallado
				if ((int)this.optReporte.Value == 1)
				{
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Numero"].Width = 80;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Numero"].CellActivation = Activation.ActivateOnly;

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Fecha"].Width = 70;				
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Fecha"].CellActivation = Activation.ActivateOnly;
				
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["FormaPago"].Width = 90;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["FormaPago"].Header.Caption = "Forma De Pago";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["FormaPago"].CellActivation = Activation.ActivateOnly;

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ruc"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ruc"].Header.Caption = "Cedula";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ruc"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Ruc"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Nombre"].Width = 220;								
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Nombre"].CellActivation = Activation.ActivateOnly;

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Valor"].Width = 70;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Valor"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Valor"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Valor", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Valor"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Valor"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Valor"].DisplayFormat = "{0: #,##0.00}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Medio"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Medio"].CellActivation = Activation.ActivateOnly;
				
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["OtrosEncuesta"].Width = 150;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["OtrosEncuesta"].Header.Caption = "Notas";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["OtrosEncuesta"].CellActivation = Activation.ActivateOnly;
				}
				#endregion Detallado

				#region Por Local
				if ((int)this.optReporte.Value == 2)
				{
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["FACEBOOK"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["FACEBOOK"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["FACEBOOK"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("FACEBOOK", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["FACEBOOK"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["FACEBOOK"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["FACEBOOK"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["CATALOGO"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["CATALOGO"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["CATALOGO"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("CATALOGO", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["CATALOGO"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["CATALOGO"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["CATALOGO"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["RADIO"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["RADIO"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["RADIO"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("RADIO", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["RADIO"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["RADIO"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["RADIO"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["RECOMENDADO"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["RECOMENDADO"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["RECOMENDADO"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("RECOMENDADO", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["RECOMENDADO"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["RECOMENDADO"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["RECOMENDADO"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["REITERATIVO"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["REITERATIVO"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["REITERATIVO"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("REITERATIVO", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["REITERATIVO"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["REITERATIVO"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["REITERATIVO"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PASO POR EL LOCAL"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PASO POR EL LOCAL"].Header.Caption = "DE PASO";
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PASO POR EL LOCAL"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PASO POR EL LOCAL"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("PASO POR EL LOCAL", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["PASO POR EL LOCAL"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PASO POR EL LOCAL"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PASO POR EL LOCAL"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PUERTEO"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PUERTEO"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["PUERTEO"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("PUERTEO", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["PUERTEO"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PUERTEO"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["PUERTEO"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["SMS"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["SMS"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["SMS"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("SMS", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["SMS"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["SMS"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["SMS"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["WEB"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["WEB"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["WEB"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("WEB", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["WEB"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["WEB"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["WEB"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["WHATSAPP"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["WHATSAPP"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["WHATSAPP"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("WHATSAPP", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["WHATSAPP"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["WHATSAPP"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["WHATSAPP"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["OTROS"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["OTROS"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["OTROS"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("OTROS", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["OTROS"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["OTROS"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["OTROS"].DisplayFormat = "{0: #,##0}";

					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 100;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.ActivateOnly;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor = Color.Yellow;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.BackColor2 = Color.Yellow;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Center;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0}";

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInformacion.Rows.All)
					{
						dr.Cells["Total"].Value = Convert.ToInt32(dr.Cells["FACEBOOK"].Value) + 
							Convert.ToInt32(dr.Cells["CATALOGO"].Value) + 
							Convert.ToInt32(dr.Cells["RADIO"].Value) + 
							Convert.ToInt32(dr.Cells["RECOMENDADO"].Value) + 
							Convert.ToInt32(dr.Cells["REITERATIVO"].Value) + 
							Convert.ToInt32(dr.Cells["PASO POR EL LOCAL"].Value) + 
							Convert.ToInt32(dr.Cells["PUERTEO"].Value) + 
							Convert.ToInt32(dr.Cells["SMS"].Value) + 
							Convert.ToInt32(dr.Cells["WEB"].Value) + 
						Convert.ToInt32(dr.Cells["WHATSAPP"].Value) + 
							Convert.ToInt32(dr.Cells["OTROS"].Value);
					}
				}
				#endregion Por Local				
			}
			#endregion Diseño Grid
		}

		private void frmReportesFacturacionEncuestas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void uGridInformacion_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
