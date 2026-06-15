using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_ReportesColocacion.
	/// </summary>
	public class frmCre_ReportesColocacion : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInformacion;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optReporte;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optPagos;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_ReportesColocacion()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_ReportesColocacion));
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("", -1);
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			this.optReporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btnVer = new System.Windows.Forms.Button();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.uGridInformacion = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnExportar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.optPagos = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optPagos)).BeginInit();
			this.SuspendLayout();
			// 
			// optReporte
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optReporte.Appearance = appearance1;
			this.optReporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optReporte.CheckedIndex = 0;
			this.optReporte.ItemAppearance = appearance2;
			this.optReporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Total";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Capital";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Interes";
			this.optReporte.Items.Add(valueListItem1);
			this.optReporte.Items.Add(valueListItem2);
			this.optReporte.Items.Add(valueListItem3);
			this.optReporte.ItemSpacingVertical = 10;
			this.optReporte.Location = new System.Drawing.Point(440, 8);
			this.optReporte.Name = "optReporte";
			this.optReporte.Size = new System.Drawing.Size(200, 24);
			this.optReporte.TabIndex = 1085;
			this.optReporte.Text = "Total";
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(344, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(80, 24);
			this.btnVer.TabIndex = 1083;
			this.btnVer.Text = "&Consultar";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(176, 12);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(34, 16);
			this.label8.TabIndex = 1082;
			this.label8.Text = "Hasta";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 12);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(37, 16);
			this.label6.TabIndex = 1081;
			this.label6.Text = "Desde";
			// 
			// dtHasta
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance3;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(224, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 1080;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance4;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(56, 10);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(104, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 1079;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// uGridInformacion
			// 
			this.uGridInformacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInformacion.Cursor = System.Windows.Forms.Cursors.Default;
			appearance5.BackColor = System.Drawing.Color.White;
			this.uGridInformacion.DisplayLayout.Appearance = appearance5;
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInformacion.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInformacion.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInformacion.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.uGridInformacion.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInformacion.DisplayLayout.Override.CellAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInformacion.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.uGridInformacion.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.RowSelectorAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInformacion.DisplayLayout.Override.SelectedRowAppearance = appearance10;
			this.uGridInformacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInformacion.Location = new System.Drawing.Point(8, 56);
			this.uGridInformacion.Name = "uGridInformacion";
			this.uGridInformacion.Size = new System.Drawing.Size(1112, 296);
			this.uGridInformacion.TabIndex = 1086;
			// 
			// btnExportar
			// 
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(1048, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(75, 24);
			this.btnExportar.TabIndex = 1087;
			this.btnExportar.Text = "Excel";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 40);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1128, 8);
			this.groupBox1.TabIndex = 1088;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(656, 12);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 16);
			this.lblContador.TabIndex = 1089;
			// 
			// optPagos
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.optPagos.Appearance = appearance11;
			this.optPagos.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optPagos.CheckedIndex = 0;
			this.optPagos.ItemAppearance = appearance12;
			this.optPagos.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Todo";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "Cobros";
			valueListItem6.DataValue = 3;
			valueListItem6.DisplayText = "Notas De Crédito";
			this.optPagos.Items.Add(valueListItem4);
			this.optPagos.Items.Add(valueListItem5);
			this.optPagos.Items.Add(valueListItem6);
			this.optPagos.ItemSpacingVertical = 10;
			this.optPagos.Location = new System.Drawing.Point(656, 8);
			this.optPagos.Name = "optPagos";
			this.optPagos.Size = new System.Drawing.Size(248, 24);
			this.optPagos.TabIndex = 1090;
			this.optPagos.Text = "Todo";
			// 
			// frmCre_ReportesColocacion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1128, 358);
			this.Controls.Add(this.optPagos);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.uGridInformacion);
			this.Controls.Add(this.optReporte);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.KeyPreview = true;
			this.Name = "frmCre_ReportesColocacion";
			this.Text = "Evolución Colocación Crédito";
			this.Load += new System.EventHandler(this.frmCre_ReportesColocacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.optReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridInformacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optPagos)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCre_ReportesColocacion_Load(object sender, System.EventArgs e)
		{			
			this.dtDesde.Value = DateTime.Parse("01/01/" + DateTime.Today.Year);
			this.dtHasta.Value = DateTime.Today.AddDays(- 1);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbValidaFechasDesdeHasta(this.dtDesde, this.dtHasta)) return;

			this.uGridInformacion.DataSource = null;

			string sSQL = string.Format("Exec Cre_EvolucionColocacionCredito '{0}', '{1}', {2}, {3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), 
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), 
				(int)this.optReporte.Value, 
				(int)this.optPagos.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInformacion);

			this.DiseñoGrid();

			this.lblContador.Text = this.uGridInformacion.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void DiseñoGrid()
		{
			if (this.uGridInformacion.Rows.Count > 0)
			{				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Colocacion"].Width = 60;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Numero"].Width = 80;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Almacen"].Width = 140;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cedula"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cliente"].Width = 200;				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Banco"].Width = 120;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Retiro"].Width = 80;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["NotaDeCredito"].Width = 80;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Monto"].Width = 100;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["TotalCredito"].Width = 100;
				//this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Width = 105;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PrimerPago"].Width = 80;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["FDGCaida"].Width = 80;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["TotalCredito"].Header.Caption = "Total Credito";	
			  this.uGridInformacion.DisplayLayout.Bands[0].Columns["PrimerPago"].Header.Caption = "Primer Pago";
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["FDGCaida"].Header.Caption = "Caida F. D. G.";
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["NotaDeCredito"].Header.Caption = "N/C";
				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Colocacion"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Numero"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Almacen"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cedula"].CellActivation = Activation.Disabled;				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cliente"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Banco"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Retiro"].CellActivation = Activation.Disabled;			
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["NotaDeCredito"].CellActivation = Activation.Disabled;			
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Monto"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["TotalCredito"].CellActivation = Activation.Disabled;
				//this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PrimerPago"].CellActivation = Activation.Disabled;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["FDGCaida"].CellActivation = Activation.Disabled;

				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Colocacion"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Numero"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Cedula"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Banco"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Retiro"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["NotaDeCredito"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Monto"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["TotalCredito"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				//this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["PrimerPago"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["FDGCaida"].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center;
				
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["Monto"].Format = "#,##0.00";
				this.uGridInformacion.DisplayLayout.Bands[0].Columns["TotalCredito"].Format = "#,##0.00";
				//this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"].Format = "#,##0.00";

				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Monto", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Monto"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Monto"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Monto"].DisplayFormat = "{0: #,##0.00}";

				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("TotalCredito", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["TotalCredito"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["TotalCredito"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["TotalCredito"].DisplayFormat = "{0: #,##0.00}";

//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add("Total", Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns["Total"], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
//				this.uGridInformacion.DisplayLayout.Bands[0].Summaries["Total"].DisplayFormat = "{0: #,##0.00}";

				int iContColumnas = (int)this.uGridInformacion.DisplayLayout.Bands[0].Columns.Count;

				int iContador = 12;

				string nombrCol = "";

				for (int j = iContador; j < iContColumnas; j++)
				{
					nombrCol = this.uGridInformacion.DisplayLayout.Bands[0].Columns[j].Header.Caption;
					
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellActivation = Activation.Disabled;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol].Width = 100;

					this.uGridInformacion.DisplayLayout.Bands[0].Summaries.Add(nombrCol, Infragistics.Win.UltraWinGrid.SummaryType.Sum, this.uGridInformacion.DisplayLayout.Bands[0].Columns[nombrCol], Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn);
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].Appearance.TextHAlign = Infragistics.Win.HAlign.Right;
					this.uGridInformacion.DisplayLayout.Bands[0].Summaries[nombrCol].DisplayFormat = "{0: #,##0.00}";					
				}
			}
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridInformacion);
		}
	}
}
