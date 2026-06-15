using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCom_InformeIndividualComisiones.
	/// </summary>
	public class frmCom_InformeIndividualComisiones : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridInforme;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblPeriodo;
		private System.Windows.Forms.Label lblValor;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCom_InformeIndividualComisiones()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int idPersonal = 0;
		string sPeriodo = "";
		int idCargo = 0;
		int iBodega = 0;
		DateTime dtDesde = DateTime.Today;
		DateTime dtHasta = DateTime.Today;
		string sNombre = "";
		bool bImprimir = false;
		decimal dValorARecibir = 0.00m;

		public frmCom_InformeIndividualComisiones(int IdPersonal, string SPeriodo, int IdCargo, int IBodega, DateTime DTDesde, DateTime DTHasta, string SNombre, bool BImprimir)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			idPersonal = IdPersonal;
			sPeriodo = SPeriodo;
			idCargo = IdCargo;
			iBodega = IBodega;
			dtDesde = DTDesde;
			dtHasta = DTHasta;
			sNombre = SNombre;
			bImprimir = BImprimir;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Meta");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Venta");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("VentaTotal");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeAlcanzado");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RangoComision");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorComision");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Comision", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Comision", 9, true);
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCom_InformeIndividualComisiones));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Concepto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Meta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Venta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("VentaTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeAlcanzado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RangoComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorComision");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comision");
			this.uGridInforme = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblPeriodo = new System.Windows.Forms.Label();
			this.lblValor = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.uGridInforme)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridInforme
			// 
			this.uGridInforme.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridInforme.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridInforme.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridInforme.DisplayLayout.Appearance = appearance1;
			this.uGridInforme.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance2.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn1.CellAppearance = appearance2;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 8;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 400;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance3.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn4.CellAppearance = appearance3;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance4.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn5.CellAppearance = appearance4;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance5.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn6.CellAppearance = appearance5;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 80;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance6.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn7.CellAppearance = appearance6;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "% Alcanzado";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 80;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance7.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn8.CellAppearance = appearance7;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.Caption = "% Comision";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance8.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Valor";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance9.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn10.CellAppearance = appearance9;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
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
																										 ultraGridColumn10});
			appearance10.TextHAlign = Infragistics.Win.HAlign.Left;
			ultraGridBand1.Header.Appearance = appearance10;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance11;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.uGridInforme.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridInforme.DisplayLayout.ColumnScrollbarSmallChange = 50;
			this.uGridInforme.DisplayLayout.GroupByBox.Hidden = true;
			this.uGridInforme.DisplayLayout.GroupByBox.Prompt = " ";
			this.uGridInforme.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridInforme.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridInforme.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance12.BackColor = System.Drawing.Color.Transparent;
			this.uGridInforme.DisplayLayout.Override.CardAreaAppearance = appearance12;
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridInforme.DisplayLayout.Override.CellAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance14.FontData.BoldAsString = "True";
			appearance14.FontData.Name = "Arial";
			appearance14.FontData.SizeInPoints = 8F;
			appearance14.ForeColor = System.Drawing.Color.White;
			appearance14.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridInforme.DisplayLayout.Override.HeaderAppearance = appearance14;
			this.uGridInforme.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance15.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance15.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.RowAlternateAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.RowSelectorAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridInforme.DisplayLayout.Override.SelectedRowAppearance = appearance17;
			this.uGridInforme.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.uGridInforme.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridInforme.Location = new System.Drawing.Point(8, 96);
			this.uGridInforme.Name = "uGridInforme";
			this.uGridInforme.Size = new System.Drawing.Size(909, 264);
			this.uGridInforme.TabIndex = 121;
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImprimir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(840, 8);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(80, 23);
			this.btnImprimir.TabIndex = 652;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Visible = false;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.BackColor = System.Drawing.Color.Transparent;
			this.lblVendedor.Font = new System.Drawing.Font("Tahoma", 11F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblVendedor.ForeColor = System.Drawing.Color.Firebrick;
			this.lblVendedor.Location = new System.Drawing.Point(16, 0);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(0, 21);
			this.lblVendedor.TabIndex = 653;
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn4.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn5.DataType = typeof(System.Decimal);
			ultraDataColumn5.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn6.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn7.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
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
																																 ultraDataColumn10});
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 80);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(928, 8);
			this.groupBox1.TabIndex = 654;
			this.groupBox1.TabStop = false;
			// 
			// lblPeriodo
			// 
			this.lblPeriodo.AutoSize = true;
			this.lblPeriodo.BackColor = System.Drawing.Color.Transparent;
			this.lblPeriodo.Font = new System.Drawing.Font("Tahoma", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblPeriodo.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPeriodo.Location = new System.Drawing.Point(16, 29);
			this.lblPeriodo.Name = "lblPeriodo";
			this.lblPeriodo.Size = new System.Drawing.Size(0, 18);
			this.lblPeriodo.TabIndex = 655;
			this.lblPeriodo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblValor
			// 
			this.lblValor.AutoSize = true;
			this.lblValor.BackColor = System.Drawing.Color.Transparent;
			this.lblValor.Font = new System.Drawing.Font("Tahoma", 9F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblValor.ForeColor = System.Drawing.Color.Firebrick;
			this.lblValor.Location = new System.Drawing.Point(16, 55);
			this.lblValor.Name = "lblValor";
			this.lblValor.Size = new System.Drawing.Size(0, 18);
			this.lblValor.TabIndex = 656;
			this.lblValor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmCom_InformeIndividualComisiones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(923, 368);
			this.Controls.Add(this.lblValor);
			this.Controls.Add(this.lblPeriodo);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.uGridInforme);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmCom_InformeIndividualComisiones";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Informe Individual Comisiones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCom_InformeIndividualComisiones_KeyDown);
			this.Load += new System.EventHandler(this.frmCom_InformeIndividualComisiones_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridInforme)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCom_InformeIndividualComisiones_Load(object sender, System.EventArgs e)
		{
			if (bImprimir) this.btnImprimir.Visible = true;

			this.lblVendedor.Text = sNombre;
			this.lblPeriodo.Text =  "PERIODO : " + dtDesde.ToString("MMMM/yyyy");

			string sSQL = string.Format("Exec Com_InformeComisionesIndividual {0}, '{1}', {2}, {3}, '{4}', '{5}', '{6}'", 
				idPersonal, sPeriodo, idCargo, iBodega, dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"), sNombre);
      FuncionesProcedimientos.dsGeneral(sSQL, this.uGridInforme);	
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridInforme.Rows.All)
				dValorARecibir = dValorARecibir + Convert.ToDecimal(dr.Cells["Comision"].Value);
			
			this.lblValor.Text = "VALOR A RECIBIR : " + dValorARecibir.ToString();

			this.uGridInforme.DisplayLayout.Bands[0].SortedColumns.Add("Tipo", false, true);			
			this.uGridInforme.Rows.ExpandAll(true);
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			ParameterFields paramFields = new ParameterFields ();

			ParameterField idPersonal = new ParameterField ();
			ParameterDiscreteValue didPersonal = new ParameterDiscreteValue ();
			idPersonal.ParameterFieldName = "@idPersonal";
			didPersonal.Value = idPersonal;
			idPersonal.CurrentValues.Add (didPersonal);
			paramFields.Add (idPersonal);

			ParameterField Periodo = new ParameterField ();
			ParameterDiscreteValue dPeriodo = new ParameterDiscreteValue ();
			Periodo.ParameterFieldName = "@Periodo";
			dPeriodo.Value = sPeriodo;
			Periodo.CurrentValues.Add (dPeriodo);
			paramFields.Add (Periodo);
			
			ParameterField idCargo = new ParameterField ();
			ParameterDiscreteValue didCargo = new ParameterDiscreteValue ();
			idCargo.ParameterFieldName = "@idCargo";
			didCargo.Value = idCargo;
			idCargo.CurrentValues.Add (didCargo);
			paramFields.Add (idCargo);

			ParameterField Bodega = new ParameterField ();
			ParameterDiscreteValue dBodega = new ParameterDiscreteValue ();
			Bodega.ParameterFieldName = "@Bodega";
			dBodega.Value = iBodega;
			Bodega.CurrentValues.Add (dBodega);
			paramFields.Add (Bodega);

			ParameterField Desde = new ParameterField ();
			ParameterDiscreteValue dDesde = new ParameterDiscreteValue ();
			Desde.ParameterFieldName = "@Desde";
			dDesde.Value = dtDesde;
			Desde.CurrentValues.Add (dDesde);
			paramFields.Add (Desde);

			ParameterField Hasta = new ParameterField ();
			ParameterDiscreteValue dHasta = new ParameterDiscreteValue ();
			Hasta.ParameterFieldName = "@Hasta";
			dHasta.Value = dtHasta;
			Hasta.CurrentValues.Add (dHasta);
			paramFields.Add (Hasta);

			ParameterField Nombre = new ParameterField ();
			ParameterDiscreteValue dNombre = new ParameterDiscreteValue ();
			Nombre.ParameterFieldName = "@Nombre";
			dNombre.Value = sNombre;
			Nombre.CurrentValues.Add (dNombre);
			paramFields.Add (Nombre);

			Reporte Rep = new Reporte("Com_InformeComisionesIndividual.rpt", "");
			Rep.MdiParent = this.MdiParent;			
			Rep.crVista.ParameterFieldInfo = paramFields;
			Rep.Show();
						
			Cursor = Cursors.Default;
		}

		private void frmCom_InformeIndividualComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
