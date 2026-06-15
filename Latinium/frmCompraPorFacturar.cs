using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCompraPorFacturar.
	/// </summary>
	public class frmCompraPorFacturar : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid grdDetalle;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnGenerar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public static int idCompra = 0;

		public frmCompraPorFacturar(int IdCompra)
		{
			idCompra = IdCompra;
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRecibida");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPorFacturar");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadFacturada");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoFacturacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManejaSerial");
			aopilatuñaInfragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Row", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnSeriales", 1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCompraPorFacturar));
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Cantidad", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Cantidad", 4, true);
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CantidadRecibida", 6, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CantidadRecibida", 6, true);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CantidadPorFacturar", 7, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CantidadPorFacturar", 7, true);
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "CantidadFacturada", 8, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "CantidadFacturada", 8, true);
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Row", 0, false, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Row", 0, false);
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadRecibida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadPorFacturar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadFacturada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoFacturacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ManejaSerial");
			this.grdDetalle = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.grdDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// grdDetalle
			// 
			this.grdDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdDetalle.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdDetalle.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdDetalle.DisplayLayout.Appearance = appearance1;
			this.grdDetalle.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 43;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 68;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 7;
			ultraGridColumn4.Width = 194;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellButtonAppearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 103;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 109;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance4;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellButtonAppearance = appearance5;
			ultraGridColumn7.Header.Caption = "C. Recibida";
			ultraGridColumn7.Header.VisiblePosition = 9;
			ultraGridColumn7.Width = 96;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance6;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellButtonAppearance = appearance7;
			ultraGridColumn8.Header.Caption = "C. Por Facturar";
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Width = 109;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellButtonAppearance = appearance9;
			ultraGridColumn9.Header.Caption = "C. Facturada";
			ultraGridColumn9.Header.VisiblePosition = 11;
			ultraGridColumn9.Width = 115;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.Caption = "Estado Fact.";
			ultraGridColumn10.Header.VisiblePosition = 8;
			ultraGridColumn10.Width = 128;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 110;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 76;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.Caption = "...";
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 38;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance10;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellButtonAppearance = appearance11;
			ultraGridColumn14.Header.Caption = "Seriales";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn14.Width = 85;
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
																										 ultraGridColumn14});
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance12;
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance13;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance14;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings4.Appearance = appearance15;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings5.Appearance = appearance16;
			summarySettings5.DisplayFormat = "{0: #,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5});
			this.grdDetalle.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdDetalle.DisplayLayout.Override.ActiveRowAppearance = appearance17;
			this.grdDetalle.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdDetalle.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdDetalle.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance18.BackColor = System.Drawing.Color.Transparent;
			this.grdDetalle.DisplayLayout.Override.CardAreaAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance19.FontData.BoldAsString = "True";
			appearance19.FontData.Name = "Arial";
			appearance19.FontData.SizeInPoints = 10F;
			appearance19.ForeColor = System.Drawing.Color.White;
			appearance19.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdDetalle.DisplayLayout.Override.HeaderAppearance = appearance19;
			appearance20.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance20.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDetalle.DisplayLayout.Override.RowAlternateAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDetalle.DisplayLayout.Override.RowSelectorAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdDetalle.DisplayLayout.Override.SelectedRowAppearance = appearance22;
			this.grdDetalle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdDetalle.Location = new System.Drawing.Point(16, 48);
			this.grdDetalle.Name = "grdDetalle";
			this.grdDetalle.Size = new System.Drawing.Size(1184, 408);
			this.grdDetalle.TabIndex = 5;
			this.grdDetalle.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.udgDetalle_BeforeCellUpdate);
			this.grdDetalle.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdDetalle_ClickCellButton);
			this.grdDetalle.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdDetalle_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(bool);
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
																																 ultraDataColumn12});
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(776, 32);
			this.label1.TabIndex = 6;
			this.label1.Text = "En esta pantalla debe indicar las cantidades y seriales que pertenecen a la factu" +
				"ra actual del proveedor. Solo se muestran productos previamente recepcionados en" +
				" bodega.";
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCancelar.Location = new System.Drawing.Point(1120, 464);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(75, 30);
			this.btnCancelar.TabIndex = 8;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.Location = new System.Drawing.Point(1040, 464);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(75, 30);
			this.btnGenerar.TabIndex = 7;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
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
			// frmCompraPorFacturar
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1224, 510);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.grdDetalle);
			this.Name = "frmCompraPorFacturar";
			this.Text = "Recepción Pendiente de Facturación";
			this.Load += new System.EventHandler(this.frmCompraPorFacturar_Load);
			((System.ComponentModel.ISupportInitialize)(this.grdDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmCompraPorFacturar_Load(object sender, System.EventArgs e)
		{
			OrdenDeCompraDetalleRecibidoCon();
		}


		private void OrdenDeCompraDetalleRecibidoCon()
		{
			MenuLatinium.MostrarModalProcesando();
			try
			{
				this.grdDetalle.DataSource =
					FuncionesProcedimientos.dtEjecucionGeneral(
					MenuLatinium.sconexionEmpresa,
					string.Format("EXEC OrdenDeCompraDetalleRecibidoCon {0}", idCompra));

				this.grdDetalle = Funcion.ContadorFilas(this.grdDetalle, "Row");
				this.ConfigurarGridFacturacion();
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Ocurrió un error al cargar los artículos: " + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				MenuLatinium.OcultarModalProcesando();
			}
		}

		private void ConfigurarGridFacturacion()
		{
			foreach (Infragistics.Win.UltraWinGrid.UltraGridRow row in this.grdDetalle.Rows)
			{
				if (row.IsAddRow) continue;

				bool manejaSerial = false;
				if (row.Cells["ManejaSerial"].Value != DBNull.Value)
					manejaSerial = Convert.ToBoolean(row.Cells["ManejaSerial"].Value);

				if (!manejaSerial)
				{
					row.Cells["btnSeriales"].Activation =
						Infragistics.Win.UltraWinGrid.Activation.Disabled;

					foreach (Infragistics.Win.UltraWinGrid.UltraGridCell cell in row.Cells)
					{
						if (cell.Column.Key == "CantidadPorFacturar" ||
							cell.Column.Key == "btnSeriales")
							continue;
						cell.Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
					}

					row.Cells["CantidadPorFacturar"].Activation =
						Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
				}
			}
		}

		private void grdDetalle_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void udgDetalle_BeforeCellUpdate(
			object sender,
			Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.Key != "CantidadPorFacturar")
				return;

			try
			{
				decimal cantidadRecibida   = 0;
				decimal cantidadFacturada  = 0;
				decimal nuevaCantidad       = 0;

				if (e.Cell.Row.Cells["CantidadRecibida"].Value != DBNull.Value)
					try { cantidadRecibida = Convert.ToDecimal(e.Cell.Row.Cells["CantidadRecibida"].Value); } catch {}

				if (e.Cell.Row.Cells["CantidadFacturada"].Value != DBNull.Value)
					try { cantidadFacturada = Convert.ToDecimal(e.Cell.Row.Cells["CantidadFacturada"].Value); } catch {}

				if (e.NewValue == null || e.NewValue == DBNull.Value ||
					e.NewValue.ToString().Trim().Length == 0)
				{ e.Cancel = true; return; }

				try { nuevaCantidad = Convert.ToDecimal(e.NewValue); }
				catch
				{
					MessageBox.Show(
						"Ingrese un valor numérico válido.",
						"Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true; return;
				}

				if (nuevaCantidad < 0)
				{
					MessageBox.Show(
						"La cantidad por facturar no puede ser negativa.",
						"Valor inválido", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true; return;
				}

				decimal disponible = cantidadRecibida - cantidadFacturada;
				if (disponible < 0) disponible = 0;

				if (nuevaCantidad > disponible)
				{
					MessageBox.Show(
						string.Format(
							"Cantidad ingresada ({0}) supera el disponible ({1}).\nRecibida: {2}   Facturada: {3}",
							nuevaCantidad, disponible, cantidadRecibida, cantidadFacturada),
						"Cantidad no válida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				e.Cancel = true;
			}
		}

		private void grdDetalle_ClickCellButton(
			object sender,
			Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				if (e.Cell.Column.Key != "btnSeriales")
					return;

				if (e.Cell.Row.Cells["idDetCompra"].Value == DBNull.Value)
					return;

				int idDetCompra = Convert.ToInt32(e.Cell.Row.Cells["idDetCompra"].Value);

				using (frmCompraPorFacturarSeriales frm =
					new frmCompraPorFacturarSeriales(idDetCompra))
				{
					frm.ShowDialog();
					OrdenDeCompraDetalleRecibidoCon();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error al abrir seriales: " + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.Cancel;
			this.Close();
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			System.Text.StringBuilder sqlUpdates = new System.Text.StringBuilder();
			int iItemsFacturar = 0;

			// --- Paso 1: validar todos los artículos antes de escribir en BD ---
			try
			{
				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr
					in this.grdDetalle.Rows.All)
				{
					if (dr.Cells["idDetCompra"].Value == DBNull.Value) continue;

					int idDetCompra = Convert.ToInt32(dr.Cells["idDetCompra"].Value);

					string articulo = dr.Cells["Articulo"].Value != DBNull.Value
						? dr.Cells["Articulo"].Value.ToString() : "(desconocido)";

					decimal cantidadPorFacturar = 0;
					if (dr.Cells["CantidadPorFacturar"].Value != DBNull.Value)
						try { cantidadPorFacturar = Convert.ToDecimal(dr.Cells["CantidadPorFacturar"].Value); } catch {}

					if (cantidadPorFacturar > 0)
					{
						iItemsFacturar++;

						// Leer ManejaSerial del grid (ya cargado, sin consulta extra a BD)
						bool bManejaSerial = false;
						if (dr.Cells["ManejaSerial"].Value != DBNull.Value)
							try { bManejaSerial = Convert.ToBoolean(dr.Cells["ManejaSerial"].Value); } catch {}

						if (bManejaSerial)
						{
							// Seriales: verificar en BD la cantidad marcada
							int iSerialesSeleccionados = 0;
							try
							{
								iSerialesSeleccionados = Convert.ToInt32(
									Funcion.iEjecucion(MenuLatinium.sconexionEmpresa,
									string.Format(
										"SELECT COUNT(*) FROM DetalleSeriales\nWHERE idDetCompra = {0} AND EstadoFacturacion = 4 AND Disponible = 1",
										idDetCompra), false));
							}
							catch {}

							if (iSerialesSeleccionados != Convert.ToInt32(cantidadPorFacturar))
							{
								MessageBox.Show(
									string.Format(
										"Artículo '{0}': seriales marcados ({1}) no coinciden con la cantidad a facturar ({2}).\nRevise los seriales antes de continuar.",
										articulo, iSerialesSeleccionados, cantidadPorFacturar),
									"Seriales no coinciden", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}
						else
						{
							// Sin serial: leer disponible del grid (ya cargado, sin consulta extra a BD)
							decimal dCantidadRecibida  = 0;
							decimal dCantidadFacturada = 0;

							if (dr.Cells["CantidadRecibida"].Value != DBNull.Value)
								try { dCantidadRecibida = Convert.ToDecimal(dr.Cells["CantidadRecibida"].Value); } catch {}
							if (dr.Cells["CantidadFacturada"].Value != DBNull.Value)
								try { dCantidadFacturada = Convert.ToDecimal(dr.Cells["CantidadFacturada"].Value); } catch {}

							decimal dDisponible = dCantidadRecibida - dCantidadFacturada;
							if (dDisponible < 0) dDisponible = 0;

							if (cantidadPorFacturar > dDisponible)
							{
								MessageBox.Show(
									string.Format(
										"Artículo '{0}': la cantidad a facturar ({1}) supera el disponible ({2}).",
										articulo, cantidadPorFacturar, dDisponible),
									"Cantidad no válida", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}

						sqlUpdates.AppendFormat(
							"UPDATE DetCompra SET CantidadPorFacturar = {0} WHERE idDetCompra = {1};",
							cantidadPorFacturar, idDetCompra);
					}
					else
					{
						sqlUpdates.AppendFormat(
							"UPDATE DetCompra SET CantidadPorFacturar = 0 WHERE idDetCompra = {0};",
							idDetCompra);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al validar: " + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if (iItemsFacturar == 0)
			{
				MessageBox.Show(
					"No hay artículos con cantidad por facturar mayor a cero.",
					"Sin artículos", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			// --- Paso 2: confirmación del usuario ---
			if (MessageBox.Show(
				string.Format(
					"¿Confirma la generación de {0} artículo(s) para facturación?",
					iItemsFacturar),
				"Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2) != DialogResult.Yes)
				return;

			// --- Paso 3: ejecutar en BD ---
			btnGenerar.Enabled  = false;
			btnCancelar.Enabled = false;
			MenuLatinium.MostrarModalProcesando();

			bool exito = false;
			try
			{
				if (sqlUpdates.Length > 0)
					Funcion.EjecusionEjecutaSQL(
						MenuLatinium.sconexionEmpresa,
						sqlUpdates.ToString(), false);

				if (idCompra > 0)
					Funcion.EjecusionEjecutaSQL(
						MenuLatinium.sconexionEmpresa,
						string.Format(
							"EXEC OrdenDeCompraEnviarPorFacturar @idCompra = {0}, @Usuario = '{1}'",
						idCompra, MenuLatinium.stUsuario.Replace("'", "''")),
						false);

				exito = true;
				this.DialogResult = DialogResult.OK;
				this.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error al generar la facturación: " + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				MenuLatinium.OcultarModalProcesando();
				if (!exito)
				{
					btnGenerar.Enabled  = true;
					btnCancelar.Enabled = true;
				}
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}
	}
}
