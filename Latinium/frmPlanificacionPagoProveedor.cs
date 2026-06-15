using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPlanificacionPagoProveedor.
	/// </summary>
	public class frmPlanificacionPagoProveedor : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnBusca;
		public Infragistics.Win.UltraWinGrid.UltraGrid udtG;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Button btnExportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPlanificacionPagoProveedor()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPlanificacionPagoProveedor));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idcompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaVencimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoAnticipo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPagado");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumPago");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "SaldoAnticipo", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "SaldoAnticipo", 5, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "ValorPagado", 9, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "ValorPagado", 9, true);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 12, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 12, true);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idcompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaVencimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoAnticipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Grupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorPagado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			this.lblHasta = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnBusca = new System.Windows.Forms.Button();
			this.udtG = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnExportar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udtG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// lblHasta
			// 
			this.lblHasta.AutoSize = true;
			this.lblHasta.BackColor = System.Drawing.Color.Transparent;
			this.lblHasta.Location = new System.Drawing.Point(8, 8);
			this.lblHasta.Name = "lblHasta";
			this.lblHasta.Size = new System.Drawing.Size(14, 16);
			this.lblHasta.TabIndex = 640;
			this.lblHasta.Text = "Al";
			// 
			// dtHasta
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance1;
			this.dtHasta.AutoSelect = false;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(64, 8);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 639;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			this.dtHasta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtHasta_KeyDown);
			// 
			// btnBusca
			// 
			this.btnBusca.Location = new System.Drawing.Point(192, 8);
			this.btnBusca.Name = "btnBusca";
			this.btnBusca.TabIndex = 641;
			this.btnBusca.Text = "&Buscar";
			this.btnBusca.Click += new System.EventHandler(this.btnBusca_Click);
			// 
			// udtG
			// 
			this.udtG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.udtG.Cursor = System.Windows.Forms.Cursors.Default;
			this.udtG.DataSource = this.ultraDataSource1;
			appearance2.BackColor = System.Drawing.Color.White;
			this.udtG.DisplayLayout.Appearance = appearance2;
			ultraGridBand1.AddButtonCaption = "Cliente";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 80;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 75;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.Caption = "Recepcion";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.Caption = "Vencimiento";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 80;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance3;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.Caption = "Saldo Factura";
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "R.U.C.";
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 150;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 200;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance4;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Valor Pagado";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 100;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance5;
			ultraGridColumn11.Header.Caption = "N. Pago";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 80;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Fecha Pago";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance6;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 12;
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
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance8;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings3.Appearance = appearance9;
			summarySettings3.DisplayFormat = "{0: #,##0.00}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3});
			ultraGridBand1.SummaryFooterCaption = "";
			this.udtG.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.udtG.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.udtG.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.udtG.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance10.BackColor = System.Drawing.Color.Transparent;
			this.udtG.DisplayLayout.Override.CardAreaAppearance = appearance10;
			this.udtG.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.udtG.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.FontData.BoldAsString = "True";
			appearance11.FontData.Name = "Arial";
			appearance11.FontData.SizeInPoints = 8.5F;
			appearance11.ForeColor = System.Drawing.Color.White;
			appearance11.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.udtG.DisplayLayout.Override.HeaderAppearance = appearance11;
			this.udtG.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udtG.DisplayLayout.Override.RowSelectorAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.udtG.DisplayLayout.Override.SelectedRowAppearance = appearance13;
			this.udtG.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.udtG.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.udtG.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.udtG.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.udtG.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.udtG.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.udtG.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.udtG.Location = new System.Drawing.Point(4, 47);
			this.udtG.Name = "udtG";
			this.udtG.Size = new System.Drawing.Size(992, 345);
			this.udtG.TabIndex = 644;
			this.udtG.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.udtG_KeyPress);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn12.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(System.Decimal);
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
																																 ultraDataColumn13});
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
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(904, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(88, 23);
			this.btnExportar.TabIndex = 870;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// frmPlanificacionPagoProveedor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1000, 390);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.udtG);
			this.Controls.Add(this.btnBusca);
			this.Controls.Add(this.lblHasta);
			this.Controls.Add(this.dtHasta);
			this.Name = "frmPlanificacionPagoProveedor";
			this.Text = "Planificacion Pago Proveedores";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPlanificacionPagoProveedor_KeyDown);
			this.Load += new System.EventHandler(this.frmPlanificacionPagoProveedor_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udtG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmPlanificacionPagoProveedor_Load(object sender, System.EventArgs e)
		{
			this.dtHasta.Value = DateTime.Today;
		}
		private void Aceptar()
		{
			if (this.udtG.Rows.Count == 0)
			{
				MessageBox.Show("No hay registros encontrados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.dtHasta.Focus();			
			}
			else if (this.udtG.ActiveRow == null)
			{
				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
			}
//			else if (this.udtG.ActiveRow.Cells["idCompra"].Value == System.DBNull.Value)
//			{
//				MessageBox.Show("Seleccione un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
//			}
			else this.DialogResult = DialogResult.OK;
		}

		private void btnBusca_Click(object sender, System.EventArgs e)
		{	
			string sSQL = string.Format("Exec ConsultaEce '{0}'", Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			this.udtG.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			if (this.udtG.Rows.Count > 0) this.udtG.Focus(); else this.dtHasta.Focus();
		}

		private void btnAcepta_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		
		}

		private void udtG_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) Aceptar();
		}

		private void frmPlanificacionPagoProveedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Escape)
				this.DialogResult = DialogResult.No;
		}

		private void dtHasta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode==Keys.Enter)
				this.btnBusca_Click(sender,e);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.udtG);
		}
	}
}
