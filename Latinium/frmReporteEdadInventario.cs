using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteEdadInventario.
	/// </summary>
	public class frmReporteEdadInventario : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button btnExportar;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private System.Windows.Forms.Button btnGenerar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteEdadInventario()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteEdadInventario));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultimo_Costo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Edad_Dias");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango_Dias");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Stock");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaMovimiento");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultimo_Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Edad_Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango_Dias");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Stock");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaMovimiento");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnExportar = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnGenerar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(8, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1005, 8);
			this.groupBox1.TabIndex = 385;
			this.groupBox1.TabStop = false;
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnExportar.Location = new System.Drawing.Point(928, 8);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 381;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 12);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 17);
			this.label5.TabIndex = 384;
			this.label5.Text = "Fecha";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 120;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 120;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Width = 120;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 3;
			ultraGridColumn5.Width = 250;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 75;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn7.CellAppearance = appearance2;
			ultraGridColumn7.Format = "#,##0.00";
			ultraGridColumn7.Header.Caption = "Ultimo Costo";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 100;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn8.CellAppearance = appearance3;
			ultraGridColumn8.Header.Caption = "Edad Dias";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 75;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance4;
			ultraGridColumn9.Header.Caption = "Rango Dias";
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 120;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.Caption = "Fecha Movimiento";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 110;
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
																										 ultraGridColumn11});
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance8;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8.5F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1000, 312);
			this.ultraGrid1.TabIndex = 382;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(System.Decimal);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.DateTime);
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
																																 ultraDataColumn11});
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
			// dtHasta
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance13;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Action = Infragistics.Win.UltraWinSchedule.DateButtonAction.AddToSelection;
			dateButton1.Caption = "Today";
			dateButton1.Type = Infragistics.Win.UltraWinSchedule.DateButtonType.Custom;
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(56, 10);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(104, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 380;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// btnGenerar
			// 
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGenerar.Location = new System.Drawing.Point(184, 8);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(75, 24);
			this.btnGenerar.TabIndex = 386;
			this.btnGenerar.Text = "Generar";
			this.btnGenerar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// frmReporteEdadInventario
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1016, 366);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.groupBox1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmReporteEdadInventario";
			this.Text = "Edad de Inventario";
			this.Load += new System.EventHandler(this.frmReporteEdadInventario_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteEdadInventario_Load(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("2006", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar al reporte Edad del Inventario", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
				
			FuncionesProcedimientos.RetornaFechaServidor(this.dtHasta, cdsSeteoF, false);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsSeteoF, "Truncate Table ReporteSerialesMgrInv");

			int iCantidad =
				Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select MAX(Bodega) From dbo.Bodega Where Inventario = 1 And '{0}' Between CONVERT(DATE, FechaApertura) And CONVERT(DATE, ISNULL(FechaCierre, '20301231'))", 
					Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));

			int iBodega = 1;

			while (iBodega <= iCantidad)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select COUNT(*) From dbo.Bodega Where Inventario = 1 And Bodega = {0} And '{1}' Between CONVERT(DATE, FechaApertura) And CONVERT(DATE, ISNULL(FechaCierre, '20301231'))", 
					iBodega, Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"))) == 1)
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec Mgr_ReportePorSerial {0}, '{1}'", iBodega, Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd")));

				iBodega++;
			}

			string sSQL = string.Format("Exec GeneraReporteEdadInventario '{0}'",
				Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"));
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}
	}
}
