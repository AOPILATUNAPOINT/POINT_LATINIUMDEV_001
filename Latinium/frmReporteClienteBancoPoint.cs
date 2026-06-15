using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteClienteBancoPoint.
	/// </summary>
	public class frmReporteClienteBancoPoint : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidad;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridReporte;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtCorte;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteClienteBancoPoint()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Saldo", 1, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Saldo", 1, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Saldo");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteClienteBancoPoint));
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbEntidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.uGridReporte = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.dtCorte = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCorte)).BeginInit();
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
			// cmbEntidad
			// 
			this.cmbEntidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbEntidad.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.Caption = "Cartera";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbEntidad.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbEntidad.DisplayMember = "Nombre";
			this.cmbEntidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidad.Location = new System.Drawing.Point(56, 40);
			this.cmbEntidad.MaxDropDownItems = 30;
			this.cmbEntidad.Name = "cmbEntidad";
			this.cmbEntidad.Size = new System.Drawing.Size(176, 22);
			this.cmbEntidad.TabIndex = 1001;
			this.cmbEntidad.ValueMember = "idEntidadFinanciera";
			this.cmbEntidad.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEntidad_InitializeLayout);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(42, 16);
			this.label2.TabIndex = 1002;
			this.label2.Text = "Cartera";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(-296, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1072, 8);
			this.groupBox1.TabIndex = 1008;
			this.groupBox1.TabStop = false;
			// 
			// uGridReporte
			// 
			this.uGridReporte.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridReporte.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridReporte.DisplayLayout.Appearance = appearance1;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 250;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance2;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 129;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			ultraGridBand2.Header.Caption = "";
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance3;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand2.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand2.SummaryFooterCaption = "";
			this.uGridReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridReporte.DisplayLayout.Override.ActiveRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.uGridReporte.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 8F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridReporte.DisplayLayout.Override.HeaderAppearance = appearance6;
			this.uGridReporte.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance7.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance7.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.RowAlternateAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridReporte.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.uGridReporte.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridReporte.Location = new System.Drawing.Point(8, 80);
			this.uGridReporte.Name = "uGridReporte";
			this.uGridReporte.Size = new System.Drawing.Size(416, 296);
			this.uGridReporte.TabIndex = 1009;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn2.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(32, 16);
			this.label3.TabIndex = 1012;
			this.label3.Text = "Corte";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtCorte
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtCorte.Appearance = appearance10;
			this.dtCorte.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtCorte.DateButtons.Add(dateButton1);
			this.dtCorte.Format = "dd/MM/yyyy";
			this.dtCorte.Location = new System.Drawing.Point(56, 8);
			this.dtCorte.Name = "dtCorte";
			this.dtCorte.NonAutoSizeHeight = 23;
			this.dtCorte.Size = new System.Drawing.Size(120, 21);
			this.dtCorte.SpinButtonsVisible = true;
			this.dtCorte.TabIndex = 1011;
			this.dtCorte.Value = ((object)(resources.GetObject("dtCorte.Value")));
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Enabled = false;
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(344, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 1010;
			this.btnExcel.Text = "&Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click_1);
			// 
			// btnVer
			// 
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(248, 40);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 1013;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// frmReporteClienteBancoPoint
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(432, 382);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.dtCorte);
			this.Controls.Add(this.uGridReporte);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbEntidad);
			this.Controls.Add(this.btnExcel);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReporteClienteBancoPoint";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Historial Cobranza Saldo Final";
			this.Load += new System.EventHandler(this.frmReporteClienteBancoPoint_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCorte)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void frmReporteClienteBancoPoint_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200616");

			if (!Funcion.Permiso("1557", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Cobranza por Edades", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			if (miAcceso.BExportar) this.btnExcel.Enabled =  true;

			this.dtCorte.Value = DateTime.Today;
			this.cmbEntidad.DataSource = 
				Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre, Codigo From Cre_EntidadFinanciera Where idEntidadFinanciera In (1, 3, 4, 14, 15, 16)");
		}

		private void cmbEntidad_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbEntidad, "Seleccione una Entidad")) return;

			int idEntidadFinanciera = (int)this.cmbEntidad.Value;

			DateTime dtCorte = (DateTime) this.dtCorte.Value;

			string sSQL = string.Format("Exec CXCReporteCobranzaPorEdadSaldoFinal '{0}', {1}, {2}", 
				dtCorte.ToString("yyyyMMdd"), 1, idEntidadFinanciera);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.uGridReporte);
		}

		private void btnExcel_Click_1(object sender, System.EventArgs e)
		{
		FuncionesProcedimientos.ExportaExcel(this.uGridReporte);
		}
	}
}
