using System.Data;
using System.Data.SqlClient;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteComparacionGastos.
	/// </summary>
	public class frmReporteComparacionGastos : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		DateTime  dtDesde ;
		DateTime dtHasta ;
		DateTime dtDesde2 ;
		DateTime dtHasta2 ;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugridPorcentaje;
		int idCuenta = 0;
		Decimal dTotalA;
		Decimal dTotalB;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label lblContadorB;
		string sGasto ;
		private System.Windows.Forms.Label lblDiferencia;
		private System.Windows.Forms.Label lblPorcentaje;
		private System.Windows.Forms.Button btnExcel; 
		string sPeriodoMes;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		bool bPeriodoMes;
		

		public frmReporteComparacionGastos(DateTime DTDesde , DateTime DTHasta , int IdCuenta  , DateTime DTDesde2 , DateTime DTHasta2, string SGasto , string SPeriodoMes, bool BPeriodoMes )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			 dtDesde = DTDesde;
       dtHasta = DTHasta;
			 dtDesde2 = DTDesde2;
		   dtHasta2 = DTHasta2;
		   idCuenta = IdCuenta;
			sGasto = SGasto;
			sPeriodoMes= SPeriodoMes;
			bPeriodoMes = BPeriodoMes ;
			 
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuentas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ir", 0);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteComparacionGastos));
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Maximum, null, "Diferencia", 4, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Diferencia", 4, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Maximum, null, "Porcentaje", 5, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Porcentaje", 5, true);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Diferencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Porcentaje");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ugridPorcentaje = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblContadorB = new System.Windows.Forms.Label();
			this.lblDiferencia = new System.Windows.Forms.Label();
			this.lblPorcentaje = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentaje)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraDataSource1
			// 
			ultraDataColumn4.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
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
			// ugridPorcentaje
			// 
			this.ugridPorcentaje.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugridPorcentaje.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugridPorcentaje.DataSource = this.ultraDataSource2;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugridPorcentaje.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 10;
			ultraGridColumn3.Format = "#,##0.00";
			ultraGridColumn3.Header.VisiblePosition = 4;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 90;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Width = 8;
			ultraGridColumn6.Format = "#,##0.00";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 8;
			ultraGridColumn7.Header.VisiblePosition = 2;
			ultraGridColumn7.Width = 200;
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			appearance3.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellButtonAppearance = appearance3;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 30;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9});
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance4;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings2.Appearance = appearance5;
			summarySettings2.DisplayFormat = "{0: #,##0.00}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ugridPorcentaje.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugridPorcentaje.DisplayLayout.GroupByBox.Hidden = true;
			appearance6.ForeColor = System.Drawing.Color.Black;
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentaje.DisplayLayout.Override.ActiveRowAppearance = appearance6;
			this.ugridPorcentaje.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridPorcentaje.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridPorcentaje.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance7.BackColor = System.Drawing.Color.Transparent;
			this.ugridPorcentaje.DisplayLayout.Override.CardAreaAppearance = appearance7;
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentaje.DisplayLayout.Override.GroupByRowAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 8.5F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridPorcentaje.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ugridPorcentaje.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance10.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance10.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.RowAlternateAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.ugridPorcentaje.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugridPorcentaje.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridPorcentaje.Location = new System.Drawing.Point(8, 152);
			this.ugridPorcentaje.Name = "ugridPorcentaje";
			this.ugridPorcentaje.Size = new System.Drawing.Size(536, 368);
			this.ugridPorcentaje.TabIndex = 660;
			this.ugridPorcentaje.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugridPorcentaje_ClickCellButton);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13});
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(552, 8);
			this.groupBox1.TabIndex = 946;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(16, 8);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 958;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblContadorB
			// 
			this.lblContadorB.AutoSize = true;
			this.lblContadorB.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContadorB.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContadorB.Location = new System.Drawing.Point(16, 40);
			this.lblContadorB.Name = "lblContadorB";
			this.lblContadorB.Size = new System.Drawing.Size(0, 17);
			this.lblContadorB.TabIndex = 959;
			this.lblContadorB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDiferencia
			// 
			this.lblDiferencia.AutoSize = true;
			this.lblDiferencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiferencia.ForeColor = System.Drawing.Color.Firebrick;
			this.lblDiferencia.Location = new System.Drawing.Point(16, 104);
			this.lblDiferencia.Name = "lblDiferencia";
			this.lblDiferencia.Size = new System.Drawing.Size(0, 17);
			this.lblDiferencia.TabIndex = 960;
			this.lblDiferencia.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblPorcentaje
			// 
			this.lblPorcentaje.AutoSize = true;
			this.lblPorcentaje.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPorcentaje.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPorcentaje.Location = new System.Drawing.Point(16, 72);
			this.lblPorcentaje.Name = "lblPorcentaje";
			this.lblPorcentaje.Size = new System.Drawing.Size(0, 17);
			this.lblPorcentaje.TabIndex = 961;
			this.lblPorcentaje.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(472, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(72, 24);
			this.btnExcel.TabIndex = 962;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// frmReporteComparacionGastos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(552, 526);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.lblPorcentaje);
			this.Controls.Add(this.lblDiferencia);
			this.Controls.Add(this.lblContadorB);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.ugridPorcentaje);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmReporteComparacionGastos";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteComparacionGastos_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteComparacionGastos_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentaje)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteComparacionGastos_Load(object sender, System.EventArgs e)
		{
			Decimal dValor = 0.00m ;
			string sPeriodo = "";
			Decimal dValorB = 0.00m ;
			string sPeriodoB = "";

			Decimal dPorcentaje = 0.00m;
			Decimal dDiferencia = 0.00m;

				

			#region consulta Periodo A

			string sSQLA = string.Format("Exec PorcentajeResumenGastos '{0}', '{1}', {2}, '{3}', '{4}',1",
				Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"),Convert.ToDateTime(dtHasta).ToString("yyyyMMdd"), idCuenta
				, Convert.ToDateTime(dtDesde2).ToString("yyyyMMdd"), Convert.ToDateTime(dtHasta2).ToString("yyyyMMdd"));
			SqlDataReader drDOa = Funcion.rEscalarSQL(cdsSeteoF,sSQLA,true);
			drDOa.Read();
			if(drDOa.HasRows)
			{
				dValor = drDOa.GetDecimal (0);
				sPeriodo = drDOa.GetString (1);
			}
			drDOa.Close();

      #endregion consulta Periodo A

      #region consulta Periodo B

			string sSQLB = string.Format("Exec PorcentajeResumenGastos '{0}', '{1}', {2}, '{3}', '{4}',2",
				Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"),Convert.ToDateTime(dtHasta).ToString("yyyyMMdd"), idCuenta
				, Convert.ToDateTime(dtDesde2).ToString("yyyyMMdd"), Convert.ToDateTime(dtHasta2).ToString("yyyyMMdd"));
			SqlDataReader drDOB = Funcion.rEscalarSQL(cdsSeteoF,sSQLB,true);
			drDOB.Read();
			if(drDOB.HasRows)
			{
				dValorB = drDOB.GetDecimal (0);
				sPeriodoB= drDOB.GetString (1);
			}
			drDOB.Close();
      
			#endregion consulta Periodo B


			this.Text = " * GRUPO DE GASTO : " + sGasto ;
			#region Por Periodo

			if (bPeriodoMes)
			{

				this.lblContador.Text =  " * PERIODO : " +  sPeriodoMes + "/" + sPeriodo + "* TOTAL : " + dValor ;
				this.lblContadorB.Text = " * PERIODO : " + sPeriodoMes + "/" + sPeriodoB + "* TOTAL : " + dValorB;
				dPorcentaje = Math.Round(((( dValor/dValorB)*100)-100),2);
				dDiferencia = Math.Round(( dValor - dValorB),2);
				this.lblPorcentaje.Text = "* PORCENTAJE :" + sPeriodoMes + "/ "+ sPeriodo + "-" + sPeriodoB + " "+ "TOTAL" + " "+dPorcentaje ;
				this.lblDiferencia.Text = "* DIFERENCIA $:" + sPeriodoMes + "/ "+ sPeriodo + "-" + sPeriodoB + " "+ "TOTAL" + " "+dDiferencia ;
			}
			else
			{
			  int iMes = Convert.ToInt32( ( Convert.ToDateTime(dtDesde).Month));
				int iMesB = Convert.ToInt32( ( Convert.ToDateTime(dtDesde2).Month));

				string sVerMes = "ENERO";
				string sVerMesB = "ENERO";

				#region sVerMes

				if (iMes == 2)  sVerMes = "FEBRERO";
				if (iMes == 3 )  sVerMes = "MARZO";
				if (iMes == 4 )  sVerMes = "ABRIL";
				if (iMes == 5 )  sVerMes = "MAYO";
				if (iMes == 6 )  sVerMes = "JUNIO";
				if (iMes == 7 )  sVerMes = "JULIO";
				if (iMes == 8 )  sVerMes = "AGOSTO";
				if (iMes == 9 )  sVerMes = "SEPTIEMBRE";
				if (iMes == 10 )  sVerMes = "OCTUBRE";
				if (iMes == 11)  sVerMes = "NOVIEMBRE";
				if (iMes == 12)  sVerMes = "DICIEMBRE";

				#endregion sVerMes

				#region sVerMesB

				if (iMesB == 2)  sVerMesB = "FEBRERO";
				if (iMesB == 3 )  sVerMesB = "MARZO";
				if (iMesB == 4 )  sVerMesB = "ABRIL";
				if (iMesB == 5 )  sVerMesB = "MAYO";
				if (iMesB == 6 )  sVerMesB = "JUNIO";
				if (iMesB == 7 )  sVerMesB = "JULIO";
				if (iMesB == 8 )  sVerMesB = "AGOSTO";
				if (iMesB == 9 )  sVerMesB = "SEPTIEMBRE";
				if (iMesB == 10 )  sVerMesB = "OCTUBRE";
				if (iMesB == 11)  sVerMesB = "NOVIEMBRE";
				if (iMesB == 12)  sVerMesB = "DICIEMBRE";

				#endregion sVerMesB

				this.lblContador.Text =  " * PERIODO : " +  sVerMes + "/" + sPeriodo + "* TOTAL : " + dValor ;
				this.lblContadorB.Text = " * PERIODO : " + sVerMesB + "/" + sPeriodoB + "* TOTAL : " + dValorB;
				dPorcentaje = Math.Round(((( dValor/dValorB)*100)-100),2);
				dDiferencia = Math.Round(( dValor - dValorB),2);
				this.lblPorcentaje.Text = "* PORCENTAJE : " + sVerMes + "-"+ sPeriodo + "/"+ sVerMesB + "-" + sPeriodoB + " "+ "TOTAL" + " "+dPorcentaje ;
				this.lblDiferencia.Text = "* DIFERENCIA $: " + sVerMes + "-"+ sPeriodo + "/"+ sVerMesB + "-" + sPeriodoB + " "+  "TOTAL" + " "+dDiferencia ;

			
			}
			#endregion Por Periodo

			string sSQL = string.Format("Exec PorcentajeResumenGastos '{0}', '{1}', {2}, '{3}', '{4}'",
				Convert.ToDateTime(dtDesde).ToString("yyyyMMdd"),Convert.ToDateTime(dtHasta).ToString("yyyyMMdd"), idCuenta
				, Convert.ToDateTime(dtDesde2).ToString("yyyyMMdd"), Convert.ToDateTime(dtHasta2).ToString("yyyyMMdd"));
			this.ugridPorcentaje.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.ugridPorcentaje);

			this.ugridPorcentaje.DisplayLayout.Bands[0].SortedColumns.Add("Cuenta", true, true);
			this.ugridPorcentaje.Rows.CollapseAll(false);

		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
				FuncionesProcedimientos.ExportaExcel(this.ugridPorcentaje);
		}

		private void frmReporteComparacionGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void ugridPorcentaje_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			string sColumna = e.Cell.Column.ToString();
//			string sPeriodo = sColumna.Substring(3).ToUpper();
			
//			if (e.Cell.Row.Cells[sPeriodo].Value == System.DBNull.Value)
//			{
//				MessageBox.Show(string.Format("No hay gastos registrado en el periodo {0}", sPeriodo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}
			int idCuentaAgrupa = 0;
			int iPeriodo = 0;
			int idcuenta = 0;
			if((int)e.Cell.Row.Cells["Tipo"].Value == 2)
			{
				
				idCuentaAgrupa = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idCuentaAgrupa from Cuenta where Codigo = '{0}' and Descripcion = '{1}'",
						e.Cell.Row.Cells["Codigo"].Value.ToString(), e.Cell.Row.Cells["Descripcion"].Value.ToString()));	

				idcuenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idcuenta from Cuenta where Codigo = '{0}' and Descripcion = '{1}'",
					e.Cell.Row.Cells["Codigo"].Value.ToString(), e.Cell.Row.Cells["Descripcion"].Value.ToString()));	

				 iPeriodo = Convert.ToInt32( ( Convert.ToDateTime(dtDesde).ToString("yyyy")));
				frmPresupuestoResumenGastos prg = new frmPresupuestoResumenGastos(dtDesde, dtHasta, idCuentaAgrupa, e.Cell.Row.Cells["Periodo"].Value.ToString(), 
					e.Cell.Row.Cells["Descripcion"].Value.ToString(), idcuenta);
				prg.ShowDialog();
			}
			else
			{
				idCuentaAgrupa = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idCuentaAgrupa from Cuenta where Codigo = '{0}' and Descripcion = '{1}'",
					e.Cell.Row.Cells["Codigo"].Value.ToString(), e.Cell.Row.Cells["Descripcion"].Value.ToString()));

				idcuenta = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idcuenta from Cuenta where Codigo = '{0}' and Descripcion = '{1}'",
					e.Cell.Row.Cells["Codigo"].Value.ToString(), e.Cell.Row.Cells["Descripcion"].Value.ToString()));	

				 iPeriodo = Convert.ToInt32( ( Convert.ToDateTime(dtDesde2).ToString("yyyy")));
				frmPresupuestoResumenGastos prg = new frmPresupuestoResumenGastos(dtDesde2, dtHasta2, idCuentaAgrupa, e.Cell.Row.Cells["Periodo"].Value.ToString(), 
					e.Cell.Row.Cells["Descripcion"].Value.ToString(), idcuenta);
				prg.ShowDialog();
			}

//			this.MostrarDetalle(e.Cell.Column.ToString(), (int)e.Cell.Row.Cells["idCuentaAgrupa"].Value, e.Cell.Row.Cells["Codigo"].Value.ToString(), 
//				e.Cell.Row.Cells["Cuenta"].Value.ToString(), iPeriodo);
		}
//		private void MostrarDetalle(string sColumna, int idCuentaAgrupa, string sPeriodo, string sGasto, int IPeriodo)
//		{			
//			string sMes = "01/";
//			if (sColumna == "btnFebrero") sMes = "02/";
//			if (sColumna == "btnMarzo") sMes = "03/";
//			if (sColumna == "btnAbril") sMes = "04/";
//			if (sColumna == "btnMayo") sMes = "05/";
//			if (sColumna == "btnJunio") sMes = "06/";
//			if (sColumna == "btnJulio") sMes = "07/";
//			if (sColumna == "btnAgosto") sMes = "08/";
//			if (sColumna == "btnSeptiembre") sMes = "09/";
//			if (sColumna == "btnOctubre") sMes = "10/";
//			if (sColumna == "btnNoviembre") sMes = "11/";
//			if (sColumna == "btnDiciembre") sMes = "12/";
////			
////			DateTime dtDesde = DateTime.Parse("01/" + sMes + IPeriodo.ToString());
////			DateTime dtHasta = DateTime.Parse(DateTime.DaysInMonth(dtDesde.Date.Year, dtDesde.Month).ToString() + "/" + dtDesde.Month.ToString() + "/" + dtDesde.Year.ToString());
//			frmPresupuestoResumenGastos prg = new frmPresupuestoResumenGastos(dtDesde, dtHasta, idCuentaAgrupa, sPeriodo, sGasto);
//			prg.ShowDialog();
//
//		}

	}
}
