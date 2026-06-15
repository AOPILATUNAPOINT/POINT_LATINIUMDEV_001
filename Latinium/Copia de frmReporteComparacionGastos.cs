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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuentas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Valor", 3, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Valor", 3, true);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteComparacionGastos));
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ugridPorcentaje = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.lblContadorB = new System.Windows.Forms.Label();
			this.lblDiferencia = new System.Windows.Forms.Label();
			this.lblPorcentaje = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugridPorcentaje)).BeginInit();
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
			this.ugridPorcentaje.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugridPorcentaje.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 250;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.Caption = "Total";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 90;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance4;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "Total Gastos";
			this.ugridPorcentaje.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ugridPorcentaje.DisplayLayout.GroupByBox.Hidden = true;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentaje.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ugridPorcentaje.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugridPorcentaje.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugridPorcentaje.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ugridPorcentaje.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugridPorcentaje.DisplayLayout.Override.GroupByRowAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugridPorcentaje.DisplayLayout.Override.HeaderAppearance = appearance8;
			this.ugridPorcentaje.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugridPorcentaje.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ugridPorcentaje.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ugridPorcentaje.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugridPorcentaje.Location = new System.Drawing.Point(0, 152);
			this.ugridPorcentaje.Name = "ugridPorcentaje";
			this.ugridPorcentaje.Size = new System.Drawing.Size(600, 368);
			this.ugridPorcentaje.TabIndex = 660;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 136);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(600, 8);
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
			this.btnExcel.Location = new System.Drawing.Point(520, 8);
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
			this.ClientSize = new System.Drawing.Size(600, 526);
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
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugridPorcentaje);

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
	}
}
