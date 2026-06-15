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
	/// Descripción breve de ReporteGuiasGruas.
	/// </summary>
	public class ReporteGuiasGruas : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btLeer;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaH;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaD;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipoGuia;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReporteGuiasGruas()
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
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFechaH = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaD = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbTipoGuia = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoGuia)).BeginInit();
			this.SuspendLayout();
			// 
			// btLeer
			// 
			this.btLeer.Location = new System.Drawing.Point(296, 16);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(88, 24);
			this.btLeer.TabIndex = 2;
			this.btLeer.Text = "Cargar Datos";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			appearance6.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance6;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(416, 224);
			this.ultraGrid1.TabIndex = 15;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(296, 56);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(96, 24);
			this.btExcel.TabIndex = 16;
			this.btExcel.Text = "Exportar  Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// ultraGridExcelExporter1
			// 
			this.ultraGridExcelExporter1.BeginExport += new Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventHandler(this.ultraGridExcelExporter1_BeginExport);
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(16, 40);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 24);
			this.label3.TabIndex = 44;
			this.label3.Text = "Fecha Hasta";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(16, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 24);
			this.label2.TabIndex = 43;
			this.label2.Text = "Fecha Desde";
			// 
			// cmbFechaH
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaH.Appearance = appearance7;
			this.cmbFechaH.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaH.DateButtons.Add(dateButton1);
			this.cmbFechaH.Format = "dd/MMM/yyyy";
			this.cmbFechaH.Location = new System.Drawing.Point(104, 40);
			this.cmbFechaH.Name = "cmbFechaH";
			this.cmbFechaH.NonAutoSizeHeight = 23;
			this.cmbFechaH.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaH.SpinButtonsVisible = true;
			this.cmbFechaH.TabIndex = 42;
			this.cmbFechaH.Value = new System.DateTime(2009, 10, 26, 0, 0, 0, 0);
			// 
			// cmbFechaD
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaD.Appearance = appearance8;
			this.cmbFechaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaD.DateButtons.Add(dateButton2);
			this.cmbFechaD.Format = "dd/MMM/yyyy";
			this.cmbFechaD.Location = new System.Drawing.Point(104, 8);
			this.cmbFechaD.Name = "cmbFechaD";
			this.cmbFechaD.NonAutoSizeHeight = 23;
			this.cmbFechaD.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaD.SpinButtonsVisible = true;
			this.cmbFechaD.TabIndex = 41;
			this.cmbFechaD.Value = new System.DateTime(2009, 10, 26, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 80);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(80, 24);
			this.label1.TabIndex = 46;
			this.label1.Text = "Tipo Guía";
			// 
			// cmbTipoGuia
			// 
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Guias Gruas";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Guias Compras";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Guias Ventas";
			this.cmbTipoGuia.Items.Add(valueListItem1);
			this.cmbTipoGuia.Items.Add(valueListItem2);
			this.cmbTipoGuia.Items.Add(valueListItem3);
			this.cmbTipoGuia.Location = new System.Drawing.Point(104, 80);
			this.cmbTipoGuia.Name = "cmbTipoGuia";
			this.cmbTipoGuia.TabIndex = 47;
			// 
			// ReporteGuiasGruas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(432, 365);
			this.Controls.Add(this.cmbTipoGuia);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbFechaH);
			this.Controls.Add(this.cmbFechaD);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btLeer);
			this.Name = "ReporteGuiasGruas";
			this.Text = "Reporte de Guías";
			this.Load += new System.EventHandler(this.ReporteGuiasGruas_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoGuia)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();

		private void btLeer_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde = (DateTime) cmbFechaD.Value;
			DateTime dtHasta = (DateTime) cmbFechaH.Value;									

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ReporteGruas '"+dtDesde.ToString("yyyyMMdd")+"', '"+			dtHasta.ToString("yyyyMMdd")+"', "+cmbTipoGuia.Value;
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetG.Reset();
			miDataSetG.Clear();
			try
			{
				myData.Fill(miDataSetG);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}
			dvG = miDataSetG.Tables[0].DefaultView;
			ultraGrid1.DataSource = dvG;		
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "ReporteGuiasR.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "ReporteGuiasR.xls");
		}

		private void ultraGridExcelExporter1_BeginExport(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventArgs e)
		{
		
		}

		private void ReporteGuiasGruas_Load(object sender, System.EventArgs e)
		{
			cmbTipoGuia.Value = 0;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
