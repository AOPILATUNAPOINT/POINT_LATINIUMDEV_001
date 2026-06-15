using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;
using System.IO;



namespace Latinium
{
	/// <summary>
	/// Descripción breve de ReporteDig.
	/// </summary>
	public class ReporteDig : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbComprobante;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.Misc.UltraButton btLeer;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private C1.Data.C1DataSet cdsArticulo;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaH;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaD;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTipo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReporteDig()
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
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			this.cmbComprobante = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.btLeer = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFechaH = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaD = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.chkTipo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbComprobante
			// 
			this.cmbComprobante.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbComprobante.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbComprobante.DisplayMember = "";
			this.cmbComprobante.Location = new System.Drawing.Point(576, 8);
			this.cmbComprobante.Name = "cmbComprobante";
			this.cmbComprobante.Size = new System.Drawing.Size(128, 21);
			this.cmbComprobante.TabIndex = 82;
			this.cmbComprobante.ValueMember = "";
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(432, 56);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(96, 24);
			this.btExcel.TabIndex = 85;
			this.btExcel.Text = "Imprimir Datos";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btLeer
			// 
			this.btLeer.Location = new System.Drawing.Point(304, 56);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(96, 24);
			this.btLeer.TabIndex = 84;
			this.btLeer.Text = "Visualizar Datos";
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
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(704, 224);
			this.ultraGrid1.TabIndex = 86;
			this.ultraGrid1.MouseDown += new System.Windows.Forms.MouseEventHandler(this.ultraGrid1_MouseDown);
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsDigitalizacion";
			this.cdsArticulo.FillOnRequest = false;
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(224, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 24);
			this.label3.TabIndex = 90;
			this.label3.Text = "Fecha Hasta";
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(80, 24);
			this.label2.TabIndex = 89;
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
			this.cmbFechaH.Location = new System.Drawing.Point(312, 8);
			this.cmbFechaH.Name = "cmbFechaH";
			this.cmbFechaH.NonAutoSizeHeight = 23;
			this.cmbFechaH.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaH.SpinButtonsVisible = true;
			this.cmbFechaH.TabIndex = 88;
			this.cmbFechaH.Value = new System.DateTime(2011, 11, 28, 0, 0, 0, 0);
			// 
			// cmbFechaD
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFechaD.Appearance = appearance8;
			this.cmbFechaD.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaD.DateButtons.Add(dateButton2);
			this.cmbFechaD.Format = "dd/MMM/yyyy";
			this.cmbFechaD.Location = new System.Drawing.Point(96, 8);
			this.cmbFechaD.Name = "cmbFechaD";
			this.cmbFechaD.NonAutoSizeHeight = 23;
			this.cmbFechaD.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaD.SpinButtonsVisible = true;
			this.cmbFechaD.TabIndex = 87;
			this.cmbFechaD.Value = new System.DateTime(2011, 11, 28, 0, 0, 0, 0);
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(448, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(128, 24);
			this.label1.TabIndex = 91;
			this.label1.Text = "Transacción Contable";
			// 
			// chkTipo
			// 
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Microsoft Sans Serif";
			this.chkTipo.Appearance = appearance9;
			this.chkTipo.Checked = true;
			this.chkTipo.CheckState = System.Windows.Forms.CheckState.Checked;
			appearance10.FontData.BoldAsString = "True";
			this.chkTipo.IndeterminateAppearance = appearance10;
			this.chkTipo.Location = new System.Drawing.Point(16, 56);
			this.chkTipo.Name = "chkTipo";
			this.chkTipo.Size = new System.Drawing.Size(272, 20);
			this.chkTipo.TabIndex = 92;
			this.chkTipo.Text = "Comprobantes Obligatorios No Digitalizados";
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
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			this.errorProvider.DataMember = "";
			// 
			// ultraButton1
			// 
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Button3D;
			this.ultraButton1.Location = new System.Drawing.Point(576, 56);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(72, 23);
			this.ultraButton1.TabIndex = 93;
			this.ultraButton1.Text = "Cerrar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ReporteDig
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(720, 365);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.chkTipo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cmbFechaH);
			this.Controls.Add(this.cmbFechaD);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btLeer);
			this.Controls.Add(this.cmbComprobante);
			this.Name = "ReporteDig";
			this.Text = "Reporte Documentos Digitalizados";
			this.Load += new System.EventHandler(this.ReporteDig_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaH)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();
		int obligatorio;
		private void ReporteDig_Load(object sender, System.EventArgs e)
		{
			string stSelect="Select * from P_DigTipoComprobante order by Nombre";
			cmbComprobante.DataSource=Funcion.dvProcedimiento(cdsArticulo, stSelect);
			cmbComprobante.DisplayMember = "Nombre";
			cmbComprobante.ValueMember = "idTipoComprobante";
			cmbComprobante.Rows.Band.Columns["idTipoComprobante"].Hidden=true;
			
		}

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btLeer_Click(object sender, System.EventArgs e)
		{
			//	string tipo;
		//	bool bOk = true;
			DateTime dtDesde = (DateTime) cmbFechaD.Value;
			DateTime dtHasta = (DateTime) cmbFechaH.Value;

			
			string tipo ="";
			try		
			{
				tipo=cmbComprobante.Value.ToString();
						errorProvider.SetError(cmbComprobante, "");
			}
			catch(Exception ex)
			{
				errorProvider.SetError(cmbComprobante, "Ingrese Tipo Transacción Contable");
					MessageBox.Show(ex.Message, "" );
				//bOk = false;
				return;
				}
			if (!chkTipo.Checked) 
			{
			obligatorio=0;
			}
			else
			{
			obligatorio=1;
			}

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec ReporteDigitalizacion '"+dtDesde.ToString("yyyyMMdd")+"', '"+dtHasta.ToString("yyyyMMdd")+"', "+tipo+", "+obligatorio;
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
			ultraGrid1.Rows.Band.Columns["Interno"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["Tipo"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["Obligatorio"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["id"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["dfecha"].Hidden=true;
			ultraGrid1.Rows.Band.Columns["Ruta"].Hidden=true;
		//	string stFormato = stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			ultraGrid1.Rows.Band.Columns["Total"].Format = "n2"; 
		//	ultraGrid1.Rows.Band.Columns["Obligatorio"].CellActivation=Infragistics.Win.UltraWinGrid.Activation.NoEdit;

		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{

			DateTime dtDesde = (DateTime) cmbFechaD.Value;
			DateTime dtHasta = (DateTime) cmbFechaH.Value;

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName =  "@FechaDesde";
			discreteVal1.Value = dtDesde.ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			ParameterField paramField2 = new ParameterField ();
			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@FechaHasta";
			discreteVal2.Value = dtHasta.ToString("yyyy-MM-dd");
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterField paramField3 = new ParameterField ();
			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			paramField3.ParameterFieldName = "@idTipoFactura";
			discreteVal3.Value = (int) cmbComprobante.Value;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			ParameterField paramField4 = new ParameterField ();
			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
			paramField4.ParameterFieldName = "@pendiente";
			if (chkTipo.Checked)
				discreteVal4.Value = (int) 1;
			else
				discreteVal4.Value = (int) 0;
			paramField4.CurrentValues.Add (discreteVal4);
			paramFields.Add (paramField4);

			string stReporte = "ReporteDigitalizacion.rpt";
			Reporte miReporte = new Reporte(stReporte, "");
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Show();					

		}

		private void ultraGrid1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			
			int IdAsiento=(int)ultraGrid1.ActiveRow.Cells["Interno"].Value;

  		int Tipo=(int)ultraGrid1.ActiveRow.Cells["Tipo"].Value;	
			if(Tipo==2||Tipo==3)
				{
				string stAsiento= IdAsiento.ToString().Substring(1);
				int asiento = System.Convert.ToInt32(stAsiento); 
					Cursor = Cursors.WaitCursor;
					Asiento miAsiento = new Asiento(asiento);
					miAsiento.MdiParent = MdiParent;
					miAsiento.Show();
					Cursor = Cursors.Default;
				}
			else if(Tipo==10)

			{
					string stAsiento= IdAsiento.ToString().Substring(2);
					int asiento = System.Convert.ToInt32(stAsiento); 
					int IdCompra = asiento*(-1);
//					Compra miCompra = new Compra(IdCompra);
//					miCompra.MdiParent = MdiParent;
//					miCompra.Show();
			}

		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
	}
}
