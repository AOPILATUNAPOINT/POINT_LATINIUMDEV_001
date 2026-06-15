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
	/// Descripción breve de AprobacionReq.
	/// </summary>
	public class AprobacionReq : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRequesicion;
		private DevExpress.XtraEditors.SimpleButton btLeer;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsRequisicion;
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private DevExpress.XtraEditors.SimpleButton btTransformar;
		private C1.Data.C1DataView cdvTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoTipo;
		private C1.Data.C1DataSet cdsGrupoArt;
		private DevExpress.XtraEditors.SimpleButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AprobacionReq()
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aprobado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OrdenCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ArticuloClase", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.udsRequisicion = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsRequesicion = new C1.Data.C1DataSet();
			this.btLeer = new DevExpress.XtraEditors.SimpleButton();
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.btTransformar = new DevExpress.XtraEditors.SimpleButton();
			this.cdvTipo = new C1.Data.C1DataView();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkTodoTipo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btExcel = new DevExpress.XtraEditors.SimpleButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.udsRequisicion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRequesicion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// udsRequisicion
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Double);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(bool);
			ultraDataColumn10.DefaultValue = false;
			this.udsRequisicion.Band.Columns.AddRange(new object[] {
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
			// cdsRequesicion
			// 
			this.cdsRequesicion.BindingContextCtrl = this;
			this.cdsRequesicion.DataLibrary = "LibFacturacion";
			this.cdsRequesicion.DataLibraryUrl = "";
			this.cdsRequesicion.DataSetDef = "dsCuartos";
			this.cdsRequesicion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRequesicion.Name = "cdsRequesicion";
			this.cdsRequesicion.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsRequesicion.SchemaDef = null;
			this.cdsRequesicion.BeforeFill += new C1.Data.FillEventHandler(this.cdsRequesicion_BeforeFill);
			// 
			// btLeer
			// 
			this.btLeer.Location = new System.Drawing.Point(24, 16);
			this.btLeer.Name = "btLeer";
			this.btLeer.Size = new System.Drawing.Size(72, 24);
			this.btLeer.TabIndex = 1;
			this.btLeer.Text = "Leer";
			this.btLeer.Click += new System.EventHandler(this.btLeer_Click);
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(112, 16);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(72, 24);
			this.btGenerar.TabIndex = 2;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// btTransformar
			// 
			this.btTransformar.Location = new System.Drawing.Point(568, 16);
			this.btTransformar.Name = "btTransformar";
			this.btTransformar.Size = new System.Drawing.Size(80, 32);
			this.btTransformar.TabIndex = 3;
			this.btTransformar.Text = "Transformar";
			this.btTransformar.Visible = false;
			this.btTransformar.Click += new System.EventHandler(this.btTransformar_Click);
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsGrupoArt;
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "ArticuloClase";
			// 
			// cdsGrupoArt
			// 
			this.cdsGrupoArt.BindingContextCtrl = this;
			this.cdsGrupoArt.CaseSensitive = false;
			this.cdsGrupoArt.DataLibrary = "LibFacturacion";
			this.cdsGrupoArt.DataLibraryUrl = "";
			this.cdsGrupoArt.DataSetDef = "dsArticuloTabla";
			this.cdsGrupoArt.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupoArt.Name = "cdsGrupoArt";
			this.cdsGrupoArt.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupoArt.SchemaDef = null;
			// 
			// cmbTipo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance1;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.cdvTipo;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 38;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 141;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 39;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Clase";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(256, 16);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(160, 21);
			this.cmbTipo.TabIndex = 6;
			this.cmbTipo.ValueMember = "idClase";
			// 
			// chkTodoTipo
			// 
			this.chkTodoTipo.Location = new System.Drawing.Point(200, 16);
			this.chkTodoTipo.Name = "chkTodoTipo";
			this.chkTodoTipo.Size = new System.Drawing.Size(48, 16);
			this.chkTodoTipo.TabIndex = 7;
			this.chkTodoTipo.Text = "Tipo";
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(432, 16);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(64, 24);
			this.btExcel.TabIndex = 11;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance2.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance4;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			this.ultraGrid1.DisplayLayout.Override.RowFilterAction = Infragistics.Win.UltraWinGrid.RowFilterAction.HideFilteredOutRows;
			this.ultraGrid1.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			appearance7.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Override.SummaryValueAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(16, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(632, 208);
			this.ultraGrid1.TabIndex = 14;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(680, 0);
			this.label1.Name = "label1";
			this.label1.TabIndex = 15;
			this.label1.Text = "label1";
			// 
			// AprobacionReq
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(680, 301);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.chkTodoTipo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.btTransformar);
			this.Controls.Add(this.btGenerar);
			this.Controls.Add(this.btLeer);
			this.Name = "AprobacionReq";
			this.Text = "Aprobación  Requisición";
			this.Load += new System.EventHandler(this.AprobacionReq_Load);
			((System.ComponentModel.ISupportInitialize)(this.udsRequisicion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRequesicion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
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
			#region LLamo_Procedimiento

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			int idTipoGrupo = 0;
			if (chkTodoTipo.Checked) idTipoGrupo = (int) cmbTipo.Value;
			string stExec = "Exec RequisicionAprobacion " 
				+ idTipoGrupo.ToString();

			SqlDataAdapter myData = new SqlDataAdapter(stExec, strConn);
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

			#endregion Llamo_Procedimiento

			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			ultraGrid1.DisplayLayout.Bands[0].Columns["idDetCompra"].Hidden = true;

			#endregion Visibilidad
			MessageBox.Show("Movimientos Generados", "Información");
		}

		private void cdsRequesicion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRequesicion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea Generar Orden de compra de Aprobado", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			DateTime dtFecha = DateTime.Today;
			string stFactura = "Exec FacturaCrear '', '', '-1', '" + dtFecha.ToString("yyyyMMdd")
				+ "', 0, 0, 0, 0, 0, 2";
			int IdCompra = Funcion.iEscalarSQL(cdsRequesicion, stFactura, true);
			bool bExiste = false;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if ((bool) dr.Cells["Transformar"].Value == true &&
					dr.Cells["OrdenCompra"].Value.ToString().Trim().Length == 0)
				{
					bExiste = true;
					int idDet = (int) dr.Cells["idDetCompra"].Value;
					string stUpdate = "Exec RequisicionOrdenCompraLinea " + idDet.ToString()
						+ ", " + IdCompra.ToString();
					Funcion.EjecutaSQL(cdsRequesicion, stUpdate, true);
				}
			}
			if (!bExiste)
			{
				MessageBox.Show("No existe datos para Generar", "Informaciòn",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			string stDetalle = "Exec RequisicionGenera " + IdCompra.ToString();
			Funcion.EjecutaSQL(cdsRequesicion, stDetalle, true);
			MessageBox.Show("Orden de compra generada", "Información",
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btTransformar_Click(object sender, System.EventArgs e)
		{
			MessageBox.Show("Desea transformar en Compra", "Informar");

		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "ReqAprobadas.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "ReqAprobadas.xls");

		}

		private void AprobacionReq_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
