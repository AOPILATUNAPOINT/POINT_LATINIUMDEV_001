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
	/// Descripción breve de Busqueda.
	/// </summary>
	public class RequisicionListado : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTexto;
		private C1.Data.C1DataSet cdsGrupo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNumero;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCampo;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnHasta;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTablas;
		private Infragistics.Win.Misc.UltraButton btReq;
		private C1.Data.C1DataView cdvTipo;
		private C1.Data.C1DataSet cdsGrupoArt;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;

		string stTabla = "Articulo";
		public RequisicionListado()
		{
			InitializeComponent();
		}

		public RequisicionListado(string StTabla)
		{
			InitializeComponent();
			stTabla = StTabla;
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
			this.cmbCampo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtTexto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.spnNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.spnHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.cmbTablas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btReq = new Infragistics.Win.Misc.UltraButton();
			this.cdvTipo = new C1.Data.C1DataView();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cmbCampo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTexto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTablas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCampo
			// 
			this.cmbCampo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCampo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCampo.DisplayMember = "";
			this.cmbCampo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCampo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCampo.Location = new System.Drawing.Point(448, 24);
			this.cmbCampo.Name = "cmbCampo";
			this.cmbCampo.Size = new System.Drawing.Size(104, 21);
			this.cmbCampo.TabIndex = 0;
			this.cmbCampo.ValueMember = "";
			this.cmbCampo.Visible = false;
			this.cmbCampo.ValueChanged += new System.EventHandler(this.cmbCampo_ValueChanged);
			// 
			// txtTexto
			// 
			this.txtTexto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTexto.Location = new System.Drawing.Point(160, 32);
			this.txtTexto.Name = "txtTexto";
			this.txtTexto.Size = new System.Drawing.Size(0, 21);
			this.txtTexto.TabIndex = 1;
			// 
			// cdsGrupo
			// 
			this.cdsGrupo.BindingContextCtrl = this;
			this.cdsGrupo.DataLibrary = "LibFacturacion";
			this.cdsGrupo.DataLibraryUrl = "";
			this.cdsGrupo.DataSetDef = "dsSeteoF";
			this.cdsGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsGrupo.Name = "cdsGrupo";
			this.cdsGrupo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsGrupo.SchemaDef = null;
			this.cdsGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// spnNumero
			// 
			this.spnNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnNumero.FormatString = "#,##0.00";
			this.spnNumero.Location = new System.Drawing.Point(448, 0);
			this.spnNumero.Name = "spnNumero";
			this.spnNumero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnNumero.PromptChar = ' ';
			this.spnNumero.Size = new System.Drawing.Size(88, 21);
			this.spnNumero.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnNumero.TabIndex = 3;
			this.spnNumero.Visible = false;
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(448, 0);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 4;
			this.cmbFecha.Value = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbFecha.Visible = false;
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(24, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(64, 24);
			this.btBuscar.TabIndex = 5;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(448, 24);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(104, 21);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 7;
			this.cmbHasta.Value = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbHasta.Visible = false;
			// 
			// spnHasta
			// 
			this.spnHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnHasta.FormatString = "#,##0.00";
			this.spnHasta.Location = new System.Drawing.Point(456, 0);
			this.spnHasta.Name = "spnHasta";
			this.spnHasta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnHasta.PromptChar = ' ';
			this.spnHasta.Size = new System.Drawing.Size(88, 21);
			this.spnHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnHasta.TabIndex = 8;
			this.spnHasta.Visible = false;
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(216, 8);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(64, 24);
			this.btExcel.TabIndex = 10;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Location = new System.Drawing.Point(96, 8);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(64, 24);
			this.btGrabar.TabIndex = 11;
			this.btGrabar.Tag = "";
			this.btGrabar.Text = "&Aprobar";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// cmbTablas
			// 
			this.cmbTablas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTablas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTablas.DisplayMember = "";
			this.cmbTablas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTablas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTablas.Location = new System.Drawing.Point(448, 0);
			this.cmbTablas.Name = "cmbTablas";
			this.cmbTablas.Size = new System.Drawing.Size(104, 21);
			this.cmbTablas.TabIndex = 9;
			this.cmbTablas.ValueMember = "";
			this.cmbTablas.Visible = false;
			this.cmbTablas.ValueChanged += new System.EventHandler(this.cmbTablas_ValueChanged);
			this.cmbTablas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTablas_InitializeLayout);
			// 
			// btReq
			// 
			this.btReq.Location = new System.Drawing.Point(168, 8);
			this.btReq.Name = "btReq";
			this.btReq.Size = new System.Drawing.Size(40, 24);
			this.btReq.TabIndex = 13;
			this.btReq.Text = "&Req";
			this.btReq.Click += new System.EventHandler(this.btReq_Click);
			// 
			// cdvTipo
			// 
			this.cdvTipo.BindingContextCtrl = this;
			this.cdvTipo.DataSet = this.cdsGrupoArt;
			this.cdvTipo.TableName = "";
			this.cdvTipo.TableViewName = "ArticuloTipo";
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
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
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
			this.ultraGrid1.Location = new System.Drawing.Point(16, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(616, 208);
			this.ultraGrid1.TabIndex = 15;
			// 
			// RequisicionListado
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(648, 278);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btReq);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.cmbTablas);
			this.Controls.Add(this.spnHasta);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.spnNumero);
			this.Controls.Add(this.txtTexto);
			this.Controls.Add(this.cmbCampo);
			this.Name = "RequisicionListado";
			this.Text = "Lista Requisiciones No Aprobadas";
			this.Load += new System.EventHandler(this.Busqueda_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCampo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTexto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTablas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private string FiltroCampo()
		{
			if (cmbCampo.Text.Length == 0) return "";
			string stFiltroAntes = "%";
			string stFiltroDespues = "%";
//			int iTipo = (int) optBusca.Value;
//			if (iTipo == 0) stFiltroAntes = "";
//			if (iTipo == 1) 
//			{
//				stFiltroAntes = "";
//				stFiltroDespues = "";
//			}
			string stFiltro = " Where " +  stTabla + "." + cmbCampo.Text;
			if (txtTexto.Visible) 
				stFiltro += " Like '" + stFiltroAntes + txtTexto.Text.ToString() + stFiltroDespues + "'";
			if (spnNumero.Visible)
			{
				stFiltro += " Between " + spnNumero.Text.ToString() + " And " + spnHasta.Value.ToString();
			}
			if (cmbFecha.Visible) 
			{
				DateTime dtFecha = (DateTime) cmbFecha.Value;
				DateTime dtHasta = (DateTime) cmbHasta.Value;
				stFiltro += " Between '" + dtFecha.ToString("yyyyMMdd") 
					+ "' And '" + dtHasta.ToString("yyyyMMdd") + "'";
			}
			return stFiltro;
		}

		private void cmbCampo_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbCampo.ActiveRow == null) return;
			int iTipo = (int)	cmbCampo.ActiveRow.Cells["TipoCampo"].Value;
			txtTexto.Visible = false;
			spnNumero.Visible = false;
			spnHasta.Visible = false;
			cmbFecha.Visible =false;
			cmbHasta.Visible =false;
			switch(iTipo)
			{
				case 1:

					break;
				case 2:
				case 3:
					spnNumero.Left = txtTexto.Left;
					spnNumero.Top = txtTexto.Top;
					spnNumero.Visible = true;
					spnHasta.Left = spnNumero.Right + 10;
					spnHasta.Top = spnNumero.Top;
					spnHasta.Visible = true;
					spnNumero.SelectAll();
					break;
				case 4:
					cmbFecha.Left = txtTexto.Left;
					cmbFecha.Top = txtTexto.Top;
					cmbFecha.Visible = true;
					cmbFecha.SelectAll();
					cmbHasta.Left = cmbFecha.Right + 10;
					cmbHasta.Top = txtTexto.Top;
					cmbHasta.Visible =true;
					break;
				case 5:
					txtTexto.Visible = true;
					txtTexto.SelectAll();
					break;
			}
		}

		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();

		private void LlenaCampos(string StTabla)
		{
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			SqlDataAdapter myData = new SqlDataAdapter(
				"EXEC TablaCampoFiltro '" + StTabla + "', 1", strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " );
				return;
			}

			dv = miDataSet.Tables[0].DefaultView;
			cmbCampo.DataSource = dv;
			cmbCampo.DisplayLayout.Bands[0].Columns["TipoCampo"].Hidden = true;
			cmbCampo.DataBind();
		}
		private Acceso miAcceso;

		private void Busqueda_Load(object sender, System.EventArgs e)
		{
			miAcceso =new Acceso("050A7", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			miDataSetG.Reset();
			miDataSetG.Clear();
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
			string stProcedimiento = "Select Distinct(Tabla) As Tabla From TablaCampo";
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			
			miDataSetT.Reset();
			miDataSetT.Clear();
			try
			{
				myData.Fill(miDataSetT);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir Tablas de TablaCompra " );
				return;
			}

			dvT = miDataSetT.Tables[0].DefaultView;
			cmbTablas.DataSource = dvT;
			cmbTablas.DataBind();
			cmbTablas.Value = "Articulo";
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
//			if (!miAcceso.Buscar)
//			{
//				MessageBox.Show("No tiene acceso a Buscar " + this.Text.Trim(),
//					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				return;
//			}

			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec RequisicionListado " ;
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

			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			#endregion Visibilidad
		}

		private void btBuscaInst_Click(object sender, System.EventArgs e)
		{
			BusquedaInstruccion miInstr = new BusquedaInstruccion();
			miInstr.MdiParent = this.MdiParent;
			miInstr.Show();
		}

		private void cmbTablas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTablas_ValueChanged(object sender, System.EventArgs e)
		{
			stTabla = cmbTablas.Text;
			LlenaCampos(stTabla);
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "ReqNoAprobadas.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "ReqNoAprobadas.xls");
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			if (!Funcion.Permiso("701", this.cdsSeguridad))
			{
				MessageBox.Show("No tiene acceso a dar Autorizacion",
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
			if (dr.Cells["idCompra"].Value != System.DBNull.Value && 
					  dr.Cells["Autoriza"].Value != System.DBNull.Value && 
					(bool)dr.Cells["Autoriza"].Value == true && 
					(bool)dr.Cells["Compra_Local"].Value == true)
				{
					int idCompra = (int) dr.Cells["idCompra"].Value;
					string stExec = "Exec RequisicionListadoGrabar " + idCompra.ToString() + ", 1";
					Funcion.EjecutaSQL(cdsGrupo, stExec, true);
				}
				if (dr.Cells["idCompra"].Value != System.DBNull.Value && 
					dr.Cells["Compra_Local"].Value != System.DBNull.Value && 
					(bool)dr.Cells["Autoriza"].Value == true && 
					(bool)dr.Cells["Compra_Local"].Value == false)
				{
					int idCompra = (int) dr.Cells["idCompra"].Value;
					string stExec = "Exec RequisicionListadoGrabar " + idCompra.ToString() + ", 0";
					Funcion.EjecutaSQL(cdsGrupo, stExec, true);
				}
			}
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec RequisicionListado " ;
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

			#region Visibilidad
			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompra"].Hidden = true;
			#endregion Visibilidad

			MessageBox.Show("Requisiciones Aprobadas");








		}

		private void btReq_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver la Requisicion", "Información");
				return;
			}

			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			if (IdCompra == 0)
			{
				MessageBox.Show("No existe una Reqisicion", "Información");
				return;
			}
			//			int TipoFactura = 36;
			Cursor = Cursors.WaitCursor;
//			Compra miCompra = new Compra((int) 36, IdCompra);
//			miCompra.Show();
			Cursor = Cursors.Default;
		}

		private void cmbTipo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
