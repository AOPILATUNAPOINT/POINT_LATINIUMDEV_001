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
	public class Busqueda : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTexto;
		private C1.Data.C1DataSet cdsGrupo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnNumero;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCampo;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btBuscaInst;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor spnHasta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTablas;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;

		string stTabla = "Articulo";
		public Busqueda()
		{
			InitializeComponent();
		}

		public Busqueda(string StTabla)
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
			this.cmbCampo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtTexto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsGrupo = new C1.Data.C1DataSet();
			this.spnNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btBuscaInst = new Infragistics.Win.Misc.UltraButton();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.spnHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbTablas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			((System.ComponentModel.ISupportInitialize)(this.cmbCampo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTexto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTablas)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbCampo
			// 
			this.cmbCampo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCampo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCampo.DisplayMember = "";
			this.cmbCampo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCampo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCampo.Location = new System.Drawing.Point(32, 32);
			this.cmbCampo.Name = "cmbCampo";
			this.cmbCampo.Size = new System.Drawing.Size(104, 21);
			this.cmbCampo.TabIndex = 0;
			this.cmbCampo.ValueMember = "";
			this.cmbCampo.ValueChanged += new System.EventHandler(this.cmbCampo_ValueChanged);
			// 
			// txtTexto
			// 
			this.txtTexto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTexto.Location = new System.Drawing.Point(160, 32);
			this.txtTexto.Name = "txtTexto";
			this.txtTexto.Size = new System.Drawing.Size(160, 21);
			this.txtTexto.TabIndex = 1;
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
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
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
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(24, 64);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(600, 192);
			this.ultraGrid1.TabIndex = 2;
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
			this.spnNumero.Location = new System.Drawing.Point(192, 0);
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
			this.cmbFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(192, 24);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 4;
			this.cmbFecha.Visible = false;
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(160, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(64, 24);
			this.btBuscar.TabIndex = 5;
			this.btBuscar.Text = "Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btBuscaInst
			// 
			this.btBuscaInst.Location = new System.Drawing.Point(232, 8);
			this.btBuscaInst.Name = "btBuscaInst";
			this.btBuscaInst.Size = new System.Drawing.Size(24, 24);
			this.btBuscaInst.TabIndex = 6;
			this.btBuscaInst.Text = "...";
			this.btBuscaInst.Click += new System.EventHandler(this.btBuscaInst_Click);
			// 
			// cmbHasta
			// 
			this.cmbHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(312, 24);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(104, 21);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 7;
			this.cmbHasta.Visible = false;
			// 
			// spnHasta
			// 
			this.spnHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnHasta.FormatString = "#,##0.00";
			this.spnHasta.Location = new System.Drawing.Point(312, 0);
			this.spnHasta.Name = "spnHasta";
			this.spnHasta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnHasta.PromptChar = ' ';
			this.spnHasta.Size = new System.Drawing.Size(88, 21);
			this.spnHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnHasta.TabIndex = 8;
			this.spnHasta.Visible = false;
			// 
			// cmbTablas
			// 
			this.cmbTablas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTablas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTablas.DisplayMember = "";
			this.cmbTablas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTablas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTablas.Location = new System.Drawing.Point(32, 8);
			this.cmbTablas.Name = "cmbTablas";
			this.cmbTablas.Size = new System.Drawing.Size(104, 21);
			this.cmbTablas.TabIndex = 9;
			this.cmbTablas.ValueMember = "";
			this.cmbTablas.ValueChanged += new System.EventHandler(this.cmbTablas_ValueChanged);
			this.cmbTablas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTablas_InitializeLayout);
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(272, 8);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(64, 24);
			this.btExcel.TabIndex = 10;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// Busqueda
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 278);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.cmbTablas);
			this.Controls.Add(this.spnHasta);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.btBuscaInst);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.spnNumero);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtTexto);
			this.Controls.Add(this.cmbCampo);
			this.Name = "Busqueda";
			this.Text = "Búsqueda";
			this.Load += new System.EventHandler(this.Busqueda_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCampo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTexto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTablas)).EndInit();
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

		private void Busqueda_Load(object sender, System.EventArgs e)
		{
//			LlenaCampos(stTabla);
//			cmbCampo.Value = "Articulo";

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
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stInstruccion = "Select Top 1 Instruccion From TablaBuscar Where Tabla = '"
				+ cmbTablas.Text + "'";
			string stProcedimiento = Funcion.sEscalarSQL(cdsGrupo, stInstruccion, true);
			stProcedimiento += FiltroCampo();
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
			string stClave = Funcion.sEscalarSQL(cdsGrupo, "Exec TablaCampoFiltro '" + stTabla + "', 3", true);
			if (stClave.Length == 0)
				MessageBox.Show("No existe clave principal", "Información");

			SqlDataReader dr = Funcion.rEscalarSQL(cdsGrupo, "Exec TablaCampoFiltro '" + stTabla + "', 4", true);
			while(dr.Read())
			{
				string stCampo = dr.GetString(0);
				int iTipo = dr.GetInt32(1);
				if (iTipo == 2)
				{
					ultraGrid1.DisplayLayout.Bands[0].Columns[stCampo].Format = "#,##0";
					ultraGrid1.DisplayLayout.Bands[0].Columns[stCampo].CellAppearance.TextHAlign =
						Infragistics.Win.HAlign.Right;
				}
				if (iTipo == 3)
				{
					ultraGrid1.DisplayLayout.Bands[0].Columns[stCampo].Format = "#,##0.00";
					ultraGrid1.DisplayLayout.Bands[0].Columns[stCampo].CellAppearance.TextHAlign =
						Infragistics.Win.HAlign.Right;
				}
				if (iTipo == 4)
					ultraGrid1.DisplayLayout.Bands[0].Columns[stCampo].Format = "dd/MMM/yyyy";
			}
			dr.Close();

			#region Visibilidad
			dr = Funcion.rEscalarSQL(cdsGrupo, "Exec TablaCampoFiltro '" + stTabla + "', 5", true);
			while(dr.Read())
			{
				string stCampo = dr.GetString(0);
				ultraGrid1.DisplayLayout.Bands[0].Columns[stCampo].Hidden = true;
			}
			dr.Close();
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
			ultraGridExcelExporter1.Export(ultraGrid1, "Busqueda.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "Busqueda.xls");
		}
	}
}
