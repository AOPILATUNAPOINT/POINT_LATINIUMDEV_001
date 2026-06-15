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
	public class ResumenCostosDI : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTexto;
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
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTablas;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoTipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private C1.Data.C1DataSet cdsGrupoArt;
		private C1.Data.C1DataView cdvTipo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private C1.Data.C1DataSet cdsCentroCosto;

		string stTabla = "Articulo";
		public ResumenCostosDI()
		{
			InitializeComponent();
		}

		public ResumenCostosDI(string StTabla)
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			this.cmbCampo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtTexto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsCentroCosto = new C1.Data.C1DataSet();
			this.spnNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.spnHasta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.cmbTablas = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkTodoTipo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvTipo = new C1.Data.C1DataView();
			this.cdsGrupoArt = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			((System.ComponentModel.ISupportInitialize)(this.cmbCampo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTexto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTablas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).BeginInit();
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
			this.cmbCampo.Location = new System.Drawing.Point(520, 40);
			this.cmbCampo.Name = "cmbCampo";
			this.cmbCampo.Size = new System.Drawing.Size(88, 21);
			this.cmbCampo.TabIndex = 0;
			this.cmbCampo.ValueMember = "";
			this.cmbCampo.Visible = false;
			// 
			// txtTexto
			// 
			this.txtTexto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTexto.Location = new System.Drawing.Point(160, 32);
			this.txtTexto.Name = "txtTexto";
			this.txtTexto.Size = new System.Drawing.Size(0, 21);
			this.txtTexto.TabIndex = 1;
			// 
			// cdsCentroCosto
			// 
			this.cdsCentroCosto.BindingContextCtrl = this;
			this.cdsCentroCosto.DataLibrary = "LibContabilidad";
			this.cdsCentroCosto.DataLibraryUrl = "";
			this.cdsCentroCosto.DataSetDef = "dsProyecto";
			this.cdsCentroCosto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCentroCosto.Name = "cdsCentroCosto";
			this.cdsCentroCosto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCentroCosto.SchemaDef = null;
			this.cdsCentroCosto.BeforeFill += new C1.Data.FillEventHandler(this.cdsGrupo_BeforeFill);
			// 
			// spnNumero
			// 
			this.spnNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnNumero.FormatString = "#,##0.00";
			this.spnNumero.Location = new System.Drawing.Point(536, 40);
			this.spnNumero.Name = "spnNumero";
			this.spnNumero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnNumero.PromptChar = ' ';
			this.spnNumero.Size = new System.Drawing.Size(64, 21);
			this.spnNumero.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnNumero.TabIndex = 3;
			this.spnNumero.Visible = false;
			this.spnNumero.ValueChanged += new System.EventHandler(this.spnNumero_ValueChanged);
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbFecha.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(192, 8);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.Size = new System.Drawing.Size(104, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 4;
			this.cmbFecha.Value = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			// 
			// btBuscar
			// 
			this.btBuscar.Location = new System.Drawing.Point(24, 8);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(64, 24);
			this.btBuscar.TabIndex = 5;
			this.btBuscar.Text = "&Generar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			this.cmbHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(384, 8);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.Size = new System.Drawing.Size(104, 21);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 7;
			this.cmbHasta.Value = new System.DateTime(2009, 5, 29, 0, 0, 0, 0);
			// 
			// spnHasta
			// 
			this.spnHasta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.spnHasta.FormatString = "#,##0.00";
			this.spnHasta.Location = new System.Drawing.Point(536, 40);
			this.spnHasta.Name = "spnHasta";
			this.spnHasta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.spnHasta.PromptChar = ' ';
			this.spnHasta.Size = new System.Drawing.Size(64, 21);
			this.spnHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.spnHasta.TabIndex = 8;
			this.spnHasta.Visible = false;
			// 
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(24, 40);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(64, 24);
			this.btExcel.TabIndex = 10;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// cmbTablas
			// 
			this.cmbTablas.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTablas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTablas.DisplayMember = "";
			this.cmbTablas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTablas.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTablas.Location = new System.Drawing.Point(512, 40);
			this.cmbTablas.Name = "cmbTablas";
			this.cmbTablas.Size = new System.Drawing.Size(104, 21);
			this.cmbTablas.TabIndex = 9;
			this.cmbTablas.ValueMember = "";
			this.cmbTablas.Visible = false;
			this.cmbTablas.ValueChanged += new System.EventHandler(this.cmbTablas_ValueChanged);
			this.cmbTablas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTablas_InitializeLayout);
			// 
			// chkTodoTipo
			// 
			this.chkTodoTipo.Location = new System.Drawing.Point(192, 40);
			this.chkTodoTipo.Name = "chkTodoTipo";
			this.chkTodoTipo.Size = new System.Drawing.Size(120, 16);
			this.chkTodoTipo.TabIndex = 17;
			this.chkTodoTipo.Text = "Centro de Costo";
			this.chkTodoTipo.CheckedChanged += new System.EventHandler(this.chkTodoTipo_CheckedChanged);
			// 
			// cmbTipo
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipo.Appearance = appearance1;
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataSource = this.cdsCentroCosto;
			this.cmbTipo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "ArticuloTipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 12;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 8;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 71;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 18;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Width = 18;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 18;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 16;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 16;
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
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "Nombre";
			this.cmbTipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(328, 40);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(160, 21);
			this.cmbTipo.TabIndex = 16;
			this.cmbTipo.ValueMember = "idProyecto";
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
			this.ultraGrid1.Location = new System.Drawing.Point(16, 80);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(608, 184);
			this.ultraGrid1.TabIndex = 18;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.Location = new System.Drawing.Point(111, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(72, 23);
			this.ultraLabel1.TabIndex = 19;
			this.ultraLabel1.Text = "Fecha Desde";
			this.ultraLabel1.Click += new System.EventHandler(this.ultraLabel1_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.Location = new System.Drawing.Point(312, 8);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(72, 23);
			this.ultraLabel2.TabIndex = 20;
			this.ultraLabel2.Text = "Fecha Hasta";
			// 
			// ResumenCostosDI
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(640, 278);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.chkTodoTipo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.cmbTablas);
			this.Controls.Add(this.spnHasta);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.spnNumero);
			this.Controls.Add(this.txtTexto);
			this.Controls.Add(this.cmbCampo);
			this.Name = "ResumenCostosDI";
			this.Text = "Requisición Estado";
			this.Load += new System.EventHandler(this.Busqueda_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbCampo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTexto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.spnHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTablas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsGrupoArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsGrupoArt.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		DataSet miDataSet = new DataSet();
		DataView dv = new DataView();
		DataSet miDataSetG = new DataSet();
		DataView dvG = new DataView();
		DataSet miDataSetT = new DataSet();
		DataView dvT = new DataView();


		private void Busqueda_Load(object sender, System.EventArgs e)
		{
//			LlenaCampos(stTabla);
//			cmbCampo.Value = "Articulo";

//			Funcion miFuncion = new Funcion();
//			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
//			string stProcedimiento = "Select Distinct(Tabla) As Tabla From TablaCampo";
//			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, strConn);
//
//			myData.TableMappings.Add("Tabla", "ExcelSube");
//			
//			miDataSetT.Reset();
//			miDataSetT.Clear();
//			try
//			{
//				myData.Fill(miDataSetT);
//			}
//			catch(Exception ex)
//			{
//				MessageBox.Show(ex.Message, "Error al abrir Tablas de TablaCompra " );
//				return;
//			}
//
//			dvT = miDataSetT.Tables[0].DefaultView;
//			cmbTablas.DataSource = dvT;
//			cmbTablas.DataBind();
//			cmbTablas.Value = "Articulo";
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			DateTime dtFechaIni = (DateTime) this.cmbFecha.Value;
			DateTime dtFechaFin = (DateTime) this.cmbHasta.Value;
			int idTipoGrupo = 0;
			if (chkTodoTipo.Checked) idTipoGrupo = (int) cmbTipo.Value;
		
			Funcion miFuncion = new Funcion();
			string strConn = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
			string stProcedimiento = "exec Agro_Comp_Result_Conta" 
				+ " '" + dtFechaIni.ToString("yyyyMMdd")
				+ "', '" + dtFechaFin.ToString("yyyyMMdd") + "', "
				+ idTipoGrupo.ToString();
			
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
//			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompraReq"].Hidden = true;
//			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompraOrc"].Hidden = true;
//			ultraGrid1.DisplayLayout.Bands[0].Columns["idCompraCom"].Hidden = true;
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
//			LlenaCampos(stTabla);
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "ResumenCostoDI.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "ResumenCostoDI.xls");
		}

		private void spnNumero_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraLabel1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkTodoTipo_CheckedChanged(object sender, System.EventArgs e)
		{
			this.cmbTipo.Enabled = this.chkTodoTipo.Checked;

		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
