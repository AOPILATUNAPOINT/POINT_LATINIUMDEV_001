using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de LiquidaOP.
	/// </summary>
	public class LiquidaOP : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsOrdenPLiquida;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.Misc.UltraButton btFiltro;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtManoObra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastosFab;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMateriaPrima;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaIni;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFin;
		private Infragistics.Win.Misc.UltraButton btOrdenP;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdAsiento;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel6;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdOrdenPLiquida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.Misc.UltraLabel ultraLabel7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsHojaCostos;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.Misc.UltraButton btManoObra;
		private Infragistics.Win.Misc.UltraButton btGastosInd;
		private Infragistics.Win.Misc.UltraButton btMateriaPrima;
		private Infragistics.Win.Misc.UltraButton btCosto;
		private Infragistics.Win.Misc.UltraButton btDesperdicio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDesperdicio;
		private System.ComponentModel.IContainer components;

		public LiquidaOP()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrdenP", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrdenProd");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GastosInd");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManoObra");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desperdicio");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 11, true, "Band 0", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Total", 11, true);
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("OrdenP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CodArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrdenProd");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GastosInd");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ManoObra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desperdicio");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.udsHojaCostos = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsOrdenPLiquida = new C1.Data.C1DataSet();
			this.cmbFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btFiltro = new Infragistics.Win.Misc.UltraButton();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.txtManoObra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastosFab = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMateriaPrima = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btOrdenP = new Infragistics.Win.Misc.UltraButton();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraLabel6 = new Infragistics.Win.Misc.UltraLabel();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.btCosto = new Infragistics.Win.Misc.UltraButton();
			this.txtIdOrdenPLiquida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel7 = new Infragistics.Win.Misc.UltraLabel();
			this.btManoObra = new Infragistics.Win.Misc.UltraButton();
			this.btGastosInd = new Infragistics.Win.Misc.UltraButton();
			this.btMateriaPrima = new Infragistics.Win.Misc.UltraButton();
			this.btDesperdicio = new Infragistics.Win.Misc.UltraButton();
			this.txtDesperdicio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsHojaCostos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrdenPLiquida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtManoObra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFab)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMateriaPrima)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrdenPLiquida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesperdicio)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "Band 0";
			this.ultraGrid1.DataSource = this.udsHojaCostos;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "OrdenProduccion";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Width = 84;
			ultraGridColumn2.Formula = "dd/MMM/yyyy";
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 72;
			ultraGridColumn3.Header.VisiblePosition = 12;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 3;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "#,##0.00";
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 11;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Width = 82;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn10.Header.Caption = "idOrdenP";
			ultraGridColumn10.Header.VisiblePosition = 13;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Hidden = true;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 5;
			ultraGridColumn12.Width = 93;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.VisiblePosition = 8;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.Width = 91;
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
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance7;
			summarySettings1.DisplayFormat = "{0:#,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 128);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(544, 184);
			this.ultraGrid1.TabIndex = 0;
			// 
			// udsHojaCostos
			// 
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn4.DataType = typeof(System.Double);
			ultraDataColumn5.DataType = typeof(System.Double);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(System.Double);
			ultraDataColumn13.DataType = typeof(System.Double);
			ultraDataColumn14.DataType = typeof(System.Double);
			ultraDataColumn15.DataType = typeof(System.Double);
			this.udsHojaCostos.Band.Columns.AddRange(new object[] {
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
																															ultraDataColumn13,
																															ultraDataColumn14,
																															ultraDataColumn15});
			// 
			// cmbFechaIni
			// 
			this.cmbFechaIni.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.FechaIni"));
			dateButton1.Caption = "Today";
			this.cmbFechaIni.DateButtons.Add(dateButton1);
			this.cmbFechaIni.Format = "dd/MMM/yyyy";
			this.cmbFechaIni.Location = new System.Drawing.Point(192, 40);
			this.cmbFechaIni.Name = "cmbFechaIni";
			this.cmbFechaIni.NonAutoSizeHeight = 23;
			this.cmbFechaIni.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaIni.TabIndex = 1;
			this.cmbFechaIni.Value = new System.DateTime(2009, 3, 16, 0, 0, 0, 0);
			// 
			// cdsOrdenPLiquida
			// 
			this.cdsOrdenPLiquida.BindingContextCtrl = this;
			this.cdsOrdenPLiquida.DataLibrary = "LibFacturacion";
			this.cdsOrdenPLiquida.DataLibraryUrl = "";
			this.cdsOrdenPLiquida.DataSetDef = "dsOrdenPLiquida";
			this.cdsOrdenPLiquida.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsOrdenPLiquida.Name = "cdsOrdenPLiquida";
			this.cdsOrdenPLiquida.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsOrdenPLiquida.SchemaDef = null;
			this.cdsOrdenPLiquida.BeforeFill += new C1.Data.FillEventHandler(this.cdsOrdenPLiquida_BeforeFill);
			// 
			// cmbFechaFin
			// 
			this.cmbFechaFin.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.FechaFin"));
			dateButton2.Caption = "Today";
			this.cmbFechaFin.DateButtons.Add(dateButton2);
			this.cmbFechaFin.Format = "dd/MMM/yyyy";
			this.cmbFechaFin.Location = new System.Drawing.Point(360, 40);
			this.cmbFechaFin.Name = "cmbFechaFin";
			this.cmbFechaFin.NonAutoSizeHeight = 23;
			this.cmbFechaFin.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaFin.TabIndex = 2;
			this.cmbFechaFin.Value = new System.DateTime(2009, 3, 16, 0, 0, 0, 0);
			// 
			// btFiltro
			// 
			this.btFiltro.Location = new System.Drawing.Point(480, 40);
			this.btFiltro.Name = "btFiltro";
			this.btFiltro.Size = new System.Drawing.Size(40, 23);
			this.btFiltro.TabIndex = 5;
			this.btFiltro.Text = "Filtro";
			this.btFiltro.Click += new System.EventHandler(this.btFiltro_Click);
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = true;
			this.barraDatoSQL1.DataMember = "OrdenPLiquida";
			this.barraDatoSQL1.DataNombreId = "IdOrdenPLiquida";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsOrdenPLiquida;
			this.barraDatoSQL1.DataTabla = "OrdenPLiquida";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(32, 328);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(304, 24);
			this.barraDatoSQL1.TabIndex = 6;
			this.barraDatoSQL1.VisibleBorrar = true;
			this.barraDatoSQL1.VisibleBuscar = true;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(144, 40);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 8;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(312, 40);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 9;
			this.ultraLabel2.Text = "Hasta";
			// 
			// txtManoObra
			// 
			this.txtManoObra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.ManoObra"));
			this.txtManoObra.FormatString = "#,##0.00";
			this.txtManoObra.Location = new System.Drawing.Point(144, 88);
			this.txtManoObra.Name = "txtManoObra";
			this.txtManoObra.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtManoObra.PromptChar = ' ';
			this.txtManoObra.Size = new System.Drawing.Size(88, 21);
			this.txtManoObra.TabIndex = 10;
			// 
			// txtGastosFab
			// 
			this.txtGastosFab.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.GastosInd"));
			this.txtGastosFab.FormatString = "#,##0.00";
			this.txtGastosFab.Location = new System.Drawing.Point(240, 88);
			this.txtGastosFab.Name = "txtGastosFab";
			this.txtGastosFab.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastosFab.PromptChar = ' ';
			this.txtGastosFab.Size = new System.Drawing.Size(88, 21);
			this.txtGastosFab.TabIndex = 11;
			// 
			// txtMateriaPrima
			// 
			this.txtMateriaPrima.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.MateriaPrima"));
			this.txtMateriaPrima.FormatString = "#,##0.00";
			this.txtMateriaPrima.Location = new System.Drawing.Point(336, 88);
			this.txtMateriaPrima.Name = "txtMateriaPrima";
			this.txtMateriaPrima.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMateriaPrima.PromptChar = ' ';
			this.txtMateriaPrima.Size = new System.Drawing.Size(88, 21);
			this.txtMateriaPrima.TabIndex = 12;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btOrdenP
			// 
			this.btOrdenP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btOrdenP.Location = new System.Drawing.Point(344, 328);
			this.btOrdenP.Name = "btOrdenP";
			this.btOrdenP.Size = new System.Drawing.Size(32, 23);
			this.btOrdenP.TabIndex = 16;
			this.btOrdenP.Text = "OP";
			this.toolTip1.SetToolTip(this.btOrdenP, "Ver orden de Producción Seleccionada");
			this.btOrdenP.Click += new System.EventHandler(this.btOrdenP_Click);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.idAsiento"));
			this.txtIdAsiento.Location = new System.Drawing.Point(504, 320);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.Size = new System.Drawing.Size(56, 21);
			this.txtIdAsiento.TabIndex = 17;
			// 
			// btAsiento
			// 
			this.btAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAsiento.Location = new System.Drawing.Point(472, 328);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.Size = new System.Drawing.Size(24, 23);
			this.btAsiento.TabIndex = 18;
			this.btAsiento.Text = "A";
			this.toolTip1.SetToolTip(this.btAsiento, "Crea / Muestra Asiento");
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			this.btAsiento.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btAsiento_MouseDown);
			// 
			// optTipo
			// 
			this.optTipo.CheckedIndex = 0;
			this.optTipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.idTipo"));
			this.optTipo.ItemAppearance = appearance12;
			this.optTipo.ItemOrigin = new System.Drawing.Point(4, 4);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Materia Prima";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Mano de Obra";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.ItemSpacingVertical = 4;
			this.optTipo.Location = new System.Drawing.Point(16, 40);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(112, 48);
			this.optTipo.TabIndex = 19;
			this.optTipo.Text = "Materia Prima";
			// 
			// ultraLabel6
			// 
			this.ultraLabel6.AutoSize = true;
			this.ultraLabel6.Location = new System.Drawing.Point(32, 16);
			this.ultraLabel6.Name = "ultraLabel6";
			this.ultraLabel6.Size = new System.Drawing.Size(64, 14);
			this.ultraLabel6.TabIndex = 20;
			this.ultraLabel6.Text = "Distribución";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.Location = new System.Drawing.Point(384, 328);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 23);
			this.ultraButton1.TabIndex = 24;
			this.ultraButton1.Text = "F";
			this.toolTip1.SetToolTip(this.ultraButton1, "Ver Factura Seleccionada");
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btCosto
			// 
			this.btCosto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCosto.Location = new System.Drawing.Point(416, 328);
			this.btCosto.Name = "btCosto";
			this.btCosto.Size = new System.Drawing.Size(24, 23);
			this.btCosto.TabIndex = 30;
			this.btCosto.Text = "C";
			this.toolTip1.SetToolTip(this.btCosto, "Actualiza costo de producción");
			this.btCosto.Click += new System.EventHandler(this.btCosto_Click);
			// 
			// txtIdOrdenPLiquida
			// 
			this.txtIdOrdenPLiquida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdOrdenPLiquida.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.idOrdenPLiquida"));
			this.txtIdOrdenPLiquida.Location = new System.Drawing.Point(504, 344);
			this.txtIdOrdenPLiquida.Name = "txtIdOrdenPLiquida";
			this.txtIdOrdenPLiquida.Size = new System.Drawing.Size(56, 21);
			this.txtIdOrdenPLiquida.TabIndex = 21;
			// 
			// txtCodigo
			// 
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.Codigo"));
			this.txtCodigo.Location = new System.Drawing.Point(192, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(104, 21);
			this.txtCodigo.TabIndex = 22;
			// 
			// ultraLabel7
			// 
			this.ultraLabel7.AutoSize = true;
			this.ultraLabel7.Location = new System.Drawing.Point(144, 8);
			this.ultraLabel7.Name = "ultraLabel7";
			this.ultraLabel7.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel7.TabIndex = 23;
			this.ultraLabel7.Text = "Código";
			// 
			// btManoObra
			// 
			this.btManoObra.Location = new System.Drawing.Point(144, 64);
			this.btManoObra.Name = "btManoObra";
			this.btManoObra.Size = new System.Drawing.Size(88, 23);
			this.btManoObra.TabIndex = 25;
			this.btManoObra.Text = "Mano de Obra";
			this.btManoObra.Click += new System.EventHandler(this.btManoObra_Click);
			// 
			// btGastosInd
			// 
			this.btGastosInd.Location = new System.Drawing.Point(240, 64);
			this.btGastosInd.Name = "btGastosInd";
			this.btGastosInd.Size = new System.Drawing.Size(88, 23);
			this.btGastosInd.TabIndex = 28;
			this.btGastosInd.Text = "Gastos Indirec.";
			this.btGastosInd.Click += new System.EventHandler(this.btGastosInd_Click);
			// 
			// btMateriaPrima
			// 
			this.btMateriaPrima.Location = new System.Drawing.Point(336, 64);
			this.btMateriaPrima.Name = "btMateriaPrima";
			this.btMateriaPrima.Size = new System.Drawing.Size(88, 23);
			this.btMateriaPrima.TabIndex = 29;
			this.btMateriaPrima.Text = "Materia Prima";
			this.btMateriaPrima.Click += new System.EventHandler(this.btMateriaPrima_Click);
			// 
			// btDesperdicio
			// 
			this.btDesperdicio.Location = new System.Drawing.Point(440, 64);
			this.btDesperdicio.Name = "btDesperdicio";
			this.btDesperdicio.Size = new System.Drawing.Size(88, 23);
			this.btDesperdicio.TabIndex = 32;
			this.btDesperdicio.Text = "Desperdicio";
			this.btDesperdicio.Click += new System.EventHandler(this.btDesperdicio_Click);
			// 
			// txtDesperdicio
			// 
			this.txtDesperdicio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsOrdenPLiquida, "OrdenPLiquida.Desperdicio"));
			this.txtDesperdicio.FormatString = "#,##0.00";
			this.txtDesperdicio.Location = new System.Drawing.Point(440, 88);
			this.txtDesperdicio.Name = "txtDesperdicio";
			this.txtDesperdicio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDesperdicio.PromptChar = ' ';
			this.txtDesperdicio.Size = new System.Drawing.Size(88, 21);
			this.txtDesperdicio.TabIndex = 31;
			// 
			// LiquidaOP
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(576, 373);
			this.Controls.Add(this.btDesperdicio);
			this.Controls.Add(this.txtDesperdicio);
			this.Controls.Add(this.btCosto);
			this.Controls.Add(this.btMateriaPrima);
			this.Controls.Add(this.btGastosInd);
			this.Controls.Add(this.btManoObra);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.ultraLabel7);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtIdOrdenPLiquida);
			this.Controls.Add(this.ultraLabel6);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btOrdenP);
			this.Controls.Add(this.txtMateriaPrima);
			this.Controls.Add(this.txtGastosFab);
			this.Controls.Add(this.txtManoObra);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.btFiltro);
			this.Controls.Add(this.cmbFechaFin);
			this.Controls.Add(this.cmbFechaIni);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "LiquidaOP";
			this.Text = "Liquidacion de Ordenes de Produccion";
			this.Load += new System.EventHandler(this.LiquidaOP_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsHojaCostos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOrdenPLiquida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtManoObra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastosFab)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMateriaPrima)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrdenPLiquida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDesperdicio)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
			if (txtIdAsiento.Text.Trim().Length > 0)
			{
				MessageBox.Show("No puede Borrar Registro porque tiene un Asiento ya Generado",
					"Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea borrar la Liquidacion de Orden de Producción?"
				+ ".\nEste proceso no se puede reversar", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
				MessageBoxDefaultButton.Button2)) return;

			barraDatoSQL1.BorraRegistro();
		}
	
		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		}

		private void ValoresIniciales()
		{
			this.barraDatoSQL1_Refresca(this, new System.EventArgs());
			drLiquida["Codigo"] = "";
			drLiquida["FechaIni"] = DateTime.Today;
			drLiquida["FechaFin"] = DateTime.Today;
			cmbFechaIni.Select();
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			barraDatoSQL1.CrearRegistro();
			ValoresIniciales();
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			BorraErrores();
			barraDatoSQL1.DeshacerRegistro();
		}

		private void BorraErrores()
		{
			errorProvider.SetError(txtManoObra, "");
			errorProvider.SetError(txtMateriaPrima, "");
			errorProvider.SetError(txtGastosFab, "");
			errorProvider.SetError(txtCodigo, "");
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			
			barraDatoSQL1.EditarRegistro();
		}

		private bool VerificaNumero(string stNumero)
		{
			bool bOk = true;
			BorraErrores();
			string stSelect = "Select Count(*) From OrdenPLiquida Where Codigo = '"
				+ stNumero.Trim() + "' And idOrdenPLiquida <> " + txtIdOrdenPLiquida.Text;
			int iCuenta = Funcion.iEscalarSQL(cdsOrdenPLiquida, stSelect, true);
			if (iCuenta > 0)
			{
				errorProvider.SetError(txtCodigo, "Codigo ya Existe");
				bOk = false;
			}
			return bOk;
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			if (barraDatoSQL1.bNuevo) // Numeración
			{
				string stNumero = Funcion.NumeraFactura(cdsOrdenPLiquida, 32);
				if (stNumero.Trim().Length > 0) drLiquida["Codigo"] = stNumero;
				else
				{
					if (txtCodigo.Value == null || txtCodigo.Value.ToString().Trim().Length == 0)
					{
						this.errorProvider.SetError(txtCodigo, "Ingrese Codigo");
						return;
					}
				}
			}
			if (!VerificaNumero(drLiquida["Codigo"].ToString())) return;
			
			barraDatoSQL1.GrabaRegistro();		
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaIni = (DateTime ) cmbFechaIni.Value;
			DateTime dtFechaFin = (DateTime ) cmbFechaFin.Value;
			string stFiltro = "";
			Reporte miRep = new Reporte("OrdenPHojaCosto.rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaIni = (DateTime ) cmbFechaIni.Value;
			DateTime dtFechaFin = (DateTime ) cmbFechaFin.Value;
			string stFiltro = "{OrdenProduccion.Fecha} >= #"
				+ dtFechaIni.ToString("MM-dd-yyyy") + "# And {OrdenProduccion.Fecha} < #"
				+ dtFechaFin.AddDays(1).ToString("MM-dd-yyyy") + "#";
			Reporte miRep = new Reporte("OrdenPLiquida.rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		C1.Data.C1DataRow drLiquida;
		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				drLiquida = cdsOrdenPLiquida.TableViews["OrdenPLiquida"].Rows[0];
			}
			catch{};		
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsOrdenPLiquida_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsOrdenPLiquida.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void LiquidaOP_Load(object sender, System.EventArgs e)
		{
			txtIdAsiento.Width = 0;
			txtIdOrdenPLiquida.Width = 0;
			barraDatoSQL1.ProximoId(4);
		}

		private void btFiltro_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec OrdenPHoja " + txtIdOrdenPLiquida.Text;



      SqlDataReader dr = Funcion.rEscalarSQL(cdsOrdenPLiquida, stExec, true);
			udsHojaCostos.Rows.Clear();
			int iFila = 0;
			while(dr.Read())
			{
				try
				{
					udsHojaCostos.Rows.Add();
					udsHojaCostos.Rows[iFila]["Codigo"]=dr.GetString(0);
					udsHojaCostos.Rows[iFila]["Fecha"]=dr.GetDateTime(1);
					udsHojaCostos.Rows[iFila]["Cantidad"]=dr.GetDouble(2);
					udsHojaCostos.Rows[iFila]["Costo"]=dr.GetDouble(3);
					udsHojaCostos.Rows[iFila]["Tipo"]=dr.GetString(4);
					udsHojaCostos.Rows[iFila]["idOrdenProd"]=dr.GetInt32(5);
					udsHojaCostos.Rows[iFila]["idCompra"]=dr.GetInt32(6);
					udsHojaCostos.Rows[iFila]["Articulo"]=dr.GetString(7);
					udsHojaCostos.Rows[iFila]["CodArticulo"]=dr.GetString(8);
					udsHojaCostos.Rows[iFila]["OrdenP"]=dr.GetString(9);
					udsHojaCostos.Rows[iFila]["Total"]=dr.GetDouble(2)*dr.GetDouble(3);
					udsHojaCostos.Rows[iFila]["GastosInd"]= dr.GetDouble(10);
					udsHojaCostos.Rows[iFila]["ManoObra"]= dr.GetDouble(11);
					udsHojaCostos.Rows[iFila]["Desperdicio"]= dr.GetDouble(12);
					iFila ++;
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Llenar grilla");
				}
			}
			dr.Close();
			ultraGrid1.Refresh();
		}

		private void btOrdenP_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null || ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value == null)
			{
				MessageBox.Show("Seleccione una Fila", "Información");
				return;
			}
			int IdOrdenP = (int) ultraGrid1.ActiveRow.Cells["idOrdenProd"].Value;
//			OrdenProduccion miOrden = new OrdenProduccion(IdOrdenP);
//			miOrden.MdiParent = this.MdiParent;
//			miOrden.Show();
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			if (txtIdOrdenPLiquida.Text.Length == 0)
			{
				MessageBox.Show("Grabe la Liquidacion para generar el asiento",
					"Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			string stExec = "Exec CreaAsientoOP " + txtIdOrdenPLiquida.Text;
			string stResultado = Funcion.sEscalarSQL(cdsOrdenPLiquida, stExec, true);
			if (stResultado != null && stResultado.Length > 0)
			{
				if (stResultado.StartsWith("No"))
				{
					MessageBox.Show(stResultado, "Información", MessageBoxButtons.OK,
						MessageBoxIcon.Information);
					Cursor = Cursors.Default;
					return;
				}
				if (DialogResult.No == MessageBox.Show(
					stResultado + " ¿Desea continuar?",
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) 
				{
					Cursor = Cursors.Default;
					return;
				}
				stExec += ", 1";
				stResultado = Funcion.sEscalarSQL(cdsOrdenPLiquida, stExec, true);
			}
			MessageBox.Show("Asiento Generado", "Información", MessageBoxButtons.OK,
				MessageBoxIcon.Information);
			barraDatoSQL1.ProximoId(5);
			Cursor = Cursors.Default;
		}

		private void btAsiento_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (txtIdAsiento.Text.Trim().Length == 0)
				{
					MessageBox.Show("No existe un Asiento Generado",
						"Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				int idAsiento = int.Parse(txtIdAsiento.Text); 
				Asiento miAsiento = new Asiento(idAsiento);
				miAsiento.MdiParent = this.MdiParent;
				miAsiento.Show();
			}
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.ActiveRow == null || ultraGrid1.ActiveRow.Cells["idCompra"].Value == null)
			{
				MessageBox.Show("Seleccione una Fila", "Información");
				return;
			}
			int IdCompra = (int) ultraGrid1.ActiveRow.Cells["idCompra"].Value;
			string stSelect = "Select idTipoFactura From Compra Where idCompra = " + IdCompra.ToString();
			int idTipo = Funcion.iEscalarSQL(cdsOrdenPLiquida, stSelect, true);
//			Compra miOrden = new Compra(idTipo, IdCompra);
//			miOrden.MdiParent = this.MdiParent;
//			miOrden.Show();
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btManoObra_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(txtManoObra, "");
			if (txtManoObra.Value == null || (double) txtManoObra.Value == 0)
			{
				errorProvider.SetError(txtManoObra, "Ingreso valor de Mano de Obra");
				return;
			}
			double dManoObra = (double) txtManoObra.Value;
			string stSelect = "EXEC OrdenPLiquidaMO " + dManoObra.ToString("0.00", us);
			Funcion.EjecutaSQL(cdsOrdenPLiquida, stSelect, true);
			MessageBox.Show("Mano de Obra repartida, para ver el resultado presione Filtro",
				"Información");
		}

		private void btGastosInd_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(txtGastosFab, "");
			if (txtGastosFab.Value == null || (double) txtGastosFab.Value == 0)
			{
				errorProvider.SetError(txtGastosFab, "Ingrese Gastos de Fabricacion");
				return;
			}
			double dMateriaP = (double) txtGastosFab.Value;
			string stSelect = "EXEC OrdenPLiquidaGI " + dMateriaP.ToString("0.00", us);
			Funcion.EjecutaSQL(cdsOrdenPLiquida, stSelect, true);
			MessageBox.Show("Gastos Indirectas repartidos, para ver el resultado presione Filtro",
				"Información");
		}

		private void btMateriaPrima_Click(object sender, System.EventArgs e)
		{
			double dMateriaP = (double) txtGastosFab.Value;
			string stSelect = "EXEC OrdenPLiquidaMP ";
			Funcion.EjecutaSQL(cdsOrdenPLiquida, stSelect, true);
			MessageBox.Show("Materia Prima recalculada",
				"Información");
		}

		private void btCosto_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec OrdenPCosto " + txtIdOrdenPLiquida.Text;
			Funcion.EjecutaSQL(cdsOrdenPLiquida, stExec, true);
			MessageBox.Show("Costo en producto terminado calculado", "Información");
		}

		private void btDesperdicio_Click(object sender, System.EventArgs e)
		{
			errorProvider.SetError(txtDesperdicio, "");
			if (txtDesperdicio.Value == null || (double) txtDesperdicio.Value == 0)
			{
				errorProvider.SetError(txtDesperdicio, "Ingrese Desperdicio");
				return;
			}
			double dMateriaP = (double) txtDesperdicio.Value;
			string stSelect = "EXEC OrdenPLiquidaDesp " + dMateriaP.ToString("0.00", us);
			Funcion.EjecutaSQL(cdsOrdenPLiquida, stSelect, true);
			MessageBox.Show("Desperdicio repartido, para ver el resultado presione Filtro",
				"Información");		
		}
	}
}
