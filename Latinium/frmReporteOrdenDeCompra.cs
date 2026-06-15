using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using Latinium.Services.OrdenDeCompra.Bodega;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteOrdenDeCompra.
	/// </summary>
	public class frmReporteOrdenDeCompra : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnBuscar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGenerar;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdOrdenCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGridBodega;
		private OrdenCompraBodegaGeneracionService serviceBodega;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSourceBodega;

		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteOrdenDeCompra()
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
		private Acceso miAcceso;

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdOrdenCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadAsignada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadGenerada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoPendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadGenerar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccionado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetalle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdOrdenCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreBodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadAsignada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadGenerada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantidadGenerar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Seleccionado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdDetOrdenDeCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdOrdenDeCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("GYE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Total");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantFactUIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantFactGYE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoUIO");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoGYE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("IdArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteOrdenDeCompra));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdDetOrdenDeCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdOrdenDeCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UIO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GYE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantFactUIO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantFactGYE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoUIO");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoGYE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel", 0);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnBodega", 1);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			this.ultraGridBodega = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSourceBodega = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnGenerar = new System.Windows.Forms.Button();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdOrdenCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.ultraGridBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSourceBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrdenCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGridBodega
			// 
			this.ultraGridBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGridBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGridBodega.DataSource = this.ultraDataSourceBodega;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGridBodega.DisplayLayout.Appearance = appearance1;
			this.ultraGridBodega.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 76;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 125;
			ultraGridColumn3.Header.Caption = "Bodega";
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Width = 183;
			ultraGridColumn4.Header.VisiblePosition = 4;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 168;
			ultraGridColumn5.Header.Caption = "Cantidad Asignada";
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.Header.Caption = "Cantidad Generada";
			ultraGridColumn6.Header.VisiblePosition = 6;
			ultraGridColumn6.Width = 183;
			ultraGridColumn7.Header.Caption = "Saldo Pendiente";
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 156;
			ultraGridColumn8.Header.Caption = "Cantidad Generar";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 175;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 115;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 113;
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
																										 ultraGridColumn10});
			this.ultraGridBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGridBodega.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.ultraGridBodega.DisplayLayout.GroupByBox.Hidden = true;
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGridBodega.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ultraGridBodega.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGridBodega.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGridBodega.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGridBodega.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted;
			this.ultraGridBodega.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ultraGridBodega.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGridBodega.DisplayLayout.Override.CellAppearance = appearance4;
			this.ultraGridBodega.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGridBodega.DisplayLayout.Override.CellPadding = 3;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGridBodega.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridBodega.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridBodega.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			this.ultraGridBodega.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGridBodega.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ultraGridBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGridBodega.Location = new System.Drawing.Point(0, 288);
			this.ultraGridBodega.Name = "ultraGridBodega";
			this.ultraGridBodega.Size = new System.Drawing.Size(992, 184);
			this.ultraGridBodega.TabIndex = 0;
			// 
			// ultraDataSourceBodega
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn6.DefaultValue = 0;
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn7.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(bool);
			this.ultraDataSourceBodega.Band.Columns.AddRange(new object[] {
																																			ultraDataColumn1,
																																			ultraDataColumn2,
																																			ultraDataColumn3,
																																			ultraDataColumn4,
																																			ultraDataColumn5,
																																			ultraDataColumn6,
																																			ultraDataColumn7,
																																			ultraDataColumn8,
																																			ultraDataColumn9,
																																			ultraDataColumn10});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// btnBuscar
			// 
			this.btnBuscar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(8, 480);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(104, 23);
			this.btnBuscar.TabIndex = 277;
			this.btnBuscar.Text = "&Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
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
			// btnGenerar
			// 
			this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGenerar.CausesValidation = false;
			this.btnGenerar.Enabled = false;
			this.btnGenerar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGenerar.Image = ((System.Drawing.Image)(resources.GetObject("btnGenerar.Image")));
			this.btnGenerar.ImageAlign = System.Drawing.ContentAlignment.BottomRight;
			this.btnGenerar.Location = new System.Drawing.Point(864, 16);
			this.btnGenerar.Name = "btnGenerar";
			this.btnGenerar.Size = new System.Drawing.Size(112, 40);
			this.btnGenerar.TabIndex = 379;
			this.btnGenerar.Text = "Generar Ordenes De Compra ";
			this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(24, 48);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(29, 16);
			this.lblIdentificacion.TabIndex = 387;
			this.lblIdentificacion.Text = "RUC";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(224, 48);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 386;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(224, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 382;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(24, 8);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 381;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance9.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance9;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "_DetCompra";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 56;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 50;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 105;
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.Width = 354;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 119;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.Width = 27;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 27;
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.Width = 123;
			ultraGridColumn19.Header.VisiblePosition = 9;
			ultraGridColumn19.Width = 95;
			ultraGridColumn20.Header.VisiblePosition = 10;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 41;
			ultraGridColumn21.Header.VisiblePosition = 11;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 42;
			ultraGridColumn22.Header.VisiblePosition = 12;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 31;
			ultraGridColumn23.Header.VisiblePosition = 13;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 33;
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 107;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 67;
			ultraGridColumn26.DataType = typeof(bool);
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 58;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellAppearance = appearance10;
			appearance11.Image = ((object)(resources.GetObject("appearance11.Image")));
			appearance11.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn27.CellButtonAppearance = appearance11;
			ultraGridColumn27.Header.Caption = "Asignación";
			ultraGridColumn27.Header.VisiblePosition = 16;
			ultraGridColumn27.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn27.Width = 117;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			ultraGridBand2.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance12.ForeColor = System.Drawing.Color.Black;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance12;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance13.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance13;
			appearance14.ForeColor = System.Drawing.Color.Black;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance15.FontData.BoldAsString = "True";
			appearance15.FontData.Name = "Arial";
			appearance15.FontData.SizeInPoints = 8.5F;
			appearance15.ForeColor = System.Drawing.Color.White;
			appearance15.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance15;
			appearance16.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance16.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(992, 168);
			this.ultraGrid1.TabIndex = 388;
			this.ultraGrid1.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_CellChange);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			// 
			// txtRuc
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance19;
			this.txtRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(72, 48);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 21);
			this.txtRuc.TabIndex = 385;
			// 
			// txtNombre
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance20;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(280, 48);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(216, 21);
			this.txtNombre.TabIndex = 384;
			// 
			// dtFecha
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance21;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 383;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// txtNumero
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance22;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(280, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 380;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ForeColor = System.Drawing.Color.Black;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(120, 480);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(96, 23);
			this.btnSalir.TabIndex = 390;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(904, 472);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 21);
			this.txtIdComprobante.TabIndex = 416;
			this.txtIdComprobante.Visible = false;
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(920, 472);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdGrupoCliente.TabIndex = 414;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(936, 472);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 21);
			this.txtIdTipoRuc.TabIndex = 413;
			this.txtIdTipoRuc.Visible = false;
			// 
			// txtIdOrdenCompra
			// 
			this.txtIdOrdenCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdOrdenCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdOrdenCompra.Enabled = false;
			this.txtIdOrdenCompra.Location = new System.Drawing.Point(952, 472);
			this.txtIdOrdenCompra.Name = "txtIdOrdenCompra";
			this.txtIdOrdenCompra.PromptChar = ' ';
			this.txtIdOrdenCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdOrdenCompra.TabIndex = 412;
			this.txtIdOrdenCompra.Visible = false;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(968, 472);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdCliente.TabIndex = 411;
			this.txtIdCliente.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(888, 472);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 417;
			this.txtEstado.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 264);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(240, 16);
			this.label1.TabIndex = 418;
			this.label1.Text = "Generación de ordenes de compra por Bodega";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmReporteOrdenDeCompra
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(992, 510);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.ultraGridBodega);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.txtIdOrdenCompra);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.btnGenerar);
			this.Controls.Add(this.btnBuscar);
			this.Name = "frmReporteOrdenDeCompra";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = " Reporte Orden De Compra";
			this.Load += new System.EventHandler(this.frmReporteOrdenDeCompra_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGridBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSourceBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdOrdenCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmReporteOrdenDeCompra_Load(object sender, System.EventArgs e)
		{
			serviceBodega =
				new OrdenCompraBodegaGeneracionService(
				cdsSeteoF
				);

			miAcceso = new Acceso(cdsSeteoF, "0637");
			if (!Funcion.Permiso("2180", cdsSeteoF))
			{				
				MessageBox.Show("No Puede Ingresar a Reporte Ordenes de Compra", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(this.Close));				

				return;
			}
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			using (frmBuscaOrdenDeCompra Busqueda = new frmBuscaOrdenDeCompra())
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.ConsultaMaestro((int)Busqueda.uGridOrdenesDeCompra.ActiveRow.Cells["IdOrdenDeCompra"].Value);
		}	
	
		public void ConsultaMaestro(int IdOrdenCompra)
		{
			try
			{
				// =========================
				// 1. CABECERA
				// =========================
				string sSQL = string.Format(
					"Exec ConsultaOrdenDeCompraMaestro {0}",
					IdOrdenCompra
					);

				SqlDataReader dr =
					Funcion.rEscalarSQL(
					cdsSeteoF,
					sSQL,
					true
					);

				if (dr.Read())
				{
					this.txtIdOrdenCompra.Value = IdOrdenCompra;
					this.txtIdCliente.Value = dr.GetInt32(1);
					this.txtIdGrupoCliente.Value = dr.GetInt32(2);
					this.txtRuc.Text = dr.GetString(3);
					this.txtNombre.Text = dr.GetString(4);
					this.dtFecha.Value = dr.GetDateTime(5);
					this.txtNumero.Text = dr.GetString(6);
					this.txtEstado.Value = dr.GetInt32(11);
					this.txtIdComprobante.Value = dr.GetInt32(14);
				}

				dr.Close();


				// =========================
				// 2. GRID MAESTRO
				// =========================
				DataTable dtMaestro =
					FuncionesProcedimientos.dtGeneral(
					string.Format(
					"Exec ConsultaMaestroOrdenCompraNuevo '',1,{0}",
					IdOrdenCompra
					)
					);

				this.ultraGrid1.DataSource = dtMaestro;
				this.ultraGrid1.DataBind();
				this.ultraGrid1.Refresh();


				// =========================
				// 3. GRID BODEGA (LIMPIEZA SEGURA)
				// =========================

				DataTable dtBodega =
					this.ultraGridBodega.DataSource as DataTable;

				if (dtBodega == null)
				{
					// SOLO UNA VEZ: inicializa estructura
					dtBodega = new DataTable();
					this.ultraGridBodega.DataSource = dtBodega;
					this.ultraGridBodega.DataBind();
				}
				else
				{
					// SOLO LIMPIA DATOS (NO COLUMNAS)
					dtBodega.Clear();
				}

				// refresco seguro sin romper layout
				this.ultraGridBodega.UpdateData();
				this.ultraGridBodega.Refresh();


				// =========================
				// 4. CONFIG GRID BODEGA
				// =========================
				ConfigurarGridBodega();


				// =========================
				// 5. ESTADO
				// =========================
				int est =
					(this.txtEstado.Value == null)
					? 0
					: Convert.ToInt32(this.txtEstado.Value);

				this.btnGenerar.Enabled = (est == 5);

				FuncionesProcedimientos.EstadoGrids(
					(est == 5),
					this.ultraGrid1
					);

				FuncionesProcedimientos.EstadoGrids(
					(est == 5),
					this.ultraGridBodega
					);
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error al Consultar: " + ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}
		private void ConfigurarGridBodega()
		{
			if (this.ultraGridBodega.DisplayLayout.Bands.Count == 0)
				return;

			UltraGridBand band =
				this.ultraGridBodega.DisplayLayout.Bands[0];

			// =========================
			// PERMITIR EDICIÓN (SI APLICA)
			// =========================
			this.ultraGridBodega.DisplayLayout.Override.AllowUpdate =
				Infragistics.Win.DefaultableBoolean.True;

			// =========================
			// EVITAR ELIMINACIÓN (CLAVE)
			// =========================
			this.ultraGridBodega.DisplayLayout.Override.AllowDelete =
				Infragistics.Win.DefaultableBoolean.False;

			// Evita acción de borrar por teclado o UI
			this.ultraGridBodega.DisplayLayout.Override.AllowRowFiltering =
				Infragistics.Win.DefaultableBoolean.False;

			this.ultraGridBodega.DisplayLayout.Override.RowSelectors =
				Infragistics.Win.DefaultableBoolean.False;

			// Bloquea comportamiento de edición tipo delete row
			this.ultraGridBodega.DisplayLayout.Override.CellClickAction =
				CellClickAction.Edit;

			// =========================
			// OPCIONAL (SEGURIDAD EXTRA)
			// =========================
			this.ultraGridBodega.DisplayLayout.Bands[0].Override.AllowDelete =
				Infragistics.Win.DefaultableBoolean.False;
		}
//		public void ConsultaMaestro(int IdOrdenCompra)
//		{
//			try
//			{
//				// =========================
//				// 1. CABECERA
//				// =========================
//				string sSQL = string.Format("Exec ConsultaOrdenDeCompraMaestro {0}", IdOrdenCompra);
//				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
//
//				if (dr.Read())
//				{
//					this.txtIdOrdenCompra.Value = IdOrdenCompra;
//					this.txtIdCliente.Value = dr.GetInt32(1);
//					this.txtIdGrupoCliente.Value = dr.GetInt32(2);
//					this.txtRuc.Text = dr.GetString(3);
//					this.txtNombre.Text = dr.GetString(4);
//					this.dtFecha.Value = dr.GetDateTime(5);
//					this.txtNumero.Text = dr.GetString(6);
//					this.txtEstado.Value = dr.GetInt32(11);
//					this.txtIdComprobante.Value = dr.GetInt32(14);
//				}
//				dr.Close();
//
//
//				// =========================
//				// 2. DETALLE GRID
//				// =========================
//				DataTable dt =
//					FuncionesProcedimientos.dtGeneral(
//					string.Format("Exec ConsultaMaestroOrdenCompraNuevo '',1,{0}", IdOrdenCompra)
//					);
//
//				// =========================
//				// 3. COLUMNAS NUEVAS (CLAVE)
//				// =========================
//				if (!dt.Columns.Contains("CantidadTemporal"))
//					dt.Columns.Add("CantidadTemporal", typeof(int));
//
//				if (!dt.Columns.Contains("Sel"))
//					dt.Columns.Add("Sel", typeof(int));
//
//
//				// =========================
//				// 4. NORMALIZAR VALORES
//				// =========================
//				foreach (DataRow r in dt.Rows)
//				{
//					if (r["CantidadTemporal"] == DBNull.Value)
//						r["CantidadTemporal"] = 0;
//
//					if (r["Sel"] == DBNull.Value)
//						r["Sel"] = 0;
//				}
//
//
//				// =========================
//				// 5. BIND GRID (IMPORTANTE)
//				// =========================
//				this.ultraGrid1.DataSource = dt;
//
//				// fuerza refresh real del binding Infragistics
//				this.ultraGrid1.DataBind();
//				this.ultraGrid1.Refresh();
//
//
//				// =========================
//				// 6. CONFIGURAR GRID
//				// =========================
//				ConfigurarGrid();
//
//
//				// =========================
//				// 7. ESTADO BOTÓN
//				// =========================
//				int est = (this.txtEstado.Value == null)
//					? 0
//					: Convert.ToInt32(this.txtEstado.Value);
//
//				if (est == 5)
//				{
//					this.btnGenerar.Enabled = true;
//					FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
//				}
//				else
//				{
//					this.btnGenerar.Enabled = false;
//					FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
//				}
//			}
//			catch (Exception ex)
//			{
//				MessageBox.Show(
//					string.Format("Error al Consultar: {0}", ex.Message),
//					"Point Technology",
//					MessageBoxButtons.OK,
//					MessageBoxIcon.Error
//					);
//			}
//		}

		private void uGridOrdenesDeCompra_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{	

		}

		private void uGridOrdenesDeCompra_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		
		}	

		private void uGridOrdenesDeCompra_InitializeGroupByRow(object sender, Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs e)
		{
		
		}			

//		private void btnGenerar_Click(object sender, System.EventArgs e)
//		{
//			#region Validaciones
//			bool haySeleccion = false;
//			bool hayUIOConCantidad = false;
//			bool hayGYEConCantidad = false;
//			bool haySeleccionCeroEnAmbos = false;
//
//			for (int i = 0; i < this.ultraGrid1.Rows.Count; i++)
//			{
//				Infragistics.Win.UltraWinGrid.UltraGridRow r = this.ultraGrid1.Rows[i];
//
//				bool sel = (r.Cells["Sel"].Value != null && r.Cells["Sel"].Value != System.DBNull.Value &&
//					System.Convert.ToBoolean(r.Cells["Sel"].Value));
//				if (!sel) continue;
//
//				haySeleccion = true;
//
//				int cantUIO = 0;
//				int cantGYE = 0;
//
//				if (r.Cells["CantFactUIO"].Value != null && r.Cells["CantFactUIO"].Value != System.DBNull.Value)
//				{
//						try { cantUIO = System.Convert.ToInt32(r.Cells["CantFactUIO"].Value); } 
//					catch { cantUIO = 0; } 
//				}
//
//				if (r.Cells["CantFactGYE"].Value != null && r.Cells["CantFactGYE"].Value != System.DBNull.Value)
//				{
//						try { cantGYE = System.Convert.ToInt32(r.Cells["CantFactGYE"].Value); } 
//					catch { cantGYE = 0; } 
//				}
//
//				if (cantUIO > 0) hayUIOConCantidad = true;
//				if (cantGYE > 0) hayGYEConCantidad = true;
//				if (cantUIO <= 0 && cantGYE <= 0) haySeleccionCeroEnAmbos = true;
//			}
//
//			if (!haySeleccion)
//			{
//				System.Windows.Forms.MessageBox.Show("Debe seleccionar al menos una fila (columna 'Sel').", "Point Technology",
//					System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
//				return;
//			}
//			if (!hayUIOConCantidad && !hayGYEConCantidad)
//			{
//				System.Windows.Forms.MessageBox.Show("Debe ingresar una cantidad mayor a 0 en 'CantFactUIO' y/o 'CantFactGYE' en las filas seleccionadas.", "Point Technology",
//					System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
//				return;
//			}
//			if (haySeleccionCeroEnAmbos)
//			{
//				System.Windows.Forms.MessageBox.Show("Hay filas seleccionadas con 'CantFactUIO' = 0 y 'CantFactGYE' = 0. Ajuste las cantidades o desmarque esas filas.", "Point Technology",
//					System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
//				return;
//			}
//			#endregion Validaciones
//
//			SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID=" + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024");
//			oConexion.Open();
//			SqlTransaction oTransaction = null;
//			SqlCommand oCmdActualiza = oConexion.CreateCommand();
//			oCmdActualiza.Connection = oConexion;
//			oCmdActualiza.CommandType = CommandType.Text;
//			oCmdActualiza.CommandTimeout = 120;
//
//			try
//			{
//				oTransaction = oConexion.BeginTransaction();
//				oCmdActualiza.Transaction = oTransaction;
//
//				oCmdActualiza.CommandText = string.Format(
//					"UPDATE d SET " +
//					"  SaldoUIO = CASE WHEN d.SaldoUIO = 0 AND (ISNULL(d.SaldoUIO,0) + ISNULL(d.SaldoGYE,0)) = 0 AND ISNULL(d.UIO,0) > 0 THEN ISNULL(d.UIO,0) ELSE d.SaldoUIO END, " +
//					"  SaldoGYE = CASE WHEN d.SaldoGYE = 0 AND (ISNULL(d.SaldoUIO,0) + ISNULL(d.SaldoGYE,0)) = 0 AND ISNULL(d.GYE,0) > 0 THEN ISNULL(d.GYE,0) ELSE d.SaldoGYE END " +
//					"FROM DetOrdenDeCompra d WHERE d.IdOrdenDeCompra = {0};",
//					Convert.ToInt32(this.txtIdOrdenCompra.Value)
//					);
//				oCmdActualiza.ExecuteNonQuery();
//
//				int[] bodegas = new int[] { 1, 101 };
//				System.Collections.ArrayList ocGeneradas = new System.Collections.ArrayList();
//
//				for (int i = 0; i < bodegas.Length; i++)
//				{
//					int bodega = bodegas[i];
//					bool hayDatos = false;
//
//					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in this.ultraGrid1.Rows)
//					{
//						bool sel = (r.Cells["Sel"].Value != DBNull.Value && r.Cells["Sel"].Value != null && Convert.ToBoolean(r.Cells["Sel"].Value));
//						if (!sel) continue;
//
//						int cantUIO = (r.Cells["CantFactUIO"].Value == null) ? 0 : Convert.ToInt32(r.Cells["CantFactUIO"].Value);
//						int cantGYE = (r.Cells["CantFactGYE"].Value == null) ? 0 : Convert.ToInt32(r.Cells["CantFactGYE"].Value);
//						int qtySolic = (bodega == 1) ? cantUIO : cantGYE;
//						if (qtySolic <= 0) continue;
//
//						int idDet = Convert.ToInt32(r.Cells["IdDetOrdenDeCompra"].Value);
//						string sSelSaldo = string.Format(
//							"SELECT ISNULL(SaldoUIO,0) AS SalUIO, ISNULL(SaldoGYE,0) AS SalGYE " +
//							"FROM DetOrdenDeCompra WHERE IdDetOrdenDeCompra = {0}",
//							idDet
//							);
//
//						oCmdActualiza.CommandText = sSelSaldo;
//						using (SqlDataReader rd = oCmdActualiza.ExecuteReader())
//						{
//							if (rd.Read())
//							{
//								int saldo = (bodega == 1) ? (int)Convert.ToDecimal(rd.GetValue(0)) : (int)Convert.ToDecimal(rd.GetValue(1));
//								if (Math.Min(qtySolic, saldo) > 0) { hayDatos = true; }
//							}
//						}
//						if (hayDatos) break;
//					}
//					if (!hayDatos) continue;
//
//					string notas = "";
//					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows)
//					{
//						bool sel = (dr.Cells["Sel"].Value != DBNull.Value && dr.Cells["Sel"].Value != null && Convert.ToBoolean(dr.Cells["Sel"].Value));
//						if (sel && dr.Cells["Notas"].Value != null && dr.Cells["Notas"].Value.ToString() != "")
//						{ notas = dr.Cells["Notas"].Value.ToString(); break; }
//					}
//					notas = notas.Replace("'", "''");
//
//					decimal subExentoCalc = 0m;
//					decimal subIvaCalc = 0m;
//					decimal ivaCalc = 0m;
//					decimal totalCalc = 0m;
//					bool hayLineaValida = false;
//
//					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow r in this.ultraGrid1.Rows)
//					{
//						bool sel = (r.Cells["Sel"].Value != null && r.Cells["Sel"].Value != DBNull.Value && Convert.ToBoolean(r.Cells["Sel"].Value));
//						if (!sel) continue;
//
//						int qtySolic = (bodega == 1)
//							? (r.Cells["CantFactUIO"].Value == null ? 0 : Convert.ToInt32(r.Cells["CantFactUIO"].Value))
//							: (r.Cells["CantFactGYE"].Value == null ? 0 : Convert.ToInt32(r.Cells["CantFactGYE"].Value));
//						if (qtySolic <= 0) continue;
//
//						int idDet = Convert.ToInt32(r.Cells["IdDetOrdenDeCompra"].Value);
//						string sSelSaldo = string.Format(
//							"SELECT ISNULL(SaldoUIO,0) AS SalUIO, ISNULL(SaldoGYE,0) AS SalGYE " +
//							"FROM DetOrdenDeCompra WHERE IdDetOrdenDeCompra = {0}",
//							idDet
//							);
//
//						oCmdActualiza.CommandText = sSelSaldo;
//						int saldoFila = 0;
//						using (SqlDataReader rSaldo = oCmdActualiza.ExecuteReader())
//						{
//							if (rSaldo.Read())
//							{
//								saldoFila = (bodega == 1) ? (int)Convert.ToDecimal(rSaldo.GetValue(0)) : (int)Convert.ToDecimal(rSaldo.GetValue(1));
//							}
//						}
//						if (saldoFila <= 0) continue;
//
//						int qty = qtySolic > saldoFila ? saldoFila : qtySolic;
//						if (qty <= 0) continue;
//
//						hayLineaValida = true;
//
//						decimal precio = (r.Cells["Costo"].Value == null ? 0m : Convert.ToDecimal(r.Cells["Costo"].Value));
//						decimal lineaSub = Math.Round(precio * qty, 2);
//						decimal porcIva = 15.00m;
//
//						if (porcIva > 0m) { subIvaCalc += lineaSub; ivaCalc += Math.Round(lineaSub * porcIva / 100m, 2); }
//						else { subExentoCalc += lineaSub; }
//					}
//					totalCalc = subExentoCalc + subIvaCalc + ivaCalc;
//					if (!hayLineaValida || totalCalc <= 0m) continue;
//
//					oCmdActualiza.CommandText = "Exec FacturaNumeroEstacion 2";
//					string numero = Convert.ToString(oCmdActualiza.ExecuteScalar());
//					DateTime fDoc = Convert.ToDateTime(this.dtFecha.Value);
//					String fTentativa = fDoc.AddDays(3).ToString("yyyyMMdd");
//
//					string sSQLMaestro = string.Format(
//						"Exec GrabaMaestroOrdenDeCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', {11}, {12}, '{13}', {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21},{22},{23},{24},{25},{26},{27}",
//						0, 2, (int)this.txtIdComprobante.Value, 0, 0, Convert.ToInt32(this.txtIdCliente.Value),
//						bodega, numero, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), fTentativa,
//						DateTime.Today.ToString("yyyyMMdd HH:mm"), 0, 5, notas, subExentoCalc, 0, subIvaCalc, 0, ivaCalc, totalCalc,
//						1, 0, 0, 0, 0, 0, 0, Convert.ToInt32(this.txtIdOrdenCompra.Value));
//					oCmdActualiza.CommandText = sSQLMaestro;
//					int idCompraNueva = Convert.ToInt32(oCmdActualiza.ExecuteScalar());
//
//					foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows)
//					{
//						bool sel = (dr.Cells["Sel"].Value != null && dr.Cells["Sel"].Value != DBNull.Value && Convert.ToBoolean(dr.Cells["Sel"].Value));
//						if (!sel) continue;
//
//						int qtySolic = (bodega == 1)
//							? (dr.Cells["CantFactUIO"].Value == null ? 0 : Convert.ToInt32(dr.Cells["CantFactUIO"].Value))
//							: (dr.Cells["CantFactGYE"].Value == null ? 0 : Convert.ToInt32(dr.Cells["CantFactGYE"].Value));
//						if (qtySolic <= 0) continue;
//
//						int idDet = Convert.ToInt32(dr.Cells["IdDetOrdenDeCompra"].Value);
//						string sSelSaldo = string.Format(
//							"SELECT ISNULL(SaldoUIO,0) AS SalUIO, ISNULL(SaldoGYE,0) AS SalGYE " +
//							"FROM DetOrdenDeCompra WHERE IdDetOrdenDeCompra = {0}",
//							idDet
//							);
//						oCmdActualiza.CommandText = sSelSaldo;
//						int saldoFila = 0;
//						using (SqlDataReader rd2 = oCmdActualiza.ExecuteReader())
//						{
//							if (rd2.Read())
//							{
//								saldoFila = (bodega == 1) ? (int)Convert.ToDecimal(rd2.GetValue(0)) : (int)Convert.ToDecimal(rd2.GetValue(1));
//							}
//						}
//						if (saldoFila <= 0) continue;
//
//						int qty = qtySolic;
//						if (saldoFila > 0 && qty > saldoFila) qty = saldoFila;
//						if (qty <= 0) continue;
//
//						decimal precio = (dr.Cells["Costo"].Value == null ? 0m : Convert.ToDecimal(dr.Cells["Costo"].Value));
//						decimal subDet = Math.Round(precio * qty, 2);
//						int idProyecto = (bodega == 1) ? 19 : 107;
//
//						string sSQLDetalle = string.Format(
//							"Exec GuardaDetalleOrdenDeCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}",
//							0, idCompraNueva, Convert.ToInt32(dr.Cells["IdArticulo"].Value), 0, qty, 0, Convert.ToDecimal(dr.Cells["Costo"].Value), 15.00, 0, 0, subDet,
//							idProyecto, 0, bodega, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
//							);
//						oCmdActualiza.CommandText = sSQLDetalle;
//						oCmdActualiza.ExecuteNonQuery();
//
//						if (bodega == 1)
//						{
//							oCmdActualiza.CommandText = string.Format(
//								"UPDATE DetOrdenDeCompra SET SaldoUIO = CASE WHEN SaldoUIO < {0} THEN 0 ELSE SaldoUIO - {0} END, CantaFactUIO = 0 " +
//								"WHERE IdDetOrdenDeCompra = {1};",
//								qty, idDet
//								);
//							oCmdActualiza.ExecuteNonQuery();
//						}
//						else
//						{
//							oCmdActualiza.CommandText = string.Format(
//								"UPDATE DetOrdenDeCompra SET SaldoGYE = CASE WHEN SaldoGYE < {0} THEN 0 ELSE SaldoGYE - {0} END, CantaFactGYE = 0 " +
//								"WHERE IdDetOrdenDeCompra = {1};",
//								qty, idDet
//								);
//							oCmdActualiza.ExecuteNonQuery();
//						}
//					}
//
//					ocGeneradas.Add(new object[] { idCompraNueva, bodega, numero });
//				}
//
//				oCmdActualiza.CommandText = string.Format(
//					"UPDATE d SET " +
//					"  SaldoUIO = CASE WHEN d.SaldoUIO < 0 THEN 0 WHEN d.SaldoUIO > ISNULL(d.UIO,0) THEN ISNULL(d.UIO,0) ELSE d.SaldoUIO END, " +
//					"  SaldoGYE = CASE WHEN d.SaldoGYE < 0 THEN 0 WHEN d.SaldoGYE > ISNULL(d.GYE,0) THEN ISNULL(d.GYE,0) ELSE d.SaldoGYE END " +
//					"FROM DetOrdenDeCompra d WHERE d.IdOrdenDeCompra = {0};",
//					Convert.ToInt32(this.txtIdOrdenCompra.Value)
//					);
//				oCmdActualiza.ExecuteNonQuery();
//
//				string sSQLCierra = string.Format(
//					"IF EXISTS (SELECT 1 FROM DetOrdenDeCompra WHERE IdOrdenDeCompra = {0} AND (ISNULL(SaldoUIO,0) > 0 OR ISNULL(SaldoGYE,0) > 0)) " +
//					"BEGIN UPDATE OrdenDeCompra SET Estado = 5 WHERE IdOrdenDeCompra = {0} AND ISNULL(Estado,0) <> 7 END " +
//					"ELSE BEGIN UPDATE OrdenDeCompra SET Estado = 9 WHERE IdOrdenDeCompra = {0} AND ISNULL(Estado,0) <> 7 END",
//					Convert.ToInt32(this.txtIdOrdenCompra.Value)
//					);
//				oCmdActualiza.CommandText = sSQLCierra;
//				oCmdActualiza.ExecuteNonQuery();
//
//				oCmdActualiza.CommandText = string.Format(
//					"UPDATE DetOrdenDeCompra SET CantaFactUIO = 0, CantaFactGYE = 0 WHERE IdOrdenDeCompra = {0};",
//					Convert.ToInt32(this.txtIdOrdenCompra.Value)
//					);
//				oCmdActualiza.ExecuteNonQuery();
//
//				oTransaction.Commit();
//
//				if (ocGeneradas.Count > 0)
//				{
//					for (int k = 0; k < ocGeneradas.Count; k++)
//					{
//						object[] oc = (object[])ocGeneradas[k];
//						int ocId = (int)oc[0];
//						int ocBod = (int)oc[1];
//						string ocNum = (string)oc[2];
//
//						frmTurnosMercaderia frm = new frmTurnosMercaderia(ocId, ocBod, ocNum, true);
//						frm.StartPosition = FormStartPosition.CenterParent;
//						frm.ShowDialog(this);
//					}
//					this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaMaestroOrdenCompraNuevo '',1,{0}", (int)this.txtIdOrdenCompra.Value));
//					MessageBox.Show("Órdenes de compra generadas correctamente.", "Point Technology");
//				}
//				else
//				{
//					MessageBox.Show("No se generó ninguna orden (no hubo líneas con saldo/qty válidos).", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
//				}
//			}
//			catch (Exception ex)
//			{
//				try { oTransaction.Rollback(); MessageBox.Show("Error: " + ex.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//				catch (Exception ex2) { MessageBox.Show("Rollback Error: " + ex2.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error); }
//			}
//			finally
//			{
//				oConexion.Close();
//			}
//		}

//		private void btnGenerar_Click(object sender, System.EventArgs e)
//		{
//			#region 1. VALIDAR SELECCIÓN
//
//			ArrayList filasSeleccionadas =
//				new ArrayList();
//
//			for (int i = 0; i < this.ultraGrid1.Rows.Count; i++)
//			{
//				UltraGridRow r =
//					this.ultraGrid1.Rows[i];
//
//				if (r == null || r.IsAddRow)
//					continue;
//
//				bool sel =
//					(r.Cells["Sel"].Value != null &&
//					r.Cells["Sel"].Value != DBNull.Value &&
//					Convert.ToBoolean(r.Cells["Sel"].Value));
//
//				if (!sel)
//					continue;
//
//				filasSeleccionadas.Add(r);
//			}
//
//			if (filasSeleccionadas.Count == 0)
//			{
//				MessageBox.Show(
//					"Debe seleccionar al menos un artículo.",
//					"Point Technology",
//					MessageBoxButtons.OK,
//					MessageBoxIcon.Warning
//					);
//
//				return;
//			}
//
//			#endregion
//
//
//			#region 2. SERVICE BODEGA
//
//			OrdenCompraBodegaGeneracionService service =
//				new OrdenCompraBodegaGeneracionService(
//				this.cdsSeteoF
//				);
//
//			#endregion
//
//
//			#region 3. CONEXIÓN SQL
//
//			SqlConnection oConexion =
//				new SqlConnection(
//				"Password=Bl45o6$9;Persist Security Info=True;User ID=" +
//				MenuLatinium.stUsuario +
//				";Initial Catalog=" +
//				MenuLatinium.stDirFacturacion +
//				";Data Source=" +
//				MenuLatinium.stDirServidor +
//				"; max pool size=1024"
//				);
//
//			oConexion.Open();
//
//			SqlTransaction oTransaction =
//				oConexion.BeginTransaction();
//
//			SqlCommand oCmd =
//				oConexion.CreateCommand();
//
//			oCmd.Connection = oConexion;
//			oCmd.Transaction = oTransaction;
//			oCmd.CommandType = CommandType.Text;
//			oCmd.CommandTimeout = 120;
//
//			#endregion
//
//
//			try
//			{
//				ArrayList ocGeneradas =
//					new ArrayList();
//
//
//				#region 4. RECORRER SOLO ARTÍCULOS SELECCIONADOS
//
//				foreach (UltraGridRow row in filasSeleccionadas)
//				{
//					int idDetOrdenDeCompra =
//						Convert.ToInt32(
//						row.Cells["IdDetOrdenDeCompra"].Value
//						);
//
//					int idArticulo =
//						Convert.ToInt32(
//						row.Cells["IdArticulo"].Value
//						);
//
//					decimal costo =
//						Convert.ToDecimal(
//						row.Cells["Costo"].Value
//						);
//
//					string notas = "";
//
//					if (row.Cells["Notas"].Value != null)
//					{
//						notas =
//							row.Cells["Notas"].Value.ToString();
//					}
//
//					notas =
//						notas.Replace("'", "''");
//
//
//					#region 5. CONSULTAR BODEGAS DEL ARTÍCULO
//
//					DataTable dtBodegas =
//						service.Consultar(
//						idDetOrdenDeCompra
//						);
//
//					if (dtBodegas == null ||
//						dtBodegas.Rows.Count == 0)
//					{
//						continue;
//					}
//
//					#endregion
//
//
//					#region 6. RECORRER BODEGAS
//
//					foreach (DataRow drBodega in dtBodegas.Rows)
//					{
//						int bodega =
//							Convert.ToInt32(
//							drBodega["Bodega"]
//							);
//
//						int cantidadTemporal = 0;
//
//						if (drBodega["CantidadTemporal"] != DBNull.Value)
//						{
//							cantidadTemporal =
//								Convert.ToInt32(
//								drBodega["CantidadTemporal"]
//								);
//						}
//
//						int saldoPendiente = 0;
//
//						if (drBodega["SaldoPendiente"] != DBNull.Value)
//						{
//							saldoPendiente =
//								Convert.ToInt32(
//								drBodega["SaldoPendiente"]
//								);
//						}
//
//						if (cantidadTemporal <= 0)
//							continue;
//
//						if (saldoPendiente <= 0)
//							continue;
//
//						if (cantidadTemporal > saldoPendiente)
//						{
//							cantidadTemporal =
//								saldoPendiente;
//						}
//
//						#endregion
//
//
//						#region 7. CALCULAR TOTALES
//
//						decimal subTotal =
//							(costo * cantidadTemporal);
//
//						decimal iva =
//							Math.Round(
//							subTotal * 0.15m,
//							2
//							);
//
//						decimal total =
//							subTotal + iva;
//
//						if (total <= 0)
//							continue;
//
//						#endregion
//
//
//						#region 8. GENERAR NÚMERO
//
//						oCmd.CommandText =
//							"Exec FacturaNumeroEstacion 2";
//
//						string numero =
//							Convert.ToString(
//							oCmd.ExecuteScalar()
//							);
//
//						string numeroOC = "";
//
//						if (this.txtNumero.Text != null)
//						{
//							numeroOC = this.txtNumero.Text.Trim();
//						}
//
//						if (numero == null)
//						{
//							numero = "";
//						}
//
//						numero = numeroOC + "-" + numero;
//
//						#endregion
//
//
//						#region 9. CREAR MAESTRO OC
//
//						oCmd.CommandText =
//							string.Format(
//							"Exec GrabaMaestroOrdenDeCompra " +
//							"0,2,{0},0,0,{1},{2},'{3}','{4}','{5}',0,0,5,'{6}',{7},0,{8},{9},{10},{11},1,0,0,0,0,0,0,{12}",
//							(int)this.txtIdComprobante.Value,
//							Convert.ToInt32(this.txtIdCliente.Value),
//							bodega,
//							numero,
//							DateTime.Now.ToString("yyyyMMdd"),
//							DateTime.Now.AddDays(3).ToString("yyyyMMdd"),
//							notas,
//							subTotal,
//							0,
//							subTotal,
//							iva,
//							total,
//							Convert.ToInt32(this.txtIdOrdenCompra.Value)
//							);
//
//						int idCompraNueva =
//							Convert.ToInt32(
//							oCmd.ExecuteScalar()
//							);
//
//						#endregion
//
//
//						#region 10. CREAR DETALLE
//
//						oCmd.CommandText =
//							string.Format(
//							"Exec GuardaDetalleOrdenDeCompra " +
//							"0,{0},{1},0,{2},0,{3},15,0,0,{4},19,0,{5},0,0,0,0,0,0,0,0,0,0",
//							idCompraNueva,
//							idArticulo,
//							cantidadTemporal,
//							costo,
//							subTotal,
//							bodega
//							);
//
//						oCmd.ExecuteNonQuery();
//
//						#endregion
//
//
//						#region 11. ACTUALIZAR GENERADO
//
//						int idOrdenDeCompraBodegaDetalle =
//							Convert.ToInt32(
//							drBodega["idOrdenDeCompraBodegaDetalle"]
//							);
//
//						oCmd.CommandText =
//							string.Format(
//							"UPDATE OrdenDeCompraBodegaDetalle " +
//							"SET " +
//							"CantidadGenerada = " +
//							"ISNULL(CantidadGenerada,0) + {0}, " +
//							"CantidadTemporal = 0 " +
//							"WHERE idOrdenDeCompraBodegaDetalle = {1}",
//							cantidadTemporal,
//							idOrdenDeCompraBodegaDetalle
//							);
//
//						oCmd.ExecuteNonQuery();
//
//						#endregion
//
//
//						#region 12. GUARDAR RESULTADO
//
//						ocGeneradas.Add(
//							new object[]
//					{
//						idCompraNueva,
//						bodega,
//						numero
//					}
//							);
//
//						#endregion
//					}
//				}
//
//				#endregion
//
//
//				#region 13. ACTUALIZAR ESTADO OC
//
//				oCmd.CommandText =
//					string.Format(
//					"IF EXISTS " +
//					"( " +
//					"	SELECT 1 " +
//					"	FROM OrdenDeCompraBodegaDetalle odbd " +
//					"		INNER JOIN DetOrdenDeCompra d " +
//					"			ON d.IdDetOrdenDeCompra = odbd.IdDetOrdenDeCompra " +
//					"	WHERE d.IdOrdenDeCompra = {0} " +
//					"	AND ISNULL(odbd.SaldoPendiente,0) > 0 " +
//					") " +
//					"BEGIN " +
//					"	UPDATE OrdenDeCompra " +
//					"	SET Estado = 5 " +
//					"	WHERE IdOrdenDeCompra = {0} " +
//					"END " +
//					"ELSE " +
//					"BEGIN " +
//					"	UPDATE OrdenDeCompra " +
//					"	SET Estado = 9 " +
//					"	WHERE IdOrdenDeCompra = {0} " +
//					"END",
//					Convert.ToInt32(
//					this.txtIdOrdenCompra.Value
//					)
//					);
//
//				oCmd.ExecuteNonQuery();
//
//				#endregion
//
//
//				oTransaction.Commit();
//
//
//				#region 14. RESULTADO
//
//				if (ocGeneradas.Count > 0)
//				{
//					MessageBox.Show(
//						"Órdenes generadas correctamente.",
//						"Point Technology",
//						MessageBoxButtons.OK,
//						MessageBoxIcon.Information
//						);
//
//					ConsultaMaestro(
//						Convert.ToInt32(
//						this.txtIdOrdenCompra.Value
//						)
//						);
//				}
//				else
//				{
//					MessageBox.Show(
//						"No existen cantidades pendientes para generar.",
//						"Point Technology",
//						MessageBoxButtons.OK,
//						MessageBoxIcon.Warning
//						);
//				}
//
//				#endregion
//			}
//			catch (Exception ex)
//			{
//				try
//				{
//					oTransaction.Rollback();
//				}
//				catch
//				{
//				}
//
//				MessageBox.Show(
//					ex.Message,
//					"Point Technology",
//					MessageBoxButtons.OK,
//					MessageBoxIcon.Error
//					);
//			}
//			finally
//			{
//				oConexion.Close();
//			}
//		}


		private void btnGenerar_Click(object sender, System.EventArgs e)
		{
			#region 1. VALIDAR SELECCIÓN

			ArrayList filasSeleccionadas =
				new ArrayList();

			for (int i = 0; i < this.ultraGrid1.Rows.Count; i++)
			{
				UltraGridRow r =
					this.ultraGrid1.Rows[i];

				if (r == null || r.IsAddRow)
					continue;

				bool sel =
					(r.Cells["Sel"].Value != null &&
					r.Cells["Sel"].Value != DBNull.Value &&
					Convert.ToBoolean(r.Cells["Sel"].Value));

				if (!sel)
					continue;

				filasSeleccionadas.Add(r);
			}

			if (filasSeleccionadas.Count == 0)
			{
				MessageBox.Show(
					"Debe seleccionar al menos un artículo.",
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Warning
					);

				return;
			}

			#endregion


			#region 2. SERVICE

			OrdenCompraBodegaGeneracionService service =
				new OrdenCompraBodegaGeneracionService(
				this.cdsSeteoF
				);

			#endregion


			#region 3. CONEXIÓN

			SqlConnection oConexion =
				new SqlConnection(
				"Password=Bl45o6$9;Persist Security Info=True;User ID=" +
				MenuLatinium.stUsuario +
				";Initial Catalog=" +
				MenuLatinium.stDirFacturacion +
				";Data Source=" +
				MenuLatinium.stDirServidor +
				"; max pool size=1024"
				);

			oConexion.Open();

			SqlTransaction oTransaction =
				oConexion.BeginTransaction();

			SqlCommand oCmd =
				oConexion.CreateCommand();

			oCmd.Connection = oConexion;
			oCmd.Transaction = oTransaction;
			oCmd.CommandType = CommandType.Text;
			oCmd.CommandTimeout = 120;

			#endregion


			try
			{
				ArrayList ocGeneradas =
					new ArrayList();

				// =====================================================
				// CACHE DE OC POR BODEGA
				// =====================================================

				Hashtable comprasPorBodega =
					new Hashtable();


				#region 4. RECORRER FILAS SELECCIONADAS

				foreach (UltraGridRow row in filasSeleccionadas)
				{
					int idDetOrdenDeCompra =
						Convert.ToInt32(
						row.Cells["IdDetOrdenDeCompra"].Value
						);

					int idArticulo =
						Convert.ToInt32(
						row.Cells["IdArticulo"].Value
						);

					decimal costo =
						Convert.ToDecimal(
						row.Cells["Costo"].Value
						);

					string notas = "";

					if (row.Cells["Notas"].Value != null)
					{
						notas =
							row.Cells["Notas"].Value.ToString();
					}

					notas =
						notas.Replace("'", "''");


					#region 5. CONSULTAR BODEGAS

					DataTable dtBodegas =
						service.Consultar(
						idDetOrdenDeCompra,
						oConexion,
						oTransaction
						);

					if (dtBodegas == null ||
						dtBodegas.Rows.Count == 0)
					{
						continue;
					}

					#endregion


					#region 6. RECORRER BODEGAS

					Hashtable procesados =
						new Hashtable();

					foreach (DataRow drBodega in dtBodegas.Rows)
					{
						int bodega =
							Convert.ToInt32(
							drBodega["Bodega"]
							);

						string key =
							idDetOrdenDeCompra.ToString() +
							"_" +
							bodega.ToString();

						if (procesados.ContainsKey(key))
							continue;

						procesados.Add(key, true);

						int cantidadTemporal = 0;

						if (drBodega["CantidadTemporal"] != DBNull.Value)
						{
							cantidadTemporal =
								Convert.ToInt32(
								drBodega["CantidadTemporal"]
								);
						}

						int saldoPendiente = 0;

						if (drBodega["SaldoPendiente"] != DBNull.Value)
						{
							saldoPendiente =
								Convert.ToInt32(
								drBodega["SaldoPendiente"]
								);
						}

						if (cantidadTemporal <= 0)
							continue;

						if (saldoPendiente <= 0)
							continue;

						// =====================================================
						// NO PERMITIR PASAR EL SALDO
						// =====================================================

						if (cantidadTemporal > saldoPendiente)
						{
							cantidadTemporal =
								saldoPendiente;
						}

						#endregion


						#region 7. CALCULAR TOTALES

						decimal subtotal =
							Math.Round(
							costo * cantidadTemporal,
							2
							);

						if (subtotal <= 0)
							continue;

						decimal porcentajeIVA =
							Funcion.decEscalarSQL(
							cdsSeteoF,
							string.Format(
							"Select dbo.RetornaImpuestoIVAVigente('{0}',1)",
							DateTime.Now.ToString("yyyyMMdd")
							)
							);

						decimal iva =
							Math.Round(
							(subtotal * porcentajeIVA) / 100m,
							2
							);

						decimal total =
							subtotal + iva;

						#endregion


						#region 8. OBTENER O CREAR OC

						int idCompraNueva = 0;

						if (comprasPorBodega.ContainsKey(bodega))
						{
							idCompraNueva =
								Convert.ToInt32(
								comprasPorBodega[bodega]
								);
						}
						else
						{
							#region GENERAR NÚMERO

							oCmd.CommandText =
								"Exec FacturaNumeroEstacion 2";

							string numero =
								Convert.ToString(
								oCmd.ExecuteScalar()
								);

							string numeroOC = "";

							if (this.txtNumero.Text != null)
							{
								numeroOC =
									this.txtNumero.Text.Trim();
							}

							if (numero == null)
							{
								numero = "";
							}

							numero =
								numeroOC + "-" + numero;

							#endregion


							#region CREAR MAESTRO CORRECTAMENTE

							// =====================================================
							// IMPORTANTE:
							// LA ODC DEBE CREARSE YA CON:
							// - SUBTOTAL
							// - IVA
							// - TOTAL
							// - FORMA PAGO
							// - FECHAS
							// - ESTADO
							// =====================================================

							oCmd.CommandText =
								string.Format(
								"Exec GrabaMaestroOrdenDeCompra " +

								// idCompra
								"0," +

								// ContadoCredito
								"2," +

								// idComprobante
								"{0}," +

								// idCredTributario
								"0," +

								// Consignacion
								"0," +

								// idCliente
								"{1}," +

								// Bodega
								"{2}," +

								// Numero
								"'{3}'," +

								// Fecha
								"'{4}'," +

								// FechaTentativaLlegada
								"'{5}'," +

								// HoraLlegada
								"'{6}'," +

								// idImportacionTipo
								"0," +

								// Estado
								"5," +

								// Notas
								"'{7}'," +

								// SubtotalExcento
								"0," +

								// Descuento0
								"0," +

								// SubtotalIva
								"{8}," +

								// Descuento
								"0," +

								// IVA
								"{9}," +

								// Total
								"{10}," +

								// bFechaLlegada
								"1," +

								// bHoraLlegada
								"1," +

								// idFormaPago
								"9," +

								// Creditos
								"0,0,0,0," +

								// idOrigen
								"{11}",

								(int)this.txtIdComprobante.Value,
								Convert.ToInt32(this.txtIdCliente.Value),
								bodega,
								numero,
								DateTime.Now.ToString("yyyyMMdd"),
								DateTime.Now.AddDays(3).ToString("yyyyMMdd"),
								DateTime.Now.ToString("yyyyMMdd HH:mm:ss"),
								notas,

								subtotal,

								iva,

								total,

								Convert.ToInt32(
								this.txtIdOrdenCompra.Value
								)
								);

							idCompraNueva =
								Convert.ToInt32(
								oCmd.ExecuteScalar()
								);

							comprasPorBodega.Add(
								bodega,
								idCompraNueva
								);

							ocGeneradas.Add(
								new object[]
					{
						idCompraNueva,
						bodega,
						numero
					}
								);

							#endregion
						}

						#endregion


						#region 9. CREAR DETALLE

						int idProyecto =
							(bodega == 1)
							? 19
							: 107;

						oCmd.CommandText =
							string.Format(
							"Exec GuardaDetalleOrdenDeCompra " +
							"0,{0},{1},0,{2},0,{3},{4},0,0,{5},{6},0,{7},0,0,0,0,0,0,0,0,0,0",
							idCompraNueva,
							idArticulo,
							cantidadTemporal,
							costo.ToString(
							System.Globalization.CultureInfo.InvariantCulture
							),
							porcentajeIVA.ToString(
							System.Globalization.CultureInfo.InvariantCulture
							),
							subtotal.ToString(
							System.Globalization.CultureInfo.InvariantCulture
							),
							idProyecto,
							bodega
							);

						oCmd.ExecuteNonQuery();

						#endregion


						#region 10. ACTUALIZAR GENERADO

						int idOrdenDeCompraBodegaDetalle =
							Convert.ToInt32(
							drBodega["idOrdenDeCompraBodegaDetalle"]
							);

						oCmd.CommandText =
							string.Format(
							"UPDATE OrdenDeCompraBodegaDetalle " +
							"SET " +
							"CantidadGenerada = " +
							"ISNULL(CantidadGenerada,0) + {0}, " +
							"CantidadTemporal = 0 " +
							"WHERE idOrdenDeCompraBodegaDetalle = {1}",
							cantidadTemporal,
							idOrdenDeCompraBodegaDetalle
							);

						oCmd.ExecuteNonQuery();

						#endregion
					}
				}

				#endregion


				#region 11. RECALCULAR CABECERA

				foreach (DictionaryEntry item in comprasPorBodega)
				{
					int idCompra =
						Convert.ToInt32(
						item.Value
						);

					decimal dPorcIva =
						Funcion.decEscalarSQL(
						cdsSeteoF,
						string.Format(
						"Select dbo.RetornaImpuestoIVAVigente('{0}',1)",
						DateTime.Now.ToString("yyyyMMdd")
						)
						);

					oCmd.CommandText =
						string.Format(
						"UPDATE Compra " +
						"SET " +

						// SUBTOTAL GENERAL
						"subTotal = ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0), " +

						// SUBTOTAL IVA
						"SubtotalIva = ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0), " +

						// SUBTOTAL EXENTO
						"SubtotalExcento = 0, " +

						// PORCENTAJE IVA
						"PorcentajeIva = {1}, " +

						// IVA
						"Iva = ROUND((ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0) * {1}) / 100.00,2), " +

						// TOTAL
						"Total = " +
						"ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0) + " +
						"ROUND((ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0) * {1}) / 100.00,2), " +

						// SALDO
						"Saldo = " +
						"ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0) + " +
						"ROUND((ISNULL((SELECT SUM(Subtotal) FROM DetCompra WHERE IdCompra = {0}),0) * {1}) / 100.00,2) " +

						"WHERE idCompra = {0}",

						idCompra,

						dPorcIva.ToString(
						System.Globalization.CultureInfo.InvariantCulture
						)
						);

					oCmd.ExecuteNonQuery();
				}

				#endregion


				#region 12. ACTUALIZAR ESTADO ODC

				oCmd.CommandText =
					string.Format(
					"IF EXISTS " +
					"( " +
					"	SELECT 1 " +
					"	FROM OrdenDeCompraBodegaDetalle odbd " +
					"		INNER JOIN DetOrdenDeCompra d " +
					"			ON d.IdDetOrdenDeCompra = odbd.IdDetOrdenDeCompra " +
					"	WHERE d.IdOrdenDeCompra = {0} " +
					"	AND ISNULL(odbd.SaldoPendiente,0) > 0 " +
					") " +
					"BEGIN " +
					"	UPDATE OrdenDeCompra " +
					"	SET Estado = 5 " +
					"	WHERE IdOrdenDeCompra = {0} " +
					"END " +
					"ELSE " +
					"BEGIN " +
					"	UPDATE OrdenDeCompra " +
					"	SET Estado = 9 " +
					"	WHERE IdOrdenDeCompra = {0} " +
					"END",
					Convert.ToInt32(
					this.txtIdOrdenCompra.Value
					)
					);

				oCmd.ExecuteNonQuery();

				#endregion


				oTransaction.Commit();


				#region 13. RESULTADO

				if (ocGeneradas.Count > 0)
				{
					MessageBox.Show(
						"Órdenes generadas correctamente.",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Information
						);

					ConsultaMaestro(
						Convert.ToInt32(
						this.txtIdOrdenCompra.Value
						)
						);
				}
				else
				{
					MessageBox.Show(
						"No existen cantidades pendientes para generar.",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning
						);
				}

				#endregion
			}
			catch (Exception ex)
			{
				try
				{
					oTransaction.Rollback();
				}
				catch
				{
				}

				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
			finally
			{
				oConexion.Close();
			}
		}

		private void ultraGrid1_AfterCellUpdate(
			object sender,
			Infragistics.Win.UltraWinGrid.CellEventArgs e
			)
		{
			try
			{
				// =========================================
				// SOLO CUANDO CAMBIA EL CHECK SEL
				// =========================================
				if (e.Cell.Column.Key != "Sel")
					return;

				// =========================================
				// VALIDAR ID DETALLE
				// =========================================
				if (
					e.Cell.Row.Cells["IdDetOrdenDeCompra"].Value == null
					||
					e.Cell.Row.Cells["IdDetOrdenDeCompra"].Value == DBNull.Value
					)
				{
					return;
				}

				int idDetOrdenDeCompra =
					Convert.ToInt32(
					e.Cell.Row.Cells["IdDetOrdenDeCompra"].Value
					);

				bool sel = false;

				if (
					e.Cell.Row.Cells["Sel"].Value != null
					&&
					e.Cell.Row.Cells["Sel"].Value != DBNull.Value
					)
				{
					sel =
						Convert.ToBoolean(
						e.Cell.Row.Cells["Sel"].Value
						);
				}

				// =========================================
				// SERVICE
				// =========================================
				OrdenCompraBodegaGeneracionService service =
					new OrdenCompraBodegaGeneracionService(
					cdsSeteoF
					);

				// =========================================
				// TABLA GRID BODEGA
				// =========================================
				DataTable dtBodega = null;

				if (
					this.ultraGridBodega.DataSource == null
					)
				{
					dtBodega = new DataTable();
				}
				else
				{
					dtBodega =
						(DataTable)this.ultraGridBodega.DataSource;
				}

				// =========================================
				// SI MARCA
				// =========================================
				if (sel)
				{
					// CONSULTAR BODEGAS DEL ARTÍCULO
					DataTable dtNuevo =
						service.Consultar(
						idDetOrdenDeCompra
						);

					// PRIMERA VEZ
					if (dtBodega.Columns.Count == 0)
					{
						dtBodega = dtNuevo.Clone();
					}

					// INSERTAR SIN DUPLICAR
					foreach (DataRow r in dtNuevo.Rows)
					{
						bool existe = false;

						foreach (DataRow r2 in dtBodega.Rows)
						{
							if (
								Convert.ToInt32(
								r2["idOrdenDeCompraBodegaDetalle"]
								)
								==
								Convert.ToInt32(
								r["idOrdenDeCompraBodegaDetalle"]
								)
								)
							{
								existe = true;
								break;
							}
						}

						if (!existe)
						{
							// VALORES DEFAULT
							if (r["CantidadTemporal"] == DBNull.Value)
								r["CantidadTemporal"] = 0;

							if (r["Sel"] == DBNull.Value)
								r["Sel"] = false;

							dtBodega.ImportRow(r);
						}
					}
				}
				else
				{
					// =========================================
					// SI DESMARCA
					// ELIMINAR DEL GRID BODEGA
					// =========================================

					ArrayList eliminar =
						new ArrayList();

					foreach (DataRow r in dtBodega.Rows)
					{
						if (
							Convert.ToInt32(
							r["IdDetOrdenDeCompra"]
							)
							==
							idDetOrdenDeCompra
							)
						{
							eliminar.Add(r);
						}
					}

					foreach (DataRow r in eliminar)
					{
						dtBodega.Rows.Remove(r);
					}
				}

				// =========================================
				// REFRESH GRID BODEGA
				// =========================================

				this.ultraGridBodega.DataSource =
					dtBodega;

				this.ultraGridBodega.DataBind();

				this.ultraGridBodega.Refresh();
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					"Error al cargar bodegas: " + ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}
		private void btnSalir_Click(object sender, System.EventArgs e)
		{		
			this.Close();
		}
		private void AbrirTurnoMercaderia()
		{
			frmTurnosMercaderia frm = new frmTurnosMercaderia();
			frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;

			frm.ShowDialog(this);
		}

		private void ultraGrid1_ClickCellButton(
			object sender,
			Infragistics.Win.UltraWinGrid.CellEventArgs e
			)
		{
			try
			{
				if (e.Cell == null)
					return;

				if (e.Cell.Column.Key == "btnBodega")
				{
					UltraGridRow row =
						e.Cell.Row;

					AbrirBodegas(row);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}

		private void AbrirBodegas(
			UltraGridRow row
			)
		{
			try
			{
				if (row == null)
					return;

				// VALIDAR ID DETALLE
				if (
					row.Cells["IdDetOrdenDeCompra"].Value == DBNull.Value
					)
				{
					MessageBox.Show(
						"No existe detalle de orden.",
						"Point Technology",
						MessageBoxButtons.OK,
						MessageBoxIcon.Warning
						);

					return;
				}

				int idDetOrdenDeCompra =
					Convert.ToInt32(
					row.Cells["IdDetOrdenDeCompra"].Value
					);

				// NUEVA PANTALLA
				frmOrdenDeCompraBodegaGenerar frm =
					new frmOrdenDeCompraBodegaGenerar(
					idDetOrdenDeCompra
					);

				frm.cdsSeteoF =
					this.cdsSeteoF;

				frm.ShowDialog();

				// RECARGAR GRID AL CERRAR
				ConsultaMaestro(
						Convert.ToInt32(this.txtIdOrdenCompra.Value)
				);
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void ConfigurarGrid()
		{
			if (this.ultraGrid1.DisplayLayout.Bands.Count == 0)
				return;

			UltraGridBand band = this.ultraGrid1.DisplayLayout.Bands[0];

			if (band.Columns.Exists("Sel"))
				band.Columns["Sel"].Header.Caption = "Sel";

			if (band.Columns.Exists("CantidadTemporal"))
				band.Columns["CantidadTemporal"].Header.Caption = "Generar";

			this.ultraGrid1.DisplayLayout.Override.AllowUpdate =
				Infragistics.Win.DefaultableBoolean.True;

			this.ultraGrid1.DisplayLayout.Override.CellClickAction =
				CellClickAction.Edit;
		}

		private void ultraGrid1_CellChange(
			object sender,
			Infragistics.Win.UltraWinGrid.CellEventArgs e
			)
		{
			try
			{
				// =========================
				// SOLO CHECK SEL
				// =========================
				if (e.Cell.Column.Key != "Sel")
					return;

				// =========================
				// GUARDAR CAMBIO INMEDIATO
				// =========================
				this.ultraGrid1.PerformAction(
					UltraGridAction.ExitEditMode
					);

				this.ultraGrid1.UpdateData();

				// =========================
				// VALIDAR ID
				// =========================
				if (
					e.Cell.Row.Cells["IdDetOrdenDeCompra"].Value == null
					||
					e.Cell.Row.Cells["IdDetOrdenDeCompra"].Value == DBNull.Value
					)
				{
					return;
				}

				int idDetOrdenDeCompra =
					Convert.ToInt32(
					e.Cell.Row.Cells["IdDetOrdenDeCompra"].Value
					);

				// =========================
				// LEER CHECK
				// =========================
				bool sel = false;

				try
				{
					sel =
						Convert.ToBoolean(
						e.Cell.Row.Cells["Sel"].Value
						);
				}
				catch
				{
					sel = false;
				}

				// =========================
				// PROCESAR GRID BODEGA
				// =========================
				CargarBodegaDetalle(
					idDetOrdenDeCompra,
					sel
					);
			}
			catch (Exception ex)
			{
				MessageBox.Show(
					ex.Message,
					"Point Technology",
					MessageBoxButtons.OK,
					MessageBoxIcon.Error
					);
			}
		}


		private void CargarBodegaDetalle(
			int idDetOrdenDeCompra,
			bool sel
			)
		{
			// =========================
			// SERVICE
			// =========================
			OrdenCompraBodegaGeneracionService service =
				new OrdenCompraBodegaGeneracionService(
				cdsSeteoF
				);

			// =========================
			// TABLA ACTUAL
			// =========================
			DataTable dtBodega;

			if (
				this.ultraGridBodega.DataSource == null
				)
			{
				dtBodega = new DataTable();
			}
			else
			{
				dtBodega =
					(DataTable)this.ultraGridBodega.DataSource;
			}

			// =========================
			// MARCAR
			// =========================
			if (sel)
			{
				DataTable dtNuevo =
					service.Consultar(
					idDetOrdenDeCompra
					);

				// CREAR ESTRUCTURA
				if (dtBodega.Columns.Count == 0)
				{
					dtBodega = dtNuevo.Clone();
				}

				// INSERTAR
				foreach (DataRow r in dtNuevo.Rows)
				{
					bool existe = false;

					foreach (DataRow r2 in dtBodega.Rows)
					{
						if (
							Convert.ToInt32(
							r2["idOrdenDeCompraBodegaDetalle"]
							)
							==
							Convert.ToInt32(
							r["idOrdenDeCompraBodegaDetalle"]
							)
							)
						{
							existe = true;
							break;
						}
					}

					if (!existe)
					{
						dtBodega.ImportRow(r);
					}
				}
			}
			else
			{
				// =========================
				// DESMARCAR
				// =========================
				ArrayList eliminar =
					new ArrayList();

				foreach (DataRow r in dtBodega.Rows)
				{
					if (
						Convert.ToInt32(
						r["IdDetOrdenDeCompra"]
						)
						==
						idDetOrdenDeCompra
						)
					{
						eliminar.Add(r);
					}
				}

				foreach (DataRow r in eliminar)
				{
					dtBodega.Rows.Remove(r);
				}
			}

			// =========================
			// REFRESH GRID
			// =========================
			this.ultraGridBodega.DataSource =
				dtBodega;

			this.ultraGridBodega.DataBind();

			this.ultraGridBodega.Refresh();

			// =========================
			// OCULTAR COLUMNAS
			// =========================
			OcultarColumnasBodega();
		}


		private void OcultarColumnasBodega()
		{
			if (
				this.ultraGridBodega.DisplayLayout.Bands.Count == 0
				)
			{
				return;
			}

			UltraGridBand band =
				this.ultraGridBodega.DisplayLayout.Bands[0];

			// =========================
			// GRID SOLO VISUAL
			// =========================
			this.ultraGridBodega.DisplayLayout.Override.AllowUpdate =
				Infragistics.Win.DefaultableBoolean.False;

			this.ultraGridBodega.DisplayLayout.Override.CellClickAction =
				CellClickAction.RowSelect;

			// =========================
			// OCULTAR COLUMNAS
			// =========================
			if (band.Columns.Exists("Bodega"))
				band.Columns["Bodega"].Hidden = true;

			if (band.Columns.Exists("idOrdenDeCompraBodegaDetalle"))
				band.Columns["idOrdenDeCompraBodegaDetalle"].Hidden = true;

			if (band.Columns.Exists("IdDetOrdenDeCompra"))
				band.Columns["IdDetOrdenDeCompra"].Hidden = true;

			if (band.Columns.Exists("IdOrdenDeCompra"))
				band.Columns["IdOrdenDeCompra"].Hidden = true;

			if (band.Columns.Exists("EstadoProceso"))
				band.Columns["EstadoProceso"].Hidden = true;

			if (band.Columns.Exists("IdArticulo"))
				band.Columns["IdArticulo"].Hidden = true;

			if (band.Columns.Exists("Sel"))
				band.Columns["Sel"].Hidden = true;

			if (band.Columns.Exists("Codigo"))
				band.Columns["Codigo"].Hidden = true;

			if (band.Columns.Exists("Articulo"))
				band.Columns["Articulo"].Hidden = true;

			// =========================
			// HEADERS
			// =========================
			if (band.Columns.Exists("NombreBodega"))
				band.Columns["NombreBodega"].Header.Caption = "Bodega";

//			if (band.Columns.Exists("Articulo"))
//				band.Columns["Articulo"].Header.Caption = "Artículo";

			if (band.Columns.Exists("CantidadAsignada"))
				band.Columns["CantidadAsignada"].Header.Caption = "Asignado";

			if (band.Columns.Exists("CantidadGenerada"))
				band.Columns["CantidadGenerada"].Header.Caption = "Generado";

			if (band.Columns.Exists("SaldoPendiente"))
				band.Columns["SaldoPendiente"].Header.Caption = "Pendiente";

			if (band.Columns.Exists("CantidadTemporal"))
				band.Columns["CantidadTemporal"].Header.Caption = "Generar";

			// =========================
			// TAMAÑOS
			// =========================
			if (band.Columns.Exists("NombreBodega"))
				band.Columns["NombreBodega"].Width = 120;

//			if (band.Columns.Exists("Articulo"))
//				band.Columns["Articulo"].Width = 250;

			if (band.Columns.Exists("CantidadAsignada"))
				band.Columns["CantidadAsignada"].Width = 70;

			if (band.Columns.Exists("CantidadGenerada"))
				band.Columns["CantidadGenerada"].Width = 70;

			if (band.Columns.Exists("SaldoPendiente"))
				band.Columns["SaldoPendiente"].Width = 70;

			if (band.Columns.Exists("CantidadTemporal"))
				band.Columns["CantidadTemporal"].Width = 70;
		}
	}
}