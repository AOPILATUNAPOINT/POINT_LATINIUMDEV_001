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
	/// Descripción breve de frmFAC_Descuentos.
	/// </summary>
	public class frmFAC_Descuentos : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFiltro;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenEfectivo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtMargenSinInteres;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmFAC_Descuentos()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmFAC_Descuentos));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFAC_Descuentos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Inicio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fin");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RUC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proveedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proteccion");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo_Proteccion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioTarjeta");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioTarjetaIVA");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioEfectivo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioTarjetaSI");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenEfectivo");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioMinEfectivo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenTarjeta");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioMinTarjeta");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MargenTarjetaSSI");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioMinTarjetaSI");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFAC_Descuentos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Inicio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fin");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("RUC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proveedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Proteccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Costo_Proteccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioTarjetaIVA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioTarjetaSI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioMinEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioMinTarjeta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MargenTarjetaSSI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrecioMinTarjetaSI");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.optFiltro = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.txtMargenTarjeta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtMargenEfectivo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtMargenSinInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenSinInteres)).BeginInit();
			this.SuspendLayout();
			// 
			// btnExcel
			// 
			this.btnExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(1112, 40);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(75, 24);
			this.btnExcel.TabIndex = 365;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(280, 7);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 364;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label2.Location = new System.Drawing.Point(8, 11);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 17);
			this.label2.TabIndex = 363;
			this.label2.Text = "Buscar";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optFiltro
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFiltro.Appearance = appearance1;
			this.optFiltro.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFiltro.CheckedIndex = 0;
			this.optFiltro.ItemAppearance = appearance2;
			this.optFiltro.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Activos";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Finalizados";
			this.optFiltro.Items.Add(valueListItem1);
			this.optFiltro.Items.Add(valueListItem2);
			this.optFiltro.Items.Add(valueListItem3);
			this.optFiltro.ItemSpacingVertical = 10;
			this.optFiltro.Location = new System.Drawing.Point(368, 7);
			this.optFiltro.Name = "optFiltro";
			this.optFiltro.Size = new System.Drawing.Size(224, 24);
			this.optFiltro.TabIndex = 362;
			this.optFiltro.Text = "Todos";
			this.optFiltro.ValueChanged += new System.EventHandler(this.optFiltro_ValueChanged);
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
			// txtArticulo
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance3;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(64, 8);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(200, 22);
			this.txtArticulo.TabIndex = 360;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance4.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance4;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Format = "dd/MM/yyyy";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 80;
			ultraGridColumn3.Format = "dd/MM/yyyy";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 150;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "R.U.C.";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 120;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 150;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn9.CellAppearance = appearance5;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 90;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 150;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 14;
			ultraGridColumn11.Width = 80;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 15;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "Costo - Proteccion";
			ultraGridColumn13.Header.VisiblePosition = 16;
			ultraGridColumn13.Width = 130;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Precio Normal";
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Width = 130;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance10;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "Precio Normal I.V.A.";
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Width = 138;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance11;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Precio Efectivo";
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Width = 130;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance12;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "Precio Tarjeta S.I.";
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Width = 130;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance13;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Margen Efectivo";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 130;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance14;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.Caption = "Precio Min Efectivo";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 135;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance15;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "Margen Tarjeta";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 130;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance16;
			ultraGridColumn21.Format = "#,##0.00";
			ultraGridColumn21.Header.Caption = "Precio Min Tarjeta";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 130;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance17;
			ultraGridColumn22.Format = "#,##0.00";
			ultraGridColumn22.Header.Caption = "Margen Tarjeta S.I.";
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 130;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance18;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.Caption = "Precio Min Tarjeta S.I.";
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 130;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance19;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 90;
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
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance20;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance21.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance21;
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance23.FontData.BoldAsString = "True";
			appearance23.FontData.Name = "Arial";
			appearance23.FontData.SizeInPoints = 8F;
			appearance23.ForeColor = System.Drawing.Color.White;
			appearance23.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance23;
			appearance24.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance24.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance25;
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance26.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance26;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1176, 376);
			this.ultraGrid1.TabIndex = 366;
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(System.DateTime);
			ultraDataColumn3.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			ultraDataColumn11.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn12.DataType = typeof(System.Decimal);
			ultraDataColumn12.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn13.DataType = typeof(System.Decimal);
			ultraDataColumn13.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn14.DataType = typeof(System.Decimal);
			ultraDataColumn14.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn15.DataType = typeof(System.Decimal);
			ultraDataColumn15.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn18.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn19.DataType = typeof(System.Decimal);
			ultraDataColumn19.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn20.DataType = typeof(System.Decimal);
			ultraDataColumn20.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn21.DataType = typeof(System.Decimal);
			ultraDataColumn21.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn22.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn23.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(1111, 8);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 367;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 72);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1192, 8);
			this.groupBox1.TabIndex = 368;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.BackColor = System.Drawing.Color.Transparent;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(624, 11);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 369;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(920, 16);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(40, 17);
			this.label13.TabIndex = 373;
			this.label13.Text = "Tarjeta";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(840, 16);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(43, 17);
			this.label14.TabIndex = 372;
			this.label14.Text = "Efectivo";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenTarjeta
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenTarjeta.Appearance = appearance27;
			this.txtMargenTarjeta.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMargenTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenTarjeta.FormatString = "#,##0.00";
			this.txtMargenTarjeta.Location = new System.Drawing.Point(920, 40);
			this.txtMargenTarjeta.Name = "txtMargenTarjeta";
			this.txtMargenTarjeta.NullText = "0.00";
			this.txtMargenTarjeta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenTarjeta.PromptChar = ' ';
			this.txtMargenTarjeta.Size = new System.Drawing.Size(72, 20);
			this.txtMargenTarjeta.TabIndex = 371;
			// 
			// txtMargenEfectivo
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenEfectivo.Appearance = appearance28;
			this.txtMargenEfectivo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMargenEfectivo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenEfectivo.FormatString = "#,##0.00";
			this.txtMargenEfectivo.Location = new System.Drawing.Point(840, 40);
			this.txtMargenEfectivo.Name = "txtMargenEfectivo";
			this.txtMargenEfectivo.NullText = "0.00";
			this.txtMargenEfectivo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenEfectivo.PromptChar = ' ';
			this.txtMargenEfectivo.Size = new System.Drawing.Size(72, 20);
			this.txtMargenEfectivo.TabIndex = 370;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(1000, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(58, 17);
			this.label1.TabIndex = 375;
			this.label1.Text = "Sin Interés";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtMargenSinInteres
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMargenSinInteres.Appearance = appearance29;
			this.txtMargenSinInteres.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtMargenSinInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMargenSinInteres.FormatString = "#,##0.00";
			this.txtMargenSinInteres.Location = new System.Drawing.Point(1000, 40);
			this.txtMargenSinInteres.Name = "txtMargenSinInteres";
			this.txtMargenSinInteres.NullText = "0.00";
			this.txtMargenSinInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtMargenSinInteres.PromptChar = ' ';
			this.txtMargenSinInteres.Size = new System.Drawing.Size(72, 20);
			this.txtMargenSinInteres.TabIndex = 374;
			// 
			// frmFAC_Descuentos
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1192, 466);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtMargenSinInteres);
			this.Controls.Add(this.txtMargenTarjeta);
			this.Controls.Add(this.txtMargenEfectivo);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.optFiltro);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.btnExcel);
			this.Controls.Add(this.btnVer);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmFAC_Descuentos";
			this.Text = "Descuentos Facturación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmFAC_Descuentos_KeyDown);
			this.Load += new System.EventHandler(this.frmFAC_Descuentos_Load);
			((System.ComponentModel.ISupportInitialize)(this.optFiltro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenEfectivo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMargenSinInteres)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private bool bActualiza = true;

		private void frmFAC_Descuentos_Load(object sender, System.EventArgs e)
		{
			SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, "Select MargenEfectivo, MargenTarjeta, MargenTarjetaSI From Fac_Descuentos_Margen", true);
			drArticulo.Read();
			if (drArticulo.HasRows)
			{
				this.txtMargenEfectivo.Value = drArticulo.GetDecimal(0);
				this.txtMargenTarjeta.Value = drArticulo.GetDecimal(1);
				this.txtMargenSinInteres.Value = drArticulo.GetDecimal(2);
			}
			drArticulo.Close();

			Consulta();
		}

		private void Consulta()
		{
			string sSQL = string.Format("Exec FAC_DescuentoConsulta '{0}', {1}", this.txtArticulo.Text.ToString(), (int)this.optFiltro.Value);
			FuncionesProcedimientos.dsGeneralModulo(sSQL, this.ultraGrid1);

			this.lblContador.Text = this.ultraGrid1.Rows.Count.ToString() + " REGISTROS ENCONTRADOS";
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "RUC")
			{				
				#region Busca Proveedor
				if (e.Cell.Row.Cells["RUC"].Value.ToString().Length > 0)
				{
					string sSQL = string.Format("Exec BuscaClientePorRUC '{0}', {1}, {2}, {3}", e.Cell.Row.Cells["RUC"].Value.ToString(), 4, 0, false);
					SqlDataReader drC = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			
					drC.Read();
					if (drC.HasRows)
					{
						e.Cell.Row.Cells["idCliente"].Value = (int) drC.GetValue(0);
						e.Cell.Row.Cells["Proveedor"].Value = drC.GetString(2);	
			
						drC.Close();						
					}
					else
					{
						drC.Close();
						MessageBox.Show("Proveedor no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						e.Cell.Row.Cells["RUC"].Value = "";
						return;
					}
					drC.Close();
				}
				#endregion Busca Proveedor
			}
			
			if (!bActualiza) return;
						
			#region Codigo Articulo
			if (e.Cell.Column.ToString() == "Codigo")
			{
				bActualiza = false;

				string sCodigo = "";
				string sSQL = string.Format("Exec BuscaArticulosDetalleCombos '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
				int iCuenta = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				if (iCuenta == 1)
				{
					string sSQLArticulo = string.Format("Exec FAC_DescuentosArticulo '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());
					SqlDataReader drArticulo = Funcion.rEscalarSQL(cdsSeteoF, sSQLArticulo, true);
					drArticulo.Read();
					if (drArticulo.HasRows)
					{
						e.Cell.Row.Cells["idArticulo"].Value = drArticulo.GetInt32(0);
						sCodigo = drArticulo.GetString(1);
						e.Cell.Row.Cells["Articulo"].Value = drArticulo.GetString(2);
						e.Cell.Row.Cells["Costo"].Value = drArticulo.GetDecimal(3);
						e.Cell.Row.Cells["PrecioTarjeta"].Value = drArticulo.GetDecimal(4);
						e.Cell.Row.Cells["PrecioTarjetaIVA"].Value = drArticulo.GetDecimal(5);
						e.Cell.Row.Cells["PrecioEfectivo"].Value = drArticulo.GetDecimal(6);
						e.Cell.Row.Cells["PrecioTarjetaSI"].Value = drArticulo.GetDecimal(7);						
						
						drArticulo.Close();

						e.Cell.Row.Cells["Codigo"].Value = sCodigo;
					}
					drArticulo.Close();
				}
				else if (iCuenta != 1)
				{
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODIGO", 1, 1, 1, DateTime.Today);					
				}
			
				bActualiza = true;
			}
			#endregion Codigo Articulo

			if (e.Cell.Column.ToString() == "Proteccion")
			{	
				e.Cell.Row.Cells["Costo_Proteccion"].Value = (Convert.ToDecimal(e.Cell.Row.Cells["Costo"].Value) - Convert.ToDecimal(e.Cell.Row.Cells["Proteccion"].Value)) * 1.15m;

				e.Cell.Row.Cells["PrecioMinEfectivo"].Value = 0.00m;
				e.Cell.Row.Cells["PrecioMinTarjeta"].Value = 0.00m;
				e.Cell.Row.Cells["PrecioMinTarjetaSI"].Value = 0.00m;

				if (Convert.ToDecimal(e.Cell.Row.Cells["Costo_Proteccion"].Value) > 0.00m)
				{
					e.Cell.Row.Cells["PrecioMinEfectivo"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Costo_Proteccion"].Value)/Convert.ToDecimal(e.Cell.Row.Cells["MargenEfectivo"].Value);
					e.Cell.Row.Cells["PrecioMinTarjeta"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Costo_Proteccion"].Value)/Convert.ToDecimal(e.Cell.Row.Cells["MargenTarjeta"].Value);
					e.Cell.Row.Cells["PrecioMinTarjetaSI"].Value = Convert.ToDecimal(e.Cell.Row.Cells["Costo_Proteccion"].Value)/Convert.ToDecimal(e.Cell.Row.Cells["MargenTarjetaSSI"].Value);
				}				
			}
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{					
				if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "RUC")
				{
					using (BuscaClientes miBusqueda = new BuscaClientes(this.ultraGrid1.ActiveRow.Cells["RUC"].Value.ToString(), 4, 0, 1, 0, false))
						if (DialogResult.OK == miBusqueda.ShowDialog())
							this.ultraGrid1.ActiveRow.Cells["RUC"].Value = miBusqueda.Ruc;
				}							
																
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
					FuncionesInventario.AbreBusquedaArticulosCombos(this.ultraGrid1, "CODIGO", 1, 1, 1, DateTime.Today);			
			}

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idFAC_Descuentos"].Value = 0;
			e.Row.Cells["Inicio"].Value = DateTime.Today;
			e.Row.Cells["Fin"].Value = DateTime.Today;
			e.Row.Cells["idCliente"].Value = 0;						
			e.Row.Cells["idArticulo"].Value = 0;
			e.Row.Cells["Costo"].Value = 0.00m;
			e.Row.Cells["Costo_Proteccion"].Value = 0.00m;
			e.Row.Cells["MargenEfectivo"].Value = Convert.ToDecimal(this.txtMargenEfectivo.Value);
			e.Row.Cells["MargenTarjeta"].Value = Convert.ToDecimal(this.txtMargenTarjeta.Value);
			e.Row.Cells["MargenTarjetaSSI"].Value = Convert.ToDecimal(this.txtMargenSinInteres.Value);
			e.Row.Cells["PrecioMinEfectivo"].Value = 0.00m;
			e.Row.Cells["PrecioMinTarjeta"].Value = 0.00m;
			e.Row.Cells["PrecioMinTarjetaSI"].Value = 0.00m;
			e.Row.Cells["Proteccion"].Value = 0.00m;
			e.Row.Cells["PrecioTarjeta"].Value = 0.00m;
			e.Row.Cells["PrecioTarjetaIVA"].Value = 0.00m;
			e.Row.Cells["PrecioEfectivo"].Value = 0.00m;
			e.Row.Cells["PrecioTarjetaSI"].Value = 0.00m;	
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Proteccion", 2);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{			
			#region Graba Registro
			int idFAC_Descuentos = 0;
			int idCliente = 0;
			int idArticulo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				idFAC_Descuentos = 0;
				idCliente = 0;
				idArticulo = 0;

				if (dr.Cells["idFAC_Descuentos"].Value != System.DBNull.Value) idFAC_Descuentos = (int) dr.Cells["idFAC_Descuentos"].Value;
				if (dr.Cells["idCliente"].Value != System.DBNull.Value) idCliente = (int) dr.Cells["idCliente"].Value;
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int) dr.Cells["idArticulo"].Value;

				string sSQL = string.Format("Exec FAC_DescuentoGuarda {0}, '{1}', '{2}', '{3}', {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}", 
					idFAC_Descuentos, 
					Convert.ToDateTime(dr.Cells["Inicio"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(dr.Cells["Fin"].Value).ToString("yyyyMMdd"), dr.Cells["Descripcion"].Value.ToString(),
					idCliente, idArticulo,
					Convert.ToDecimal(dr.Cells["Costo"].Value), Convert.ToDecimal(dr.Cells["Proteccion"].Value), Convert.ToDecimal(dr.Cells["Costo_Proteccion"].Value), 					
					Convert.ToDecimal(dr.Cells["PrecioTarjeta"].Value), Convert.ToDecimal(dr.Cells["PrecioTarjetaIVA"].Value), 
					Convert.ToDecimal(dr.Cells["PrecioEfectivo"].Value), Convert.ToDecimal(dr.Cells["PrecioTarjetaSI"].Value),
					Convert.ToDecimal(dr.Cells["MargenEfectivo"].Value), Convert.ToDecimal(dr.Cells["PrecioMinEfectivo"].Value),
					Convert.ToDecimal(dr.Cells["MargenTarjeta"].Value), Convert.ToDecimal(dr.Cells["PrecioMinTarjeta"].Value),
					Convert.ToDecimal(dr.Cells["MargenTarjetaSSI"].Value), Convert.ToDecimal(dr.Cells["PrecioMinTarjetaSI"].Value));
				Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);				
			}
			#endregion Graba Registro

			MessageBox.Show("Proceso Generado Correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Consulta();
		}

		private void optFiltro_ValueChanged(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void frmFAC_Descuentos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();

			if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);
		}
	}
}